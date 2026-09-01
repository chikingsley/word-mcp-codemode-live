"""Transactional, verified batches for live Microsoft Word editing."""

import inspect
import json
import logging
import pkgutil
from collections.abc import Awaitable, Callable
from contextlib import contextmanager
from importlib import import_module
from time import perf_counter
from typing import Any, get_args, get_type_hints

from fastmcp.tools import ToolResult
from fastmcp.utilities.types import Image
from mcp.types import TextContent
from pydantic import validate_call

from word_mcp_codemode_live import tools as tools_package
from word_mcp_codemode_live.core.word_com import undo_named_record, unique_undo_name
from word_mcp_codemode_live.tools.capture import render_word_pages
from word_mcp_codemode_live.tools.metadata import word_tool

logger = logging.getLogger(__name__)

BatchTool = Callable[..., Awaitable[str]]
ValidatedBatchTool = Callable[..., Awaitable[Any]]


class BatchOperationError(RuntimeError):
    """Raised when a nested live tool reports an error."""


class BatchPreparationError(ValueError):
    """Raised when one requested operation is invalid before mutation."""

    def __init__(self, message: str, index: int) -> None:
        super().__init__(message)
        self.index = index


class BatchExecutionError(RuntimeError):
    """Preserve rollback information when a prepared operation fails."""

    def __init__(self, error: Exception, index: int | None, mutation_attempted: bool) -> None:
        super().__init__(str(error))
        self.index = index
        self.mutation_attempted = mutation_attempted


def _batch_tools() -> dict[str, BatchTool]:
    """Discover callables carrying co-located ``batchable`` tool metadata."""
    discovered: dict[str, BatchTool] = {}
    prefix = f"{tools_package.__name__}."
    for module_info in pkgutil.walk_packages(tools_package.__path__, prefix):
        module = import_module(module_info.name)
        for value in vars(module).values():
            metadata = getattr(value, "__fastmcp__", None)
            if metadata is None or "batchable" not in (metadata.tags or set()):
                continue
            name = metadata.name or value.__name__
            discovered[name] = value
    return discovered


def _document_stats(document: Any) -> dict[str, Any]:
    return {
        "pages": int(document.ComputeStatistics(2)),  # wdStatisticPages
        "words": int(document.ComputeStatistics(0)),  # wdStatisticWords
        "paragraphs": int(document.Paragraphs.Count),
        "tables": int(document.Tables.Count),
        "sections": int(document.Sections.Count),
        "saved": bool(document.Saved),
    }


def _range_page(word_range: Any) -> int | None:
    try:
        return int(word_range.Information(3))  # wdActiveEndPageNumber
    except Exception:
        return None


def _find_page(document: Any, text: str) -> int | None:
    if not text or len(text) > 255:
        return None
    word_range = document.Content.Duplicate
    word_range.Find.ClearFormatting()
    found = word_range.Find.Execute(
        FindText=text,
        Forward=True,
        MatchCase=False,
        MatchWholeWord=False,
        Wrap=0,
    )
    return _range_page(word_range) if found else None


def _operation_pages(document: Any, operation: dict[str, Any]) -> set[int]:
    """Infer pages touched by an operation without requiring every tool to change its API."""
    arguments = operation.get("arguments") or {}
    pages: set[int] = set()

    for key in ("find_text", "target_text"):
        value = arguments.get(key)
        if isinstance(value, str):
            page = _find_page(document, value)
            if page:
                pages.add(page)

    for key in ("start_paragraph", "end_paragraph", "paragraph_index"):
        value = arguments.get(key)
        if isinstance(value, int) and 1 <= value <= document.Paragraphs.Count:
            page = _range_page(document.Paragraphs(value).Range)
            if page:
                pages.add(page)

    target_index = arguments.get("target_paragraph_index")
    if isinstance(target_index, int) and 0 <= target_index < document.Paragraphs.Count:
        page = _range_page(document.Paragraphs(target_index + 1).Range)
        if page:
            pages.add(page)

    table_index = arguments.get("table_index")
    if isinstance(table_index, int) and document.Tables.Count:
        normalized = document.Tables.Count if table_index == -1 else table_index
        if 1 <= normalized <= document.Tables.Count:
            page = _range_page(document.Tables(normalized).Range)
            if page:
                pages.add(page)

    position = arguments.get("position")
    if position == "start":
        pages.add(1)
    elif position == "end":
        pages.add(int(document.ComputeStatistics(2)))
    return pages


def _parse_nested_result(tool_name: str, result: Any) -> Any:
    if isinstance(result, str):
        try:
            parsed = json.loads(result)
        except json.JSONDecodeError:
            parsed = result
    else:
        parsed = result
    if isinstance(parsed, dict):
        if parsed.get("error"):
            raise BatchOperationError(f"{tool_name}: {parsed['error']}")
        if parsed.get("success") is False or parsed.get("ok") is False:
            raise BatchOperationError(f"{tool_name}: operation reported failure: {parsed}")
    return parsed


def _is_boolean_annotation(annotation: Any) -> bool:
    if annotation is bool:
        return True
    arguments = set(get_args(annotation))
    return bool in arguments and arguments <= {bool, type(None)}


def _error_result(message: str, *, rolled_back: bool, operation_index: int | None) -> ToolResult:
    metadata = {
        "success": False,
        "error": message,
        "rolled_back": rolled_back,
        "failed_operation": operation_index,
    }
    return ToolResult(
        content=[TextContent(type="text", text=json.dumps(metadata, ensure_ascii=False))],
        structured_content=metadata,
        is_error=True,
    )


def _undo_batch(document: Any, app: Any, undo_name: str, mutation_attempted: bool) -> bool:
    """Undo this batch without accidentally undoing an older user action."""
    if not mutation_attempted:
        return False

    return undo_named_record(document, app, undo_name)


def _prepare_operation(
    operation: dict[str, Any],
    index: int,
    batch_tools: dict[str, BatchTool],
    filename: str | None,
) -> tuple[str, ValidatedBatchTool, dict[str, Any]]:
    tool_name = operation.get("tool")
    arguments = operation.get("arguments", {})
    if not isinstance(tool_name, str) or tool_name not in batch_tools:
        raise BatchPreparationError(f"Unsupported batch tool: {tool_name!r}", index)
    if not isinstance(arguments, dict):
        raise BatchPreparationError(f"Operation {index} arguments must be an object", index)
    if "filename" in arguments:
        raise BatchPreparationError(
            f"Operation {index} must not include filename; use the batch-level filename", index
        )
    call_arguments = dict(arguments)
    tool = batch_tools[tool_name]
    signature = inspect.signature(tool)
    if "filename" in signature.parameters:
        call_arguments["filename"] = filename
    try:
        signature.bind(**call_arguments)
    except TypeError as exc:
        raise BatchPreparationError(str(exc), index) from exc
    annotations = get_type_hints(tool)
    invalid_booleans = [
        name
        for name, value in call_arguments.items()
        if name in annotations
        and _is_boolean_annotation(annotations[name])
        and value is not None
        and type(value) is not bool
    ]
    if invalid_booleans:
        raise BatchPreparationError(
            f"Operation {index} requires JSON boolean values for: {invalid_booleans}", index
        )
    return tool_name, validate_call(tool), call_arguments


def _prepare_operations(
    operations: list[dict[str, Any]], filename: str | None
) -> list[tuple[str, ValidatedBatchTool, dict[str, Any]]]:
    tools = _batch_tools()
    return [
        _prepare_operation(operation, index, tools, filename)
        for index, operation in enumerate(operations)
    ]


@contextmanager
def _batch_ui_state(app: Any):
    old_screen_updating = bool(app.ScreenUpdating)
    try:
        old_pagination = bool(app.Options.Pagination)
    except Exception:
        old_pagination = True
    app.ScreenUpdating = False
    try:
        try:
            app.Options.Pagination = False
        except Exception as exc:
            logger.debug("Word pagination optimization is unavailable: %s", exc)
        yield
    finally:
        try:
            app.Options.Pagination = old_pagination
        except Exception as exc:
            logger.debug("Could not restore Word pagination option: %s", exc)
        app.ScreenUpdating = old_screen_updating


async def _execute_prepared(
    app: Any,
    undo_name: str,
    prepared: list[tuple[str, ValidatedBatchTool, dict[str, Any]]],
) -> list[dict[str, Any]]:
    from word_mcp_codemode_live.core.word_com import undo_record

    results: list[dict[str, Any]] = []
    failed_index: int | None = None
    mutation_attempted = False
    try:
        with _batch_ui_state(app), undo_record(app, undo_name):
            for index, (tool_name, tool, arguments) in enumerate(prepared):
                failed_index = index
                mutation_attempted = True
                nested = await tool(**arguments)
                results.append(
                    {"tool": tool_name, "result": _parse_nested_result(tool_name, nested)}
                )
    except Exception as exc:
        raise BatchExecutionError(exc, failed_index, mutation_attempted) from exc
    return results


def _verification_errors(
    document_text: str,
    stats: dict[str, Any],
    contains_text: list[str] | None,
    absent_text: list[str] | None,
    expected_page_count: int | None,
) -> list[str]:
    errors = [
        f"Required text not found: {snippet!r}"
        for snippet in contains_text or []
        if snippet not in document_text
    ]
    errors.extend(
        f"Forbidden text still present: {snippet!r}"
        for snippet in absent_text or []
        if snippet in document_text
    )
    if expected_page_count is not None and stats["pages"] != expected_page_count:
        errors.append(f"Expected {expected_page_count} pages; Word reports {stats['pages']}")
    return errors


def _render_batch_pages(
    document: Any,
    affected_pages: set[int],
    capture_pages: list[int] | None,
    capture_affected_pages: bool,
    capture_dpi: int,
):
    should_capture = capture_pages is not None or (capture_affected_pages and bool(affected_pages))
    pages = capture_pages
    if pages is None and affected_pages:
        pages = sorted(affected_pages)[:8]
    return render_word_pages(document, pages, dpi=capture_dpi) if should_capture else []


def _success_result(metadata: dict[str, Any], rendered: list[Any]) -> ToolResult:
    content = [
        TextContent(type="text", text=json.dumps(metadata, ensure_ascii=False)),
        *(Image(data=page.data, format="png").to_image_content() for page in rendered),
    ]
    return ToolResult(content=content, structured_content=metadata)


@word_tool(title="Batch Edit and Verify Live Word", domain="batch", change="edit")
async def word_live_edit_batch(
    operations: list[dict[str, Any]],
    filename: str | None = None,
    undo_name: str = "MCP: Batch Edit",
    save: bool = True,
    contains_text: list[str] | None = None,
    absent_text: list[str] | None = None,
    expected_page_count: int | None = None,
    capture_pages: list[int] | None = None,
    capture_affected_pages: bool = True,
    capture_dpi: int = 144,
) -> ToolResult:
    """Apply a verified live Word batch and return affected-page images.

    Each operation has ``tool`` and ``arguments`` keys. Only mutating Word-live
    tools are accepted. Nested tool undo records join this batch's one Word undo
    record. If an operation or assertion fails, the tool rolls back only when it
    can positively identify its own entry at the top of Word's undo history.

    Args:
        operations: Ordered live-tool calls, for example
            ``[{"tool": "word_live_replace_text", "arguments": {...}}]``.
        filename: Open Word document name or full path; injected into every call.
        undo_name: Label shown in Word's Undo menu.
        save: Save after verification succeeds.
        contains_text: Text snippets that must exist after the batch.
        absent_text: Text snippets that must not exist after the batch.
        expected_page_count: Exact expected page count after repagination.
        capture_pages: Explicit one-based pages to render after the batch.
        capture_affected_pages: Infer and render pages touched by the operations.
        capture_dpi: Resolution for returned page images, from 72 to 300 DPI.
    """
    if not operations:
        return _error_result(
            "operations must not be empty", rolled_back=False, operation_index=None
        )

    from word_mcp_codemode_live.core.word_com import find_document, get_word_app

    app = get_word_app()
    document = find_document(app, filename)
    before = _document_stats(document)
    affected_pages: set[int] = set()
    for operation in operations:
        affected_pages.update(_operation_pages(document, operation))

    try:
        prepared = _prepare_operations(operations, filename)
    except BatchPreparationError as exc:
        return _error_result(str(exc), rolled_back=False, operation_index=exc.index)

    timings: dict[str, float] = {}
    actual_undo_name = unique_undo_name(undo_name)
    batch_started = perf_counter()
    edit_started = batch_started
    try:
        results = await _execute_prepared(app, actual_undo_name, prepared)
        timings["edit_seconds"] = round(perf_counter() - edit_started, 4)
    except BatchExecutionError as exc:
        rolled_back = _undo_batch(document, app, actual_undo_name, exc.mutation_attempted)
        return _error_result(str(exc), rolled_back=rolled_back, operation_index=exc.index)

    try:
        repaginate_started = perf_counter()
        document.Repaginate()
        timings["repaginate_seconds"] = round(perf_counter() - repaginate_started, 4)
        for operation in operations:
            affected_pages.update(_operation_pages(document, operation))
        after = _document_stats(document)
        document_text = document.Content.Text

        verify_started = perf_counter()
        verification_errors = _verification_errors(
            document_text, after, contains_text, absent_text, expected_page_count
        )
        if verification_errors:
            raise BatchOperationError("; ".join(verification_errors))

        timings["verify_seconds"] = round(perf_counter() - verify_started, 4)

        capture_started = perf_counter()
        rendered = _render_batch_pages(
            document,
            affected_pages,
            capture_pages,
            capture_affected_pages,
            capture_dpi,
        )
        timings["capture_seconds"] = round(perf_counter() - capture_started, 4)

        if save:
            document.Save()
        after = _document_stats(document)
    except Exception as exc:
        rolled_back = _undo_batch(document, app, actual_undo_name, True)
        try:
            document.Repaginate()
        except Exception as repaginate_exc:
            logger.warning("Could not repaginate after batch rollback: %s", repaginate_exc)
        return _error_result(str(exc), rolled_back=rolled_back, operation_index=None)

    metadata = {
        "success": True,
        "document": document.Name,
        "undo_name": undo_name,
        "undo_record_name": actual_undo_name,
        "operation_count": len(operations),
        "operations": results,
        "before": before,
        "after": after,
        "affected_pages": sorted(affected_pages),
        "captured_pages": [page.page for page in rendered],
        "timings": {
            **timings,
            "total_seconds": round(perf_counter() - batch_started, 4),
        },
    }
    return _success_result(metadata, rendered)
