# ty: ignore[unresolved-import]

"""Transactional, verified batches for live Microsoft Word editing."""

import inspect
import json
from collections.abc import Awaitable, Callable
from time import perf_counter
from typing import Any

from fastmcp.tools.tool import ToolResult
from fastmcp.utilities.types import Image
from mcp.types import TextContent

from word_mcp_codemode_live.tools.page_capture_tools import render_word_pages

BatchTool = Callable[..., Awaitable[str]]


class BatchOperationError(RuntimeError):
    """Raised when a nested live tool reports an error."""


def _tool_catalog() -> dict[str, BatchTool]:
    """Return the deliberately limited set of mutating tools allowed in a batch."""
    from word_mcp_codemode_live.tools import live_layout_tools, live_read_tools, live_tools

    return {
        "word_live_insert_text": live_tools.word_live_insert_text,
        "word_live_format_text": live_tools.word_live_format_text,
        "word_live_replace_text": live_tools.word_live_replace_text,
        "word_live_insert_paragraphs": live_tools.word_live_insert_paragraphs,
        "word_live_add_table": live_tools.word_live_add_table,
        "word_live_format_table": live_tools.word_live_format_table,
        "word_live_modify_table": live_tools.word_live_modify_table,
        "word_live_delete_text": live_tools.word_live_delete_text,
        "word_live_apply_list": live_tools.word_live_apply_list,
        "word_live_setup_heading_numbering": live_tools.word_live_setup_heading_numbering,
        "word_live_toggle_track_changes": live_tools.word_live_toggle_track_changes,
        "word_live_insert_image": live_tools.word_live_insert_image,
        "word_live_insert_cross_reference": live_tools.word_live_insert_cross_reference,
        "word_live_insert_equation": live_tools.word_live_insert_equation,
        "word_live_set_page_layout": live_layout_tools.word_live_set_page_layout,
        "word_live_add_header_footer": live_layout_tools.word_live_add_header_footer,
        "word_live_add_page_numbers": live_layout_tools.word_live_add_page_numbers,
        "word_live_add_section_break": live_layout_tools.word_live_add_section_break,
        "word_live_set_paragraph_spacing": live_layout_tools.word_live_set_paragraph_spacing,
        "word_live_add_bookmark": live_layout_tools.word_live_add_bookmark,
        "word_live_add_watermark": live_layout_tools.word_live_add_watermark,
        "word_live_set_core_properties": live_read_tools.word_live_set_core_properties,
        "word_live_add_comment": live_read_tools.word_live_add_comment,
        "word_live_reply_to_comment": live_read_tools.word_live_reply_to_comment,
        "word_live_resolve_comment": live_read_tools.word_live_resolve_comment,
        "word_live_delete_comment": live_read_tools.word_live_delete_comment,
        "word_live_accept_revisions": live_read_tools.word_live_accept_revisions,
        "word_live_reject_revisions": live_read_tools.word_live_reject_revisions,
    }


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
    if isinstance(parsed, dict) and parsed.get("error"):
        raise BatchOperationError(f"{tool_name}: {parsed['error']}")
    return parsed


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

    try:
        control = app.CommandBars.FindControl(Type=6, Id=128)
        if control is not None and control.ListCount:
            latest = str(control.List(1))
            if undo_name[:64].casefold() not in latest.casefold():
                return False
    except Exception:
        # The undo dropdown is undocumented and unavailable in some Word builds.
        pass

    try:
        return bool(document.Undo(1))
    except Exception:
        return False


async def word_live_edit_batch(
    operations: list[dict[str, Any]],
    filename: str | None = None,
    undo_name: str = "MCP: Batch Edit",
    save: bool = True,
    contains_text: list[str] | None = None,
    absent_text: list[str] | None = None,
    expected_page_count: int | None = None,
    verify_layout: bool = False,
    capture_pages: list[int] | None = None,
    capture_affected_pages: bool = True,
    capture_dpi: int = 144,
) -> ToolResult:
    """Apply live Word edits atomically, verify them, and return affected-page images.

    Each operation has ``tool`` and ``arguments`` keys. Only mutating Word-live
    tools are accepted. Nested tool undo records join this batch's one Word undo
    record. If an operation or an explicit assertion fails, Word rolls the whole
    batch back.

    Args:
        operations: Ordered live-tool calls, for example
            ``[{"tool": "word_live_replace_text", "arguments": {...}}]``.
        filename: Open Word document name or full path; injected into every call.
        undo_name: Label shown in Word's Undo menu.
        save: Save after verification succeeds.
        contains_text: Text snippets that must exist after the batch.
        absent_text: Text snippets that must not exist after the batch.
        expected_page_count: Exact expected page count after repagination.
        verify_layout: Run the slower full-document structural diagnostic. Text
            assertions, repagination, document statistics, and page rendering
            still run when this is false.
        capture_pages: Explicit one-based pages to render after the batch.
        capture_affected_pages: Infer and render pages touched by the operations.
        capture_dpi: Resolution for returned page images, from 72 to 300 DPI.
    """
    if not operations:
        return _error_result(
            "operations must not be empty", rolled_back=False, operation_index=None
        )

    from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record
    from word_mcp_codemode_live.tools.live_read_tools import word_live_diagnose_layout

    catalog = _tool_catalog()
    app = get_word_app()
    document = find_document(app, filename)
    before = _document_stats(document)
    affected_pages: set[int] = set()
    for operation in operations:
        affected_pages.update(_operation_pages(document, operation))

    prepared: list[tuple[str, BatchTool, dict[str, Any]]] = []
    for index, operation in enumerate(operations):
        tool_name = operation.get("tool")
        arguments = operation.get("arguments", {})
        if not isinstance(tool_name, str) or tool_name not in catalog:
            return _error_result(
                f"Unsupported batch tool: {tool_name!r}",
                rolled_back=False,
                operation_index=index,
            )
        if not isinstance(arguments, dict):
            return _error_result(
                f"Operation {index} arguments must be an object",
                rolled_back=False,
                operation_index=index,
            )
        nested_filename = arguments.get("filename")
        if filename and nested_filename and nested_filename != filename:
            return _error_result(
                f"Operation {index} targets a different document: {nested_filename!r}",
                rolled_back=False,
                operation_index=index,
            )
        call_arguments = dict(arguments)
        tool = catalog[tool_name]
        signature = inspect.signature(tool)
        if "filename" in signature.parameters:
            call_arguments["filename"] = filename
        try:
            signature.bind(**call_arguments)
        except TypeError as exc:
            return _error_result(str(exc), rolled_back=False, operation_index=index)
        prepared.append((tool_name, tool, call_arguments))

    old_screen_updating = bool(app.ScreenUpdating)
    try:
        old_pagination = bool(app.Options.Pagination)
    except Exception:
        old_pagination = True

    results: list[dict[str, Any]] = []
    timings: dict[str, float] = {}
    failed_index: int | None = None
    mutation_attempted = False
    batch_started = perf_counter()
    edit_started = batch_started
    try:
        app.ScreenUpdating = False
        try:
            app.Options.Pagination = False
        except Exception:
            pass

        with undo_record(app, undo_name):
            for index, (tool_name, tool, call_arguments) in enumerate(prepared):
                failed_index = index
                mutation_attempted = True
                nested_result = await tool(**call_arguments)
                parsed_result = _parse_nested_result(tool_name, nested_result)
                results.append({"tool": tool_name, "result": parsed_result})
        failed_index = None
        timings["edit_seconds"] = round(perf_counter() - edit_started, 4)
    except Exception as exc:
        rolled_back = _undo_batch(document, app, undo_name, mutation_attempted)
        return _error_result(str(exc), rolled_back=rolled_back, operation_index=failed_index)
    finally:
        try:
            app.Options.Pagination = old_pagination
        except Exception:
            pass
        app.ScreenUpdating = old_screen_updating

    try:
        repaginate_started = perf_counter()
        document.Repaginate()
        timings["repaginate_seconds"] = round(perf_counter() - repaginate_started, 4)
        for operation in operations:
            affected_pages.update(_operation_pages(document, operation))
        after = _document_stats(document)
        document_text = document.Content.Text

        verify_started = perf_counter()
        verification_errors: list[str] = []
        for snippet in contains_text or []:
            if snippet not in document_text:
                verification_errors.append(f"Required text not found: {snippet!r}")
        for snippet in absent_text or []:
            if snippet in document_text:
                verification_errors.append(f"Forbidden text still present: {snippet!r}")
        if expected_page_count is not None and after["pages"] != expected_page_count:
            verification_errors.append(
                f"Expected {expected_page_count} pages; Word reports {after['pages']}"
            )
        if verification_errors:
            raise BatchOperationError("; ".join(verification_errors))

        layout_report = None
        if verify_layout:
            layout_report = _parse_nested_result(
                "word_live_diagnose_layout", await word_live_diagnose_layout(filename=filename)
            )
        timings["verify_seconds"] = round(perf_counter() - verify_started, 4)

        should_capture = capture_pages is not None or capture_affected_pages
        pages_to_capture = capture_pages
        if pages_to_capture is None and affected_pages:
            pages_to_capture = sorted(affected_pages)[:8]
        capture_started = perf_counter()
        rendered = (
            render_word_pages(document, pages_to_capture, dpi=capture_dpi) if should_capture else []
        )
        timings["capture_seconds"] = round(perf_counter() - capture_started, 4)

        if save:
            document.Save()
        after = _document_stats(document)
    except Exception as exc:
        rolled_back = _undo_batch(document, app, undo_name, mutation_attempted)
        try:
            document.Repaginate()
        except Exception:
            pass
        return _error_result(str(exc), rolled_back=rolled_back, operation_index=None)

    metadata = {
        "success": True,
        "document": document.Name,
        "undo_name": undo_name,
        "operation_count": len(operations),
        "operations": results,
        "before": before,
        "after": after,
        "affected_pages": sorted(affected_pages),
        "captured_pages": [page.page for page in rendered],
        "layout": layout_report,
        "timings": {
            **timings,
            "total_seconds": round(perf_counter() - batch_started, 4),
        },
    }
    content = [
        TextContent(type="text", text=json.dumps(metadata, ensure_ascii=False)),
        *(Image(data=page.data, format="png").to_image_content() for page in rendered),
    ]
    return ToolResult(content=content, structured_content=metadata)
