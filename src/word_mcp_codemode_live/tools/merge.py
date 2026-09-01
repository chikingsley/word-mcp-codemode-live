"""Insert complete files into live Word documents with native Word semantics."""

import os
import sys
from collections import Counter
from pathlib import Path
from typing import Annotated, Any, Literal

from pydantic import Field

from word_mcp_codemode_live.tools.metadata import word_tool


def _require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Live editing is only available on Windows")


def _resolve_source_path(source_path: str) -> Path:
    if not source_path or not source_path.strip():
        raise ValueError("source_path is required")
    source = Path(source_path).expanduser()
    if not source.is_absolute():
        raise ValueError("source_path must be an absolute path")
    try:
        source = source.resolve(strict=True)
    except FileNotFoundError as exc:
        raise ValueError(f"Source file does not exist: {source}") from exc
    if not source.is_file():
        raise ValueError(f"source_path must identify a file: {source}")
    return source


def _collection_count(document: Any, name: str) -> int:
    try:
        return int(getattr(document, name).Count)
    except Exception:
        return 0


def _style_inventory(document: Any) -> dict[str, str]:
    inventory: dict[str, str] = {}
    for index in range(1, int(document.Styles.Count) + 1):
        name = str(document.Styles(index).NameLocal)
        inventory[name.casefold()] = name
    return inventory


def _remove_imported_styles(document: Any, before: dict[str, str]) -> None:
    current = _style_inventory(document)
    introduced = [current[key] for key in current.keys() - before.keys()]
    for name in introduced:
        style = document.Styles(name)
        if bool(style.BuiltIn):
            raise RuntimeError(f"refusing to delete newly surfaced built-in style {name!r}")
        style.Delete()
    remaining = [name for key, name in _style_inventory(document).items() if key not in before]
    if remaining:
        raise RuntimeError(f"imported styles remain after rollback cleanup: {remaining}")


def _list_template_inventory(document: Any) -> list[str]:
    return [
        str(document.ListTemplates(index).Name)
        for index in range(1, int(document.ListTemplates.Count) + 1)
    ]


def _introduced_list_templates(document: Any, before: list[str]) -> list[str]:
    introduced = Counter(_list_template_inventory(document)) - Counter(before)
    return sorted(introduced.elements(), key=str.casefold)


def _rollback_definition_cleanup(
    document: Any,
    before_styles: dict[str, str],
    before_list_templates: list[str],
    before_structure: dict[str, int],
    before_text: str,
    saved_before: bool,
) -> None:
    _remove_imported_styles(document, before_styles)
    residual_list_templates = _introduced_list_templates(document, before_list_templates)
    if residual_list_templates:
        raise RuntimeError(
            "Word Undo left imported list-template definitions that the Word object "
            f"model cannot delete: {residual_list_templates}"
        )
    current_structure = _structure_counts(document)
    if current_structure != before_structure or str(document.Content.Text) != before_text:
        raise RuntimeError(
            "Word Undo and definition cleanup did not restore the exact pre-insert "
            "content and structural counts"
        )
    if bool(document.Saved) != saved_before:
        document.Saved = saved_before
    if bool(document.Saved) != saved_before:
        raise RuntimeError("Word did not restore the pre-insert saved state")


def _structure_counts(document: Any) -> dict[str, int]:
    """Return stable native-object counts useful for verifying InsertFile effects."""
    return {
        "characters": max(0, int(document.Content.End) - 1),
        "paragraphs": _collection_count(document, "Paragraphs"),
        "sections": _collection_count(document, "Sections"),
        "tables": _collection_count(document, "Tables"),
        "fields": _collection_count(document, "Fields"),
        "comments": _collection_count(document, "Comments"),
        "footnotes": _collection_count(document, "Footnotes"),
        "endnotes": _collection_count(document, "Endnotes"),
        "revisions": _collection_count(document, "Revisions"),
        "styles": _collection_count(document, "Styles"),
    }


def _count_deltas(before: dict[str, int], after: dict[str, int]) -> dict[str, int]:
    return {name: after[name] - before[name] for name in before}


def _target_range(
    document: Any,
    *,
    position: str | None,
    target_start: int | None,
    target_end: int | None,
) -> tuple[Any, dict[str, Any]]:
    has_position = position is not None
    has_start = target_start is not None
    has_end = target_end is not None
    if has_start != has_end:
        raise ValueError("target_start and target_end must be provided together")
    if has_position == has_start:
        raise ValueError(
            "Provide exactly one target: position='start'/'end' or target_start and target_end"
        )

    content_end = max(0, int(document.Content.End) - 1)
    if has_position:
        if position not in {"start", "end"}:
            raise ValueError("position must be 'start' or 'end'")
        offset = 0 if position == "start" else content_end
        return document.Range(offset, offset), {
            "kind": "position",
            "position": position,
            "start": offset,
            "end": offset,
        }

    assert target_start is not None and target_end is not None
    if target_start < 0 or target_end < 0:
        raise ValueError("target_start and target_end must be non-negative")
    if target_start > target_end:
        raise ValueError("target_start must be less than or equal to target_end")
    if target_end > content_end:
        raise ValueError(
            f"Target range {target_start}-{target_end} is outside document range 0-{content_end}"
        )
    return document.Range(target_start, target_end), {
        "kind": "range",
        "start": target_start,
        "end": target_end,
        "replaced_characters": target_end - target_start,
    }


@word_tool(
    title="Word Live Insert File",
    domain="content",
    change="edit",
)
async def word_live_insert_file(
    source_path: Annotated[
        str,
        Field(min_length=1, description="Absolute path to the Word file to insert."),
    ],
    filename: str | None = None,
    position: Literal["start", "end"] | None = None,
    target_start: Annotated[int, Field(ge=0)] | None = None,
    target_end: Annotated[int, Field(ge=0)] | None = None,
    source_bookmark: str | None = None,
) -> dict[str, Any]:
    """Insert a file using Word's native ``Range.InsertFile`` operation.

    This is a native Word merge: it does not reconstruct source paragraphs.
    Consequently, sections, styles, fields, comments, notes, and revisions are
    carried over exactly to the extent supported by the installed Word version
    and source format. Word resolves same-named style and other definition
    conflicts using its own destination-document rules; this tool does not
    reconcile or rename them. ``target_start``/``target_end`` are zero-based
    Word character positions; a non-collapsed range is replaced by the file.
    This operation is intentionally not batchable because Word Undo removes inserted
    content but can leave imported style and list-template definitions. On a failed
    direct call, this tool removes only newly imported styles and reports any residual
    list templates, which the Word object model cannot delete. A later manual Undo of
    a successful insertion has Word's same native definition-residue limitation.

    Args:
        source_path: Absolute path to an existing source file.
        filename: Open destination document name or full path (None = active).
        position: Explicit collapsed target, either ``start`` or ``end``.
        target_start: Start of an explicit destination range, inclusive.
        target_end: End of an explicit destination range, exclusive.
        source_bookmark: Optional bookmark/range name inside the source file.
    """
    _require_windows()
    source = _resolve_source_path(source_path)
    if source_bookmark is not None and not source_bookmark.strip():
        raise ValueError("source_bookmark cannot be empty or whitespace")

    from word_mcp_codemode_live.core.word_com import (
        find_document,
        get_word_app,
        undo_transaction,
    )

    app = get_word_app()
    document = find_document(app, filename)
    destination_path = Path(str(document.FullName)).resolve(strict=False)
    try:
        same_file = os.path.samefile(source, destination_path)
    except OSError:
        same_file = os.path.normcase(str(source)) == os.path.normcase(str(destination_path))
    if same_file:
        raise ValueError("source_path must not identify the destination document")

    target, resolved_target = _target_range(
        document,
        position=position,
        target_start=target_start,
        target_end=target_end,
    )
    insertion_start = int(target.Start)
    before_styles = _style_inventory(document)
    before_list_templates = _list_template_inventory(document)
    before = _structure_counts(document)
    before_text = str(document.Content.Text)
    saved_before = bool(document.Saved)

    with undo_transaction(
        app,
        document,
        "MCP: Insert File",
        rollback_cleanup=lambda: _rollback_definition_cleanup(
            document,
            before_styles,
            before_list_templates,
            before,
            before_text,
            saved_before,
        ),
    ):
        try:
            target.InsertFile(
                str(source),
                source_bookmark or "",
                False,  # ConfirmConversions: never display a conversion prompt.
                False,  # Link: insert content, not an INCLUDETEXT field.
                False,  # Attachment: relevant only to email messages.
            )
        except Exception as exc:
            raise RuntimeError(f"Word could not insert source file {str(source)!r}: {exc}") from exc

        after = _structure_counts(document)
        replaced_characters = int(resolved_target.get("replaced_characters", 0))
        inserted_characters = after["characters"] - before["characters"] + replaced_characters
        if inserted_characters < 0:
            raise RuntimeError(
                "Word returned an invalid character count after InsertFile: "
                f"computed inserted length {inserted_characters}"
            )
        # Range.InsertFile leaves a collapsed destination Range collapsed in
        # current Word, so derive the inserted span from the native document
        # character delta instead of trusting target.End.
        insertion_end = insertion_start + inserted_characters
        introduced_styles = [
            name for key, name in _style_inventory(document).items() if key not in before_styles
        ]
        introduced_list_templates = _introduced_list_templates(document, before_list_templates)
        deltas = _count_deltas(before, after)
    return {
        "success": True,
        "document": str(document.Name),
        "source_path": str(source),
        "source_bookmark": source_bookmark,
        "target": resolved_target,
        "inserted_range": {"start": insertion_start, "end": insertion_end},
        "before": before,
        "after": after,
        "deltas": deltas,
        "introduced_styles": introduced_styles,
        "introduced_list_templates": introduced_list_templates,
        "native_operation": "Range.InsertFile",
        "undo_limitation": (
            "Word Undo removes inserted content but may leave imported style and "
            "list-template definitions; list templates cannot be deleted through "
            "Word's object model."
        ),
    }
