"""Inspect and update native Microsoft Word fields in open documents."""

import logging
from typing import Any

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word.stories import collect_story_ranges

logger = logging.getLogger(__name__)

_FIELD_NAMES = {
    -1: "empty",
    3: "ref",
    4: "index_entry",
    5: "footnote_ref",
    6: "set",
    7: "if",
    8: "index",
    10: "style_ref",
    12: "sequence",
    13: "table_of_contents",
    15: "title",
    16: "subject",
    17: "author",
    18: "keywords",
    19: "comments",
    20: "last_saved_by",
    21: "create_date",
    22: "save_date",
    23: "print_date",
    24: "revision_number",
    25: "edit_time",
    26: "number_of_pages",
    27: "number_of_words",
    28: "number_of_characters",
    29: "file_name",
    30: "template",
    31: "date",
    32: "time",
    33: "page",
    34: "expression",
    35: "quote",
    36: "include",
    37: "page_ref",
    38: "ask",
    39: "fill_in",
    48: "print",
    49: "equation",
    56: "link",
    57: "symbol",
    58: "embed",
    59: "merge_field",
    64: "document_variable",
    65: "section",
    66: "section_pages",
    67: "include_picture",
    68: "include_text",
    69: "file_size",
    72: "note_ref",
    73: "table_of_authorities",
    79: "auto_text",
    85: "document_property",
    88: "hyperlink",
    90: "list_number",
    96: "citation",
    97: "bibliography",
}


def _field_rows(document: Any) -> list[tuple[dict[str, Any], Any]]:
    rows: list[tuple[dict[str, Any], Any]] = []
    story_ranges, _ = collect_story_ranges(document)
    for story in story_ranges:
        story_range = story.com_range
        fields = story_range.Fields
        for story_field_index in range(1, int(fields.Count) + 1):
            field = fields(story_field_index)
            field_type_id = int(field.Type)
            code_range = field.Code
            try:
                result_range = field.Result
                result_available = True
                result_start_offset = int(result_range.Start)
                result_end_offset = int(result_range.End)
                result_text: str | None = str(result_range.Text)
            except Exception:
                result_available = False
                result_start_offset = None
                result_end_offset = None
                result_text = None
            rows.append(
                (
                    {
                        "index": len(rows) + 1,
                        "story": story.name,
                        "story_type_id": story.story_type,
                        "story_instance_index": story.instance_index,
                        "story_field_index": story_field_index,
                        "type": _FIELD_NAMES.get(field_type_id, "unknown"),
                        "type_id": field_type_id,
                        # Word character positions are zero-based offsets, unlike
                        # every public collection index returned by these tools.
                        "code_start_offset": int(code_range.Start),
                        "code_end_offset": int(code_range.End),
                        "result_available": result_available,
                        "result_start_offset": result_start_offset,
                        "result_end_offset": result_end_offset,
                        "code": str(code_range.Text),
                        "result": result_text,
                        "locked": bool(field.Locked),
                    },
                    field,
                )
            )
    return rows


@word_tool(title="Word Live List Fields", domain="references", change="read")
async def word_live_list_fields(filename: str | None = None) -> dict[str, Any]:
    """List native Word fields across every populated document story.

    ``index``, ``story_instance_index``, and ``story_field_index`` are one-based.
    Properties ending in ``_offset`` are Word's zero-based character offsets.

    Args:
        filename: Open document name or full path (None = active document).
    """
    word_session.require_windows("Live field tools")

    document = word_session.find_document(word_session.get_word_app(), filename)
    fields = [row for row, _field in _field_rows(document)]
    return {
        "success": True,
        "document": str(document.Name),
        "field_count": len(fields),
        "fields": fields,
    }


@word_tool(
    title="Word Live Update Fields",
    domain="references",
    change="edit",
    batchable=True,
)
async def word_live_update_fields(
    filename: str | None = None,
    field_indices: list[int] | None = None,
) -> dict[str, Any]:
    """Update all native fields or selected fields in an open document.

    Args:
        filename: Open document name or full path (None = active document).
        field_indices: Optional unique one-based indexes from ``word_live_list_fields``.
            Omit to update every field across every document story.
    """
    word_session.require_windows("Live field tools")
    if field_indices is not None:
        if not field_indices:
            raise ValueError("field_indices cannot be empty; omit it to update all fields")
        if any(index < 1 for index in field_indices):
            raise ValueError("field_indices must contain only one-based positive indexes")
        if len(set(field_indices)) != len(field_indices):
            raise ValueError("field_indices must not contain duplicates")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    rows = _field_rows(document)
    selected = list(range(1, len(rows) + 1)) if field_indices is None else field_indices
    missing = [index for index in selected if index > len(rows)]
    if missing:
        raise ValueError(
            f"field_indices contains indexes outside the available range 1-{len(rows)}: {missing}"
        )

    updated: list[dict[str, Any]] = []
    with word_session.undo_record(app, "MCP: Update Fields"):
        for index in selected:
            before, field = rows[index - 1]
            try:
                update_result = field.Update()
            except Exception as exc:
                raise RuntimeError(f"Word failed to update field {index}: {exc}") from exc
            try:
                after_result: str | None = str(field.Result.Text)
                result_available = True
            except Exception:
                after_result = None
                result_available = False
            updated.append(
                {
                    "index": index,
                    "story": before["story"],
                    "type": before["type"],
                    "update_result": None if update_result is None else bool(update_result),
                    "result_available": result_available,
                    "result": after_result,
                }
            )

    return {
        "success": True,
        "document": str(document.Name),
        "available_field_count": len(rows),
        "updated_count": len(updated),
        "updated_fields": updated,
    }


@word_tool(
    title="Word Live Unlink Fields",
    domain="references",
    change="edit",
    batchable=True,
)
async def word_live_unlink_fields(
    filename: str | None = None,
    field_indices: list[int] | None = None,
) -> dict[str, Any]:
    """Replace selected native fields with their current displayed results.

    Unlinking converts each field result to ordinary text or a graphic, so the
    result will no longer update automatically. Current Word leaves index-entry
    (XE) fields linked, so requests containing them are rejected before editing.

    Args:
        filename: Open document name or full path (None = active document).
        field_indices: Unique one-based indexes from ``word_live_list_fields``.
            Omit to unlink every unlinkable field. If omitted and the document
            contains a known non-unlinkable field, no fields are changed.
    """
    word_session.require_windows("Live field tools")
    if field_indices is not None:
        if not field_indices:
            raise ValueError("field_indices cannot be empty; omit it to unlink all fields")
        if any(index < 1 for index in field_indices):
            raise ValueError("field_indices must contain only one-based positive indexes")
        if len(set(field_indices)) != len(field_indices):
            raise ValueError("field_indices must not contain duplicates")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    rows = _field_rows(document)
    selected = list(range(1, len(rows) + 1)) if field_indices is None else field_indices
    missing = [index for index in selected if index > len(rows)]
    if missing:
        raise ValueError(
            f"field_indices contains indexes outside the available range 1-{len(rows)}: {missing}"
        )

    # Word 16 silently leaves XE index-entry fields linked. SEQ fields are
    # unlinkable in current Word despite older Microsoft documentation saying
    # otherwise, so verify actual post-state instead of rejecting them.
    unsupported = [index for index in selected if rows[index - 1][0]["type_id"] == 4]
    if unsupported:
        raise ValueError(
            "Word cannot unlink index-entry (XE) fields; "
            f"remove these indexes from the request: {unsupported}"
        )

    unlinked_by_index: dict[int, dict[str, Any]] = {}
    with word_session.undo_transaction(app, document, "MCP: Unlink Fields"):
        # Work backwards so unlinking a field cannot invalidate later fields in
        # the same story. Return entries in the caller's requested order below.
        for index in sorted(selected, reverse=True):
            before, field = rows[index - 1]
            try:
                field.Unlink()
            except Exception as exc:
                raise RuntimeError(f"Word failed to unlink field {index}: {exc}") from exc
            unlinked_by_index[index] = {
                "index": index,
                "story": before["story"],
                "type": before["type"],
                "result": before["result"],
            }
        remaining_count = len(_field_rows(document))
        expected_remaining = len(rows) - len(selected)
        if remaining_count != expected_remaining:
            raise RuntimeError(
                "Word did not unlink every requested field: "
                f"expected {expected_remaining} remaining, found {remaining_count}"
            )

    unlinked = [unlinked_by_index[index] for index in selected]
    return {
        "success": True,
        "document": str(document.Name),
        "available_field_count": len(rows),
        "unlinked_count": len(unlinked),
        "unlinked_fields": unlinked,
    }
