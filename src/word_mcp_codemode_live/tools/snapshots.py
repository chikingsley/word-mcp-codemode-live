"""Persist and compare semantic snapshots of open Microsoft Word documents."""

import hashlib
import logging
import ntpath
from contextvars import ContextVar
from typing import Annotated, Any

from pydantic import Field

from word_mcp_codemode_live import snapshot_format
from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session

logger = logging.getLogger(__name__)

_CAPTURE_WARNINGS: ContextVar[set[str] | None] = ContextVar(
    "word_snapshot_capture_warnings", default=None
)
_HEADER_FOOTER_VARIANTS = {"primary": 1, "first": 2, "even": 3}
_HEADER_FOOTER_STORY_TYPES = {
    6: ("header", "even"),
    7: ("header", "primary"),
    8: ("footer", "even"),
    9: ("footer", "primary"),
    10: ("header", "first"),
    11: ("footer", "first"),
}


def _clean_text(value: Any) -> str:
    return str(value).rstrip("\r\x07")


def _safe(getter: Any, default: Any = None) -> Any:
    try:
        return getter()
    except Exception as exc:
        warnings = _CAPTURE_WARNINGS.get()
        if warnings is not None:
            warnings.add(
                "One or more optional Word properties were unavailable "
                f"({type(exc).__name__}); corresponding snapshot values are null or defaulted."
            )
        return default


def _style_name(word_range: Any) -> str | None:
    return _safe(lambda: str(word_range.Style.NameLocal), _safe(lambda: str(word_range.Style)))


def _range_record(word_range: Any) -> dict[str, Any]:
    return {
        "start_offset": int(word_range.Start),
        "end_offset": int(word_range.End),
        "text": _clean_text(word_range.Text),
    }


def _paragraphs(document: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for index in range(1, int(document.Paragraphs.Count) + 1):
        paragraph = document.Paragraphs(index)
        word_range = paragraph.Range
        list_format = word_range.ListFormat
        list_type = _safe(lambda list_format=list_format: int(list_format.ListType), 0)
        records.append(
            {
                "index": index,
                **_range_record(word_range),
                "style": _style_name(word_range),
                "outline_level": _safe(lambda paragraph=paragraph: int(paragraph.OutlineLevel)),
                "in_table": bool(
                    _safe(lambda word_range=word_range: word_range.Information(12), False)
                ),
                "list": {
                    "type_id": list_type,
                    "level": _safe(lambda list_format=list_format: int(list_format.ListLevelNumber))
                    if list_type
                    else None,
                    "label": _safe(lambda list_format=list_format: str(list_format.ListString))
                    if list_type
                    else None,
                },
            }
        )
    return records


def _tables(document: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for index in range(1, int(document.Tables.Count) + 1):
        table = document.Tables(index)
        row_count = int(table.Rows.Count)
        column_count = int(table.Columns.Count)
        cells: list[list[str | None]] = []
        for row in range(1, row_count + 1):
            cells.append(
                [
                    _safe(
                        lambda row=row, column=column, table=table: _clean_text(
                            table.Cell(row, column).Range.Text
                        )
                    )
                    for column in range(1, column_count + 1)
                ]
            )
        records.append(
            {
                "index": index,
                "start_offset": int(table.Range.Start),
                "end_offset": int(table.Range.End),
                "row_count": row_count,
                "column_count": column_count,
                # A null cell means Word rejected that grid coordinate, normally
                # because the table contains vertically or horizontally merged cells.
                "cells": cells,
            }
        )
    return records


def _sections(document: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for index in range(1, int(document.Sections.Count) + 1):
        section = document.Sections(index)
        setup = section.PageSetup
        records.append(
            {
                "index": index,
                "start_offset": int(section.Range.Start),
                "end_offset": int(section.Range.End),
                "start_type_id": _safe(lambda section=section: int(section.PageSetup.SectionStart)),
                "page_setup": {
                    "orientation_id": _safe(lambda setup=setup: int(setup.Orientation)),
                    "page_width_points": _safe(lambda setup=setup: float(setup.PageWidth)),
                    "page_height_points": _safe(lambda setup=setup: float(setup.PageHeight)),
                    "top_margin_points": _safe(lambda setup=setup: float(setup.TopMargin)),
                    "bottom_margin_points": _safe(lambda setup=setup: float(setup.BottomMargin)),
                    "left_margin_points": _safe(lambda setup=setup: float(setup.LeftMargin)),
                    "right_margin_points": _safe(lambda setup=setup: float(setup.RightMargin)),
                    "gutter_points": _safe(lambda setup=setup: float(setup.Gutter)),
                    "different_first_page": bool(
                        _safe(lambda setup=setup: setup.DifferentFirstPageHeaderFooter, False)
                    ),
                    "different_odd_even": bool(
                        _safe(lambda setup=setup: setup.OddAndEvenPagesHeaderFooter, False)
                    ),
                },
            }
        )
    return records


def _populated_header_footer_text(document: Any) -> dict[tuple[int, str, str], str]:
    populated: dict[tuple[int, str, str], str] = {}
    for story_type, (kind, variant) in _HEADER_FOOTER_STORY_TYPES.items():
        try:
            word_range = document.StoryRanges(story_type)
        except Exception as exc:
            logger.debug("Header/footer story %s is unavailable: %s", story_type, exc)
            continue
        while word_range is not None:
            section_index = int(word_range.Sections(1).Index)
            populated[(section_index, kind, variant)] = _clean_text(word_range.Text)
            word_range = _safe(lambda word_range=word_range: word_range.NextStoryRange)
    return populated


def _headers_footers(document: Any) -> list[dict[str, Any]]:
    populated_text = _populated_header_footer_text(document)
    effective_text: dict[tuple[str, str], str] = {}
    records: list[dict[str, Any]] = []
    for section_index in range(1, int(document.Sections.Count) + 1):
        section = document.Sections(section_index)
        for kind, collection in (("header", section.Headers), ("footer", section.Footers)):
            for variant, variant_id in _HEADER_FOOTER_VARIANTS.items():
                story = collection(variant_id)
                linked_to_previous = bool(_safe(lambda story=story: story.LinkToPrevious, False))
                text = populated_text.get((section_index, kind, variant))
                if text is None and linked_to_previous:
                    text = effective_text.get((kind, variant), "")
                if text is None:
                    text = ""
                effective_text[(kind, variant)] = text
                records.append(
                    {
                        "section_index": section_index,
                        "kind": kind,
                        "variant": variant,
                        "exists": bool(_safe(lambda story=story: story.Exists, False)),
                        "linked_to_previous": linked_to_previous,
                        # Reading Range.Text for an empty header/footer can
                        # materialize that story and dirty a saved document.
                        "text": text,
                    }
                )
    return records


def _notes(document: Any) -> dict[str, list[dict[str, Any]]]:
    result: dict[str, list[dict[str, Any]]] = {}
    for name, collection in (("footnotes", document.Footnotes), ("endnotes", document.Endnotes)):
        records: list[dict[str, Any]] = []
        for index in range(1, int(collection.Count) + 1):
            note = collection(index)
            records.append(
                {
                    "index": index,
                    "reference_offset": _safe(lambda note=note: int(note.Reference.Start)),
                    "text": _clean_text(note.Range.Text),
                }
            )
        result[name] = records
    return result


def _fields(document: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    collection = document.Fields
    for index in range(1, int(collection.Count) + 1):
        field = collection(index)
        records.append(
            {
                "index": index,
                "type_id": int(field.Type),
                "code": _clean_text(field.Code.Text),
                "result": _safe(lambda field=field: _clean_text(field.Result.Text)),
                "locked": bool(_safe(lambda field=field: field.Locked, False)),
            }
        )
    return records


def _bookmarks(document: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    collection = document.Bookmarks
    for index in range(1, int(collection.Count) + 1):
        bookmark = collection(index)
        records.append({"name": str(bookmark.Name), **_range_record(bookmark.Range)})
    return sorted(records, key=lambda record: (record["name"].casefold(), record["name"]))


def _comments(document: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    collection = document.Comments
    for raw_index in range(1, int(collection.Count) + 1):
        comment = collection(raw_index)
        if _safe(lambda comment=comment: comment.Ancestor) is not None:
            continue
        replies: list[dict[str, Any]] = []
        reply_collection = _safe(lambda comment=comment: comment.Replies)
        if reply_collection is not None:
            for reply_index in range(1, int(reply_collection.Count) + 1):
                reply = reply_collection(reply_index)
                replies.append(
                    {
                        "author": _safe(lambda reply=reply: str(reply.Author), ""),
                        "initial": _safe(lambda reply=reply: str(reply.Initial), ""),
                        "text": _clean_text(reply.Range.Text),
                        "resolved": bool(_safe(lambda reply=reply: reply.Done, False)),
                    }
                )
        records.append(
            {
                "index": len(records) + 1,
                "author": _safe(lambda comment=comment: str(comment.Author), ""),
                "initial": _safe(lambda comment=comment: str(comment.Initial), ""),
                "scope_text": _clean_text(_safe(lambda comment=comment: comment.Scope.Text, "")),
                "text": _clean_text(comment.Range.Text),
                "resolved": bool(_safe(lambda comment=comment: comment.Done, False)),
                "replies": replies,
            }
        )
    return records


def _revisions(document: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    collection = document.Revisions
    for index in range(1, int(collection.Count) + 1):
        revision = collection(index)
        word_range = revision.Range
        records.append(
            {
                "index": index,
                "type_id": int(revision.Type),
                "author": _safe(lambda revision=revision: str(revision.Author), ""),
                **_range_record(word_range),
            }
        )
    return records


def _capture(document: Any) -> dict[str, Any]:
    capture_warnings: set[str] = set()
    token = _CAPTURE_WARNINGS.set(capture_warnings)
    try:
        content = {
            "paragraphs": _paragraphs(document),
            "tables": _tables(document),
            "sections": _sections(document),
            "headers_footers": _headers_footers(document),
            "notes": _notes(document),
            "fields": _fields(document),
            "bookmarks": _bookmarks(document),
            "comments": _comments(document),
            "revisions": _revisions(document),
        }
    finally:
        _CAPTURE_WARNINGS.reset(token)
    source_full_path = str(document.FullName)
    snapshot = {
        "schema": snapshot_format.SCHEMA,
        "version": snapshot_format.VERSION,
        "source": {
            "name": str(document.Name),
            "full_path": source_full_path,
            "normalized_path_sha256": hashlib.sha256(
                ntpath.normcase(source_full_path).encode("utf-8")
            ).hexdigest(),
        },
        "content_sha256": snapshot_format.content_hash(content),
        "capture_warnings": sorted(capture_warnings),
        "capabilities": {
            "captures": [
                "main_story_paragraphs",
                "tables",
                "section_page_setup",
                "header_footer_text",
                "footnotes_endnotes",
                "main_story_fields",
                "bookmarks",
                "comments",
                "revisions",
            ],
            "does_not_capture": [
                "pixel_or_rendered_layout",
                "drawing_geometry",
                "embedded_object_payloads",
                "complete_character_or_paragraph_formatting",
                "revision_history",
            ],
        },
        "content": content,
    }
    snapshot["envelope_sha256"] = snapshot_format.envelope_hash(snapshot)
    return snapshot


@word_tool(title="Word Live Create Document Snapshot", domain="inspection", change="edit")
async def word_live_create_document_snapshot(
    snapshot_path: Annotated[
        str,
        Field(
            min_length=1,
            description="Destination .json path; relative paths resolve from the server working directory.",
        ),
    ],
    filename: str | None = None,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Capture an open document's semantic content and structure to versioned JSON.

    The snapshot is read-only with respect to Word. It is designed for later
    semantic comparison and does not assert visual or pixel-level equivalence.

    Args:
        snapshot_path: Destination .json path. Relative paths resolve from the server's
            current working directory; the parent directory must already exist.
        filename: Open document name or full path (None = active document).
        overwrite: Replace an existing snapshot only when explicitly true.
    """
    word_session.require_windows("Live snapshot capture")
    path = snapshot_format.snapshot_path(snapshot_path)

    document = word_session.find_document(word_session.get_word_app(), filename)
    saved_before = bool(document.Saved)
    snapshot = _capture(document)
    saved_after = bool(document.Saved)
    if saved_after != saved_before:
        raise RuntimeError(
            "Word changed the document saved state during snapshot capture; "
            "no snapshot was written and the document must be inspected"
        )
    snapshot_format.write_snapshot(path, snapshot, overwrite)
    return {
        "success": True,
        "document": str(document.Name),
        "snapshot_path": str(path),
        "schema": snapshot_format.SCHEMA,
        "version": snapshot_format.VERSION,
        "content_sha256": snapshot["content_sha256"],
        "envelope_sha256": snapshot["envelope_sha256"],
        "integrity_scope": "complete_snapshot_envelope_checksum_not_authentication",
        "capture_warnings": snapshot["capture_warnings"],
        "counts": {
            key: len(value)
            if isinstance(value, list)
            else sum(len(items) for items in value.values())
            for key, value in snapshot["content"].items()
        },
        "visual_equivalence_captured": False,
        "saved_state_unchanged": True,
        "sensitivity_warning": (
            "The snapshot contains document text, comments, notes, and metadata. "
            "Store and share it with the same care as the source document."
        ),
    }


@word_tool(title="Word Live Diff Document Snapshots", domain="inspection", change="read")
async def word_live_diff_document_snapshots(
    before_snapshot_path: Annotated[str, Field(min_length=1)],
    after_snapshot_path: Annotated[str, Field(min_length=1)],
    allow_cross_document: bool = False,
) -> dict[str, Any]:
    """Compare two version-1 semantic snapshots without opening or changing Word.

    Paragraphs use a deterministic sequence comparison so insertions do not make
    every later paragraph appear modified. Other components use JSON Pointer paths;
    Word character offsets are excluded from semantic comparison. ``same_source_path``
    means only that both snapshots recorded the same case-insensitive normalized
    Windows full path; it is not a persistent document identity.
    """
    before_path, before = snapshot_format.load_snapshot(before_snapshot_path)
    after_path, after = snapshot_format.load_snapshot(after_snapshot_path)
    before_source = before.get("source", {})
    after_source = after.get("source", {})
    same_source_path = before_source.get("normalized_path_sha256") == after_source.get(
        "normalized_path_sha256"
    )
    if not same_source_path and not allow_cross_document:
        raise ValueError(
            "Snapshots identify different source paths. Set allow_cross_document=true "
            "only when that comparison is intentional."
        )

    before_content = before["content"]
    after_content = after["content"]
    paragraph_operations = snapshot_format.paragraph_diff(
        before_content.get("paragraphs", []), after_content.get("paragraphs", [])
    )
    component_changes: dict[str, list[dict[str, Any]]] = {}
    component_names = sorted((before_content.keys() | after_content.keys()) - {"paragraphs"})
    for component in component_names:
        changes = snapshot_format.recursive_diff(
            snapshot_format.without_offsets(before_content.get(component)),
            snapshot_format.without_offsets(after_content.get(component)),
        )
        if changes:
            component_changes[component] = changes

    paragraph_change_group_count = len(paragraph_operations)
    component_leaf_change_count = sum(len(changes) for changes in component_changes.values())
    return {
        "success": True,
        "before_snapshot_path": str(before_path),
        "after_snapshot_path": str(after_path),
        "same_source_path": same_source_path,
        "source_identity_basis": "case-insensitive normalized Windows full path only",
        "identical_semantic_content": (
            paragraph_change_group_count == 0 and component_leaf_change_count == 0
        ),
        "paragraph_change_group_count": paragraph_change_group_count,
        "component_leaf_change_count": component_leaf_change_count,
        "paragraph_operations": paragraph_operations,
        "component_changes": component_changes,
        "comparison_scope": "semantic_content_and_structure",
        "visual_equivalence_compared": False,
    }
