"""Persist and compare semantic snapshots of open Microsoft Word documents."""

import hashlib
import json
import logging
import ntpath
import os
import sys
import tempfile
from contextvars import ContextVar
from difflib import SequenceMatcher
from pathlib import Path
from typing import Annotated, Any

from pydantic import Field

from word_mcp_codemode_live.tools.metadata import word_tool

logger = logging.getLogger(__name__)

_SCHEMA = "word-mcp-live.document-snapshot"
_VERSION = 1
_MAX_SNAPSHOT_BYTES = 50 * 1024 * 1024
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
_OFFSET_KEYS = {
    "start_offset",
    "end_offset",
    "reference_offset",
    "code_start_offset",
    "code_end_offset",
    "result_start_offset",
    "result_end_offset",
}


def _require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Live snapshot capture is only available on Windows")


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


def _canonical_bytes(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode(
        "utf-8"
    )


def _content_hash(content: dict[str, Any]) -> str:
    return hashlib.sha256(_canonical_bytes(content)).hexdigest()


def _envelope_hash(snapshot: dict[str, Any]) -> str:
    envelope = {key: value for key, value in snapshot.items() if key != "envelope_sha256"}
    return hashlib.sha256(_canonical_bytes(envelope)).hexdigest()


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
        "schema": _SCHEMA,
        "version": _VERSION,
        "source": {
            "name": str(document.Name),
            "full_path": source_full_path,
            "normalized_path_sha256": hashlib.sha256(
                ntpath.normcase(source_full_path).encode("utf-8")
            ).hexdigest(),
        },
        "content_sha256": _content_hash(content),
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
    snapshot["envelope_sha256"] = _envelope_hash(snapshot)
    return snapshot


def _snapshot_path(path: str) -> Path:
    if not path or not path.strip():
        raise ValueError("snapshot_path is required")
    candidate = Path(path).expanduser().resolve(strict=False)
    if candidate.suffix.casefold() != ".json":
        raise ValueError("snapshot_path must use a .json extension")
    if not candidate.parent.is_dir():
        raise ValueError(f"Snapshot parent directory does not exist: {candidate.parent}")
    if candidate.exists() and not candidate.is_file():
        raise ValueError(f"Snapshot path is not a file: {candidate}")
    return candidate


def _write_snapshot(path: Path, snapshot: dict[str, Any], overwrite: bool) -> None:
    serialized = json.dumps(snapshot, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    if len(serialized.encode("utf-8")) > _MAX_SNAPSHOT_BYTES:
        raise ValueError(
            f"Serialized snapshot exceeds the {_MAX_SNAPSHOT_BYTES // (1024 * 1024)} MiB limit"
        )
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(serialized)
            handle.flush()
            os.fsync(handle.fileno())
        if overwrite:
            os.replace(temporary_name, path)
        else:
            try:
                # This tool is Windows-only; os.rename publishes atomically and
                # refuses an existing destination instead of replacing it.
                os.rename(temporary_name, path)
            except FileExistsError as exc:
                raise FileExistsError(
                    f"Snapshot already exists: {path}. Set overwrite=true to replace it."
                ) from exc
    except Exception:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise


def _validate_snapshot_content(content: Any, path: Path) -> dict[str, Any]:
    if not isinstance(content, dict):
        raise ValueError(f"Snapshot content must be a JSON object: {path}")
    list_components = {
        "paragraphs",
        "tables",
        "sections",
        "headers_footers",
        "fields",
        "bookmarks",
        "comments",
        "revisions",
    }
    for component in sorted(list_components):
        records = content.get(component)
        if not isinstance(records, list) or any(not isinstance(item, dict) for item in records):
            raise ValueError(f"Snapshot content.{component} must be a list of objects: {path}")
    notes = content.get("notes")
    if not isinstance(notes, dict) or set(notes) != {"footnotes", "endnotes"}:
        raise ValueError(f"Snapshot content.notes must contain footnotes and endnotes: {path}")
    if any(
        not isinstance(notes[kind], list) or any(not isinstance(item, dict) for item in notes[kind])
        for kind in ("footnotes", "endnotes")
    ):
        raise ValueError(f"Snapshot note collections must be lists of objects: {path}")
    return content


def _validate_snapshot_source(snapshot: dict[str, Any], path: Path) -> None:
    source = snapshot.get("source")
    if not isinstance(source, dict):
        raise ValueError(f"Snapshot source must be a JSON object: {path}")
    if not isinstance(source.get("full_path"), str) or not source["full_path"]:
        raise ValueError(f"Snapshot source full_path must be a nonempty string: {path}")
    identity = source.get("normalized_path_sha256")
    expected_identity = hashlib.sha256(
        ntpath.normcase(source["full_path"]).encode("utf-8")
    ).hexdigest()
    if identity != expected_identity:
        raise ValueError(f"Snapshot source normalized_path_sha256 is invalid: {path}")


def _load_snapshot(path_value: str) -> tuple[Path, dict[str, Any]]:
    path = _snapshot_path(path_value)
    if not path.is_file():
        raise FileNotFoundError(f"Snapshot does not exist: {path}")
    if path.stat().st_size > _MAX_SNAPSHOT_BYTES:
        raise ValueError(f"Snapshot exceeds the {_MAX_SNAPSHOT_BYTES // (1024 * 1024)} MiB limit")
    try:
        snapshot = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read snapshot JSON {path}: {exc}") from exc
    if not isinstance(snapshot, dict):
        raise ValueError(f"Snapshot root must be a JSON object: {path}")
    if snapshot.get("schema") != _SCHEMA or snapshot.get("version") != _VERSION:
        raise ValueError(
            f"Unsupported snapshot schema/version in {path}; expected {_SCHEMA} version {_VERSION}"
        )
    if snapshot.get("envelope_sha256") != _envelope_hash(snapshot):
        raise ValueError(f"Snapshot envelope hash mismatch: {path}")
    content = _validate_snapshot_content(snapshot.get("content"), path)
    _validate_snapshot_source(snapshot, path)
    expected_hash = snapshot.get("content_sha256")
    actual_hash = _content_hash(content)
    if expected_hash != actual_hash:
        raise ValueError(f"Snapshot content hash mismatch: {path}")
    return path, snapshot


def _without_offsets(value: Any) -> Any:
    if isinstance(value, dict):
        return {
            key: _without_offsets(item) for key, item in value.items() if key not in _OFFSET_KEYS
        }
    if isinstance(value, list):
        return [_without_offsets(item) for item in value]
    return value


def _paragraph_semantic(paragraph: dict[str, Any]) -> dict[str, Any]:
    return {
        key: _without_offsets(value)
        for key, value in paragraph.items()
        if key not in {"index", *_OFFSET_KEYS}
    }


def _paragraph_diff(before: list[Any], after: list[Any]) -> list[dict[str, Any]]:
    before_semantic = [_paragraph_semantic(item) for item in before]
    after_semantic = [_paragraph_semantic(item) for item in after]
    matcher = SequenceMatcher(
        None,
        [_canonical_bytes(item) for item in before_semantic],
        [_canonical_bytes(item) for item in after_semantic],
        autojunk=False,
    )
    operations: list[dict[str, Any]] = []
    for operation, before_start, before_end, after_start, after_end in matcher.get_opcodes():
        if operation == "equal":
            continue
        operations.append(
            {
                "operation": operation,
                "before_start_index": before_start + 1,
                "before_count": before_end - before_start,
                "after_start_index": after_start + 1,
                "after_count": after_end - after_start,
                "before": before_semantic[before_start:before_end],
                "after": after_semantic[after_start:after_end],
            }
        )
    return operations


def _escape_pointer(value: str) -> str:
    return value.replace("~", "~0").replace("/", "~1")


def _recursive_diff(before: Any, after: Any, path: str = "") -> list[dict[str, Any]]:
    if type(before) is not type(after):
        return [{"path": path or "/", "operation": "replace", "before": before, "after": after}]
    if isinstance(before, dict):
        changes: list[dict[str, Any]] = []
        for key in sorted(before.keys() | after.keys()):
            child_path = f"{path}/{_escape_pointer(str(key))}"
            if key not in before:
                changes.append({"path": child_path, "operation": "add", "after": after[key]})
            elif key not in after:
                changes.append({"path": child_path, "operation": "remove", "before": before[key]})
            else:
                changes.extend(_recursive_diff(before[key], after[key], child_path))
        return changes
    if isinstance(before, list):
        changes = []
        shared = min(len(before), len(after))
        for index in range(shared):
            changes.extend(_recursive_diff(before[index], after[index], f"{path}/{index}"))
        for index in range(shared, len(before)):
            changes.append(
                {"path": f"{path}/{index}", "operation": "remove", "before": before[index]}
            )
        for index in range(shared, len(after)):
            changes.append({"path": f"{path}/{index}", "operation": "add", "after": after[index]})
        return changes
    if before != after:
        return [{"path": path or "/", "operation": "replace", "before": before, "after": after}]
    return []


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
    _require_windows()
    path = _snapshot_path(snapshot_path)
    from word_mcp_codemode_live.core.word_com import find_document, get_word_app

    document = find_document(get_word_app(), filename)
    saved_before = bool(document.Saved)
    snapshot = _capture(document)
    saved_after = bool(document.Saved)
    if saved_after != saved_before:
        raise RuntimeError(
            "Word changed the document saved state during snapshot capture; "
            "no snapshot was written and the document must be inspected"
        )
    _write_snapshot(path, snapshot, overwrite)
    return {
        "success": True,
        "document": str(document.Name),
        "snapshot_path": str(path),
        "schema": _SCHEMA,
        "version": _VERSION,
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
    before_path, before = _load_snapshot(before_snapshot_path)
    after_path, after = _load_snapshot(after_snapshot_path)
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
    paragraph_operations = _paragraph_diff(
        before_content.get("paragraphs", []), after_content.get("paragraphs", [])
    )
    component_changes: dict[str, list[dict[str, Any]]] = {}
    component_names = sorted((before_content.keys() | after_content.keys()) - {"paragraphs"})
    for component in component_names:
        changes = _recursive_diff(
            _without_offsets(before_content.get(component)),
            _without_offsets(after_content.get(component)),
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
