"""Discover, insert, and update native Microsoft Word cross-references."""

import sys
from typing import Any, Literal

from word_mcp_codemode_live.tools.metadata import word_tool

ReferenceType = Literal[
    "numbered_item",
    "heading",
    "bookmark",
    "footnote",
    "endnote",
    "figure",
    "table",
    "equation",
]
ReferenceKind = Literal[
    "content_text",
    "number_full_context",
    "number_no_context",
    "number_relative_context",
    "page_number",
    "position",
    "entire_caption",
    "only_label_and_number",
    "only_caption_text",
    "footnote_number",
    "footnote_number_formatted",
    "endnote_number",
    "endnote_number_formatted",
]
InsertPosition = Literal["start", "end"]

# WdReferenceType plus the three built-in WdCaptionLabelID values. Caption
# labels are valid ReferenceType arguments to GetCrossReferenceItems and
# Range.InsertCrossReference.
_REFERENCE_TYPES: dict[str, int] = {
    "numbered_item": 0,
    "heading": 1,
    "bookmark": 2,
    "footnote": 3,
    "endnote": 4,
    "figure": -1,
    "table": -2,
    "equation": -3,
}

# WdReferenceKind values. The previous implementation used positive values for
# the heading-number kinds; Word's actual constants are negative.
_REFERENCE_KINDS: dict[str, int] = {
    "content_text": -1,
    "number_full_context": -4,
    "number_no_context": -3,
    "number_relative_context": -2,
    "page_number": 7,
    "position": 15,
    "entire_caption": 2,
    "only_label_and_number": 3,
    "only_caption_text": 4,
    "footnote_number": 5,
    "footnote_number_formatted": 16,
    "endnote_number": 6,
    "endnote_number_formatted": 17,
}

_NUMBERED_KINDS = {
    "content_text",
    "number_full_context",
    "number_no_context",
    "number_relative_context",
    "page_number",
    "position",
}
_ALLOWED_KINDS: dict[str, set[str]] = {
    "numbered_item": _NUMBERED_KINDS,
    "heading": _NUMBERED_KINDS,
    "bookmark": _NUMBERED_KINDS,
    "footnote": {
        "footnote_number",
        "footnote_number_formatted",
        "page_number",
        "position",
    },
    "endnote": {
        "endnote_number",
        "endnote_number_formatted",
        "page_number",
        "position",
    },
    "figure": {
        "entire_caption",
        "only_label_and_number",
        "only_caption_text",
        "page_number",
        "position",
    },
    "table": {
        "entire_caption",
        "only_label_and_number",
        "only_caption_text",
        "page_number",
        "position",
    },
    "equation": {
        "entire_caption",
        "only_label_and_number",
        "only_caption_text",
        "page_number",
        "position",
    },
}


def _native_items(document: Any, reference_type: ReferenceType) -> list[str]:
    """Return Word's Cross-reference dialog items without synthesizing IDs."""
    raw_items = document.GetCrossReferenceItems(_REFERENCE_TYPES[reference_type])
    if raw_items is None:
        return []
    if isinstance(raw_items, str):
        return [raw_items]
    return [str(item) for item in raw_items]


def _insertion_range(
    document: Any,
    *,
    position: InsertPosition,
    paragraph_index: int | None,
) -> Any:
    if paragraph_index is not None:
        paragraph_count = int(document.Paragraphs.Count)
        if not 1 <= paragraph_index <= paragraph_count:
            raise ValueError(f"paragraph_index must be between 1 and {paragraph_count}")
        word_range = document.Paragraphs(paragraph_index).Range.Duplicate
        word_range.Collapse(1)  # wdCollapseStart
        return word_range

    if position == "start":
        return document.Range(0, 0)
    if position == "end":
        end = max(0, int(document.Content.End) - 1)
        return document.Range(end, end)
    raise ValueError("position must be start or end")


@word_tool(
    title="Word Live List Cross-Reference Targets",
    domain="references",
    change="read",
)
async def word_live_list_cross_reference_targets(
    filename: str | None = None,
    reference_type: ReferenceType = "heading",
) -> dict[str, Any]:
    """List targets reported by Word's native Cross-reference dialog API.

    The returned ``target_index`` is the one-based position in Word's current
    native list. It is intentionally not presented as a durable target ID; ask
    Word for the list again after editing headings, captions, bookmarks, or
    notes.

    Args:
        filename: Open document name/path (None selects the active document).
        reference_type: Native target family to query.

    Returns:
        Structured target records containing type, one-based position, and the
        exact label returned by Word.
    """
    if sys.platform != "win32":
        raise RuntimeError("Live cross-reference tools are only available on Windows")
    if reference_type not in _REFERENCE_TYPES:
        raise ValueError(f"reference_type must be one of: {', '.join(sorted(_REFERENCE_TYPES))}")

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        document = find_document(get_word_app(), filename)
        labels = _native_items(document, reference_type)
        targets = [
            {
                "reference_type": reference_type,
                "target_index": index,
                "label": label,
            }
            for index, label in enumerate(labels, start=1)
        ]
        return {
            "success": True,
            "document": str(document.Name),
            "reference_type": reference_type,
            "count": len(targets),
            "targets": targets,
            "index_stability": "native_list_position; rediscover after target edits",
        }
    except (ValueError, RuntimeError):
        raise
    except Exception as exc:
        raise RuntimeError(f"Word could not list cross-reference targets: {exc}") from exc


@word_tool(
    title="Word Live Insert Cross-Reference",
    domain="references",
    change="edit",
    batchable=True,
)
async def word_live_insert_cross_reference(
    filename: str | None = None,
    reference_type: ReferenceType = "heading",
    target_index: int = 1,
    reference_kind: ReferenceKind = "content_text",
    position: InsertPosition = "end",
    paragraph_index: int | None = None,
    insert_as_hyperlink: bool = True,
    include_position: bool = False,
) -> dict[str, Any]:
    """Insert a native Word cross-reference field at a precise range.

    First call ``word_live_list_cross_reference_targets`` and pass back its
    one-based ``target_index``. For bookmarks, Word requires the discovered
    bookmark label rather than its list position; this tool performs that
    conversion without claiming the label is a stable external identifier.

    Args:
        filename: Open document name/path (None selects the active document).
        reference_type: Native target family.
        target_index: One-based position from native target discovery.
        reference_kind: Information displayed by the inserted reference.
        position: Insert at document start/end when paragraph_index is omitted.
        paragraph_index: Optional one-based paragraph whose start is used.
        insert_as_hyperlink: Make the field clickable in Word.
        include_position: Also include Word's relative "above"/"below" text.

    Returns:
        Structured insertion details including the selected native label.
    """
    if sys.platform != "win32":
        raise RuntimeError("Live cross-reference tools are only available on Windows")
    if reference_type not in _REFERENCE_TYPES:
        raise ValueError(f"reference_type must be one of: {', '.join(sorted(_REFERENCE_TYPES))}")
    if reference_kind not in _REFERENCE_KINDS:
        raise ValueError(f"reference_kind must be one of: {', '.join(sorted(_REFERENCE_KINDS))}")
    if position not in {"start", "end"}:
        raise ValueError("position must be start or end")
    if target_index < 1:
        raise ValueError("target_index must be at least 1")
    if reference_kind not in _ALLOWED_KINDS[reference_type]:
        allowed = ", ".join(sorted(_ALLOWED_KINDS[reference_type]))
        raise ValueError(
            f"reference_kind {reference_kind!r} is not valid for {reference_type!r}; "
            f"use one of: {allowed}"
        )

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        app = get_word_app()
        document = find_document(app, filename)
        labels = _native_items(document, reference_type)
        if target_index > len(labels):
            raise ValueError(
                f"target_index must be between 1 and {len(labels)} for {reference_type!r}"
            )
        label = labels[target_index - 1]
        reference_item: int | str = label if reference_type == "bookmark" else target_index
        word_range = _insertion_range(
            document,
            position=position,
            paragraph_index=paragraph_index,
        )

        with undo_record(app, "MCP: Insert Cross-Reference"):
            word_range.InsertCrossReference(
                ReferenceType=_REFERENCE_TYPES[reference_type],
                ReferenceKind=_REFERENCE_KINDS[reference_kind],
                ReferenceItem=reference_item,
                InsertAsHyperlink=insert_as_hyperlink,
                IncludePosition=include_position,
            )

        return {
            "success": True,
            "document": str(document.Name),
            "reference_type": reference_type,
            "reference_kind": reference_kind,
            "target": {
                "reference_type": reference_type,
                "target_index": target_index,
                "label": label,
            },
            "inserted_at": (
                {"paragraph_index": paragraph_index, "location": "start"}
                if paragraph_index is not None
                else {"document_position": position}
            ),
            "insert_as_hyperlink": insert_as_hyperlink,
            "include_position": include_position,
        }
    except (ValueError, RuntimeError):
        raise
    except Exception as exc:
        raise RuntimeError(f"Word could not insert the cross-reference: {exc}") from exc
