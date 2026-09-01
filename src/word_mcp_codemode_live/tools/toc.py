"""Create, inspect, update, and delete native Microsoft Word TOCs."""

import sys
from typing import Any, Literal

from word_mcp_codemode_live.tools.metadata import word_tool

TocTarget = Literal["document_start", "document_end", "paragraph_start", "paragraph_end"]
TocUpdateMode = Literal["all", "page_numbers"]
_TOC_TARGETS = {"document_start", "document_end", "paragraph_start", "paragraph_end"}


def _require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Live table-of-contents tools are only available on Windows")


def _toc_entry(toc: Any, index: int) -> dict[str, Any]:
    word_range = toc.Range
    entry: dict[str, Any] = {
        "index": index,
        "start_offset": int(word_range.Start),
        "end_offset": int(word_range.End),
        "text": str(word_range.Text),
        "use_heading_styles": bool(toc.UseHeadingStyles),
        "upper_heading_level": int(toc.UpperHeadingLevel),
        "lower_heading_level": int(toc.LowerHeadingLevel),
        "include_page_numbers": bool(toc.IncludePageNumbers),
        "right_align_page_numbers": bool(toc.RightAlignPageNumbers),
        "use_hyperlinks": bool(toc.UseHyperlinks),
        "hide_page_numbers_in_web": bool(toc.HidePageNumbersInWeb),
    }
    try:
        entry["field_code"] = str(word_range.Fields(1).Code.Text)
    except Exception:
        entry["field_code"] = None
    return entry


def _get_toc(document: Any, toc_index: int) -> Any:
    count = int(document.TablesOfContents.Count)
    if not 1 <= toc_index <= count:
        raise ValueError(f"toc_index must be between 1 and {count}")
    return document.TablesOfContents(toc_index)


def _toc_index(collection: Any, toc: Any) -> int:
    """Locate a newly created TOC without assuming it was appended in range order."""
    target_range = toc.Range
    for index in range(1, int(collection.Count) + 1):
        candidate = collection(index)
        if candidate is toc:
            return index
        candidate_range = candidate.Range
        if int(candidate_range.Start) == int(target_range.Start) and int(
            candidate_range.End
        ) == int(target_range.End):
            return index
    raise RuntimeError("Word created a table of contents but it could not be located")


def _target_range(document: Any, target: TocTarget, paragraph_index: int | None) -> Any:
    paragraph_target = target in {"paragraph_start", "paragraph_end"}
    if paragraph_target and paragraph_index is None:
        raise ValueError(f"paragraph_index is required when target={target!r}")
    if not paragraph_target and paragraph_index is not None:
        raise ValueError("paragraph_index is only valid for a paragraph target")

    if paragraph_target:
        paragraph_count = int(document.Paragraphs.Count)
        if paragraph_index is None or not 1 <= paragraph_index <= paragraph_count:
            raise ValueError(f"paragraph_index must be between 1 and {paragraph_count}")
        paragraph_range = document.Paragraphs(paragraph_index).Range
        offset = int(paragraph_range.Start)
        if target == "paragraph_end":
            # Insert immediately before the paragraph mark, not in the next paragraph.
            offset = max(int(paragraph_range.Start), int(paragraph_range.End) - 1)
    elif target == "document_start":
        offset = int(document.Content.Start)
    else:
        # A Word document's Content ends after its mandatory final paragraph mark.
        offset = max(int(document.Content.Start), int(document.Content.End) - 1)
    return document.Range(offset, offset)


@word_tool(title="Word Live List Tables of Contents", domain="references", change="read")
async def word_live_list_tables_of_contents(filename: str | None = None) -> dict[str, Any]:
    """Inspect every native table of contents in an open Word document.

    TOC indexes are one-based. Range properties are explicitly named zero-based
    offsets to match Word's native character-position model.
    """
    _require_windows()
    from word_mcp_codemode_live.core.word_com import find_document, get_word_app

    document = find_document(get_word_app(), filename)
    collection = document.TablesOfContents
    entries = [
        _toc_entry(collection(index), index) for index in range(1, int(collection.Count) + 1)
    ]
    return {
        "success": True,
        "document": str(document.Name),
        "toc_count": len(entries),
        "tables_of_contents": entries,
    }


@word_tool(
    title="Word Live Create Table of Contents",
    domain="references",
    change="edit",
    batchable=True,
)
async def word_live_create_table_of_contents(
    filename: str | None = None,
    target: TocTarget = "document_start",
    paragraph_index: int | None = None,
    upper_heading_level: int = 1,
    lower_heading_level: int = 3,
    include_page_numbers: bool = True,
    right_align_page_numbers: bool = True,
    use_hyperlinks: bool = True,
    hide_page_numbers_in_web: bool = True,
) -> dict[str, Any]:
    """Create a genuine Word TOC field at an explicit collapsed range.

    ``paragraph_index`` is one-based and required only for ``paragraph_start``
    or ``paragraph_end``. Heading levels are inclusive and must be from 1 to 9.
    """
    _require_windows()
    if target not in _TOC_TARGETS:
        raise ValueError(f"target must be one of {sorted(_TOC_TARGETS)}")
    if not 1 <= upper_heading_level <= 9:
        raise ValueError("upper_heading_level must be between 1 and 9")
    if not 1 <= lower_heading_level <= 9:
        raise ValueError("lower_heading_level must be between 1 and 9")
    if upper_heading_level > lower_heading_level:
        raise ValueError("upper_heading_level cannot exceed lower_heading_level")

    from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

    app = get_word_app()
    document = find_document(app, filename)
    word_range = _target_range(document, target, paragraph_index)
    insertion_offset = int(word_range.Start)
    with undo_record(app, "MCP: Create Table of Contents"):
        toc = document.TablesOfContents.Add(
            Range=word_range,
            UseHeadingStyles=True,
            UpperHeadingLevel=upper_heading_level,
            LowerHeadingLevel=lower_heading_level,
            UseFields=False,
            TableID="",
            RightAlignPageNumbers=right_align_page_numbers,
            IncludePageNumbers=include_page_numbers,
            AddedStyles="",
            UseHyperlinks=use_hyperlinks,
            HidePageNumbersInWeb=hide_page_numbers_in_web,
            UseOutlineLevels=True,
        )
    index = _toc_index(document.TablesOfContents, toc)
    return {
        "success": True,
        "document": str(document.Name),
        "created_at_offset": insertion_offset,
        "table_of_contents": _toc_entry(toc, index),
    }


@word_tool(
    title="Word Live Update Table of Contents",
    domain="references",
    change="edit",
    batchable=True,
)
async def word_live_update_table_of_contents(
    filename: str | None = None,
    toc_index: int = 1,
    mode: TocUpdateMode = "all",
) -> dict[str, Any]:
    """Update one native TOC by its one-based index.

    ``mode='all'`` refreshes entries and page numbers. ``mode='page_numbers'``
    preserves the entry set and refreshes page numbers only.
    """
    _require_windows()
    if mode not in {"all", "page_numbers"}:
        raise ValueError("mode must be all or page_numbers")
    from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

    app = get_word_app()
    document = find_document(app, filename)
    toc = _get_toc(document, toc_index)
    with undo_record(app, "MCP: Update Table of Contents"):
        if mode == "all":
            toc.Update()
        else:
            toc.UpdatePageNumbers()
    return {
        "success": True,
        "document": str(document.Name),
        "mode": mode,
        "table_of_contents": _toc_entry(toc, toc_index),
    }


@word_tool(
    title="Word Live Delete Table of Contents",
    domain="references",
    change="edit",
    batchable=True,
)
async def word_live_delete_table_of_contents(
    filename: str | None = None,
    toc_index: int = 1,
) -> dict[str, Any]:
    """Delete one native TOC by its one-based index."""
    _require_windows()
    from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

    app = get_word_app()
    document = find_document(app, filename)
    toc = _get_toc(document, toc_index)
    deleted = _toc_entry(toc, toc_index)
    with undo_record(app, "MCP: Delete Table of Contents"):
        toc.Delete()
    return {
        "success": True,
        "document": str(document.Name),
        "deleted_table_of_contents": deleted,
        "remaining_toc_count": int(document.TablesOfContents.Count),
    }
