"""Inspect text, structure, and positions in live Word documents."""

import logging
from typing import Annotated, Any

from pydantic import Field

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.validation import reject_control_chars
from word_mcp_codemode_live.word import session as word_session

logger = logging.getLogger(__name__)
PositiveIndex = Annotated[int, Field(ge=1)]
NonNegativeInt = Annotated[int, Field(ge=0)]


def _safe_attr(obj: Any, attribute: str, default: Any = None) -> Any:
    """Read one COM attribute without aborting a partial inspection result."""
    try:
        return getattr(obj, attribute)
    except Exception:
        return default


def _page_text(document: Any, page: int, end_page: int | None) -> dict[str, Any]:
    """Collect paragraph text for an inclusive page range from an attached document."""
    total_pages = int(document.ComputeStatistics(2))  # wdStatisticPages
    if page > total_pages:
        raise ValueError(f"Page {page} out of range (document has {total_pages} pages)")
    final_page = min(end_page or page, total_pages)

    page_start_range = document.GoTo(What=1, Which=1, Count=page)
    range_start = int(page_start_range.Start)
    if final_page < total_pages:
        range_end = int(document.GoTo(What=1, Which=1, Count=final_page + 1).Start)
    else:
        range_end = int(document.Content.End)

    paragraphs: list[dict[str, Any]] = []
    for index in range(1, document.Paragraphs.Count + 1):
        paragraph = document.Paragraphs(index)
        paragraph_start = int(paragraph.Range.Start)
        paragraph_end = int(paragraph.Range.End)
        if paragraph_end <= range_start:
            continue
        if paragraph_start >= range_end:
            break
        paragraphs.append(
            {
                "index": index,
                "text": str(paragraph.Range.Text).rstrip("\r\x07"),
                "char_start": paragraph_start,
                "char_end": paragraph_end,
            }
        )

    return {
        "success": True,
        "document": str(document.Name),
        "pages": f"{page}" if page == final_page else f"{page}-{final_page}",
        "total_pages": total_pages,
        "paragraph_count": len(paragraphs),
        "range_start": range_start,
        "range_end": range_end,
        "paragraphs": paragraphs,
    }


@word_tool(title="Word Live Get Text", domain="inspection", change="read")
async def word_live_get_text(filename: str | None = None) -> dict[str, Any]:
    """Get paragraph text, truncating very large documents to their first three pages."""
    word_session.require_windows("Live Word tools")
    document = word_session.find_document(word_session.get_word_app(), filename)
    total_paragraphs = int(document.Paragraphs.Count)

    if total_paragraphs > 200:
        result = _page_text(document, 1, 3)
        total_pages = int(result["total_pages"])
        result.update(
            {
                "truncated": True,
                "total_paragraphs": total_paragraphs,
                "message": (
                    f"Document has {total_paragraphs} paragraphs across {total_pages} pages. "
                    "Showing first 3 pages only. Use word_live_get_page_text to read "
                    "specific pages."
                ),
            }
        )
        return result

    paragraphs = [
        {
            "index": index,
            "text": str(document.Paragraphs(index).Range.Text).rstrip("\r\x07"),
        }
        for index in range(1, total_paragraphs + 1)
    ]
    return {
        "success": True,
        "document": str(document.Name),
        "paragraph_count": len(paragraphs),
        "paragraphs": paragraphs,
    }


@word_tool(title="Word Live Get Info", domain="inspection", change="read")
async def word_live_get_info(filename: str | None = None) -> dict[str, Any]:
    """Get document statistics and best-effort built-in metadata."""
    word_session.require_windows("Live Word tools")
    document = word_session.find_document(word_session.get_word_app(), filename)
    info: dict[str, Any] = {
        "success": True,
        "name": str(document.Name),
        "full_path": str(document.FullName),
        "pages": int(document.ComputeStatistics(2)),
        "words": int(document.ComputeStatistics(0)),
        "characters": int(document.ComputeStatistics(3)),
        "lines": int(document.ComputeStatistics(1)),
        "paragraphs": int(document.Paragraphs.Count),
        "sections": int(document.Sections.Count),
        "tables": int(document.Tables.Count),
        "comments": int(document.Comments.Count),
        "track_revisions": bool(document.TrackRevisions),
        "saved": bool(document.Saved),
    }
    try:
        properties = document.BuiltInDocumentProperties
        for key, name in (("author", "Author"), ("title", "Title"), ("subject", "Subject")):
            value = properties(name).Value
            info[key] = str(value) if value else ""
    except Exception as exc:
        logger.debug("Built-in document properties are unavailable: %s", exc)
    return info


@word_tool(title="Word Live Find Text", domain="inspection", change="read")
async def word_live_find_text(
    filename: str | None = None,
    search_text: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    context_chars: NonNegativeInt = 60,
    max_results: PositiveIndex = 50,
) -> dict[str, Any]:
    """Find text using Word's native Find engine and return offsets with context."""
    word_session.require_windows("Live Word tools")
    if not search_text:
        raise ValueError("search_text is required")
    reject_control_chars("search_text", search_text)

    document = word_session.find_document(word_session.get_word_app(), filename)
    matches: list[dict[str, Any]] = []
    partial_errors: list[str] = []
    search_range = document.Content.Duplicate
    search_range.Find.ClearFormatting()

    while len(matches) < max_results:
        try:
            found = search_range.Find.Execute(
                FindText=search_text,
                MatchCase=match_case,
                MatchWholeWord=whole_word if not use_wildcards else False,
                MatchWildcards=use_wildcards,
                Forward=True,
                Wrap=0,
            )
        except Exception as exc:
            partial_errors.append(f"Find.Execute failed: {exc}")
            break
        if not found:
            break

        match_start = int(_safe_attr(search_range, "Start", -1))
        match_end = int(_safe_attr(search_range, "End", -1))
        try:
            context_range = search_range.Duplicate
            content_end = int(_safe_attr(document.Content, "End", match_end))
            context_start = max(0, match_start - context_chars) if match_start >= 0 else 0
            context_end = (
                min(content_end, match_end + context_chars) if match_end >= 0 else context_chars
            )
            context_range.SetRange(context_start, context_end)
            context_text = _safe_attr(context_range, "Text", "<unreadable>")
        except Exception as exc:
            context_text = f"<context unavailable: {exc}>"
        matches.append(
            {
                "start": match_start,
                "end": match_end,
                "text": _safe_attr(search_range, "Text", "<unreadable>"),
                "context": context_text,
            }
        )
        try:
            search_range.SetRange(
                match_end if match_end >= 0 else search_range.End,
                document.Content.End,
            )
        except Exception as exc:
            partial_errors.append(f"advance past match failed: {exc}")
            break

    result: dict[str, Any] = {
        "success": True,
        "document": _safe_attr(document, "Name", "<unknown>"),
        "search_text": search_text,
        "match_count": len(matches),
        "matches": matches,
    }
    if partial_errors:
        result["partial_errors"] = partial_errors
    return result


@word_tool(title="Word Live Get Page Text", domain="inspection", change="read")
async def word_live_get_page_text(
    filename: str | None = None,
    page: PositiveIndex = 1,
    end_page: PositiveIndex | None = None,
) -> dict[str, Any]:
    """Get paragraphs and character offsets for an inclusive page range."""
    word_session.require_windows("Live Word tools")
    if page < 1:
        raise ValueError("page must be >= 1")
    if end_page is not None and end_page < page:
        raise ValueError("end_page must be >= page")
    document = word_session.find_document(word_session.get_word_app(), filename)
    return _page_text(document, page, end_page)
