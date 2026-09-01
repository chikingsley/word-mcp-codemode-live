"""Inspect the native outline structure of an open Word document."""

from typing import Any

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session

_WD_ACTIVE_END_PAGE_NUMBER = 3
_BODY_TEXT_LEVEL = 10


@word_tool(title="Word Live Inspect Document Outline", domain="inspection", change="read")
async def word_live_inspect_document_outline(
    filename: str | None = None,
    maximum_level: int = 9,
    include_body_text: bool = False,
) -> dict[str, Any]:
    """Return paragraphs by Word outline level, including heading and custom outline styles.

    Paragraph indexes are one-based. Character positions are zero-based offsets.
    ``page_number`` is Word's active-end page number and may trigger repagination.
    """
    word_session.require_windows("Live outline tools")
    if not 1 <= maximum_level <= 9:
        raise ValueError("maximum_level must be between 1 and 9")

    document = word_session.find_document(word_session.get_word_app(), filename)
    entries: list[dict[str, Any]] = []
    outlined_count = 0
    body_count = 0
    for paragraph_index in range(1, int(document.Paragraphs.Count) + 1):
        paragraph = document.Paragraphs(paragraph_index)
        level = int(paragraph.OutlineLevel)
        is_body = level == _BODY_TEXT_LEVEL
        if is_body:
            body_count += 1
            if not include_body_text:
                continue
        elif not 1 <= level <= maximum_level:
            continue
        else:
            outlined_count += 1

        word_range = paragraph.Range
        try:
            style_name = str(word_range.Style.NameLocal)
        except Exception:
            style_name = str(word_range.Style)
        try:
            page_number = int(word_range.Information(_WD_ACTIVE_END_PAGE_NUMBER))
        except Exception:
            page_number = None
        entries.append(
            {
                "paragraph_index": paragraph_index,
                "outline_level": level,
                "is_body_text": is_body,
                "style": style_name,
                "start_offset": int(word_range.Start),
                "end_offset": int(word_range.End),
                "page_number": page_number,
                "text": str(word_range.Text).rstrip("\r\x07"),
            }
        )

    return {
        "success": True,
        "document": str(document.Name),
        "maximum_level": maximum_level,
        "include_body_text": include_body_text,
        "outline_entry_count": outlined_count,
        "body_paragraph_count": body_count,
        "returned_count": len(entries),
        "entries": entries,
    }
