"""Inspect text, structure, and positions in live Word documents."""

import json
import logging
import sys

from word_mcp_codemode_live.tools.metadata import word_tool

logger = logging.getLogger(__name__)


@word_tool(title="Word Live Get Text", domain="inspection", change="read")
async def word_live_get_text(filename: str | None = None) -> str:
    """Get all text from an open Word document, paragraph by paragraph.

    For documents with more than 200 paragraphs, only the first 3 pages are
    returned along with total page count. Use word_live_get_page_text to read
    specific pages of large documents.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with paragraphs list.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        total_paras = doc.Paragraphs.Count

        # Large document safety cap: return first 3 pages instead of all
        if total_paras > 200:
            total_pages = doc.ComputeStatistics(2)  # wdStatisticPages
            result = json.loads(await word_live_get_page_text(filename, 1, 3))
            result["truncated"] = True
            result["total_paragraphs"] = total_paras
            result["total_pages"] = total_pages
            result["message"] = (
                f"Document has {total_paras} paragraphs across {total_pages} pages. "
                f"Showing first 3 pages only. Use word_live_get_page_text(page=N, end_page=M) "
                f"to read specific pages."
            )
            return json.dumps(result, ensure_ascii=False)

        paragraphs = []
        for i in range(1, total_paras + 1):
            text = doc.Paragraphs(i).Range.Text.rstrip("\r\x07")
            paragraphs.append({"index": i, "text": text})

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "paragraph_count": len(paragraphs),
                "paragraphs": paragraphs,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Get Info", domain="inspection", change="read")
async def word_live_get_info(filename: str | None = None) -> str:
    """Get document info from an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with document metadata (pages, words, paragraphs, sections, etc.).
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        # wdStatistic constants
        WD_STAT_PAGES = 2
        WD_STAT_WORDS = 0
        WD_STAT_CHARACTERS = 3
        WD_STAT_LINES = 1

        info = {
            "name": doc.Name,
            "full_path": doc.FullName,
            "pages": doc.ComputeStatistics(WD_STAT_PAGES),
            "words": doc.ComputeStatistics(WD_STAT_WORDS),
            "characters": doc.ComputeStatistics(WD_STAT_CHARACTERS),
            "lines": doc.ComputeStatistics(WD_STAT_LINES),
            "paragraphs": doc.Paragraphs.Count,
            "sections": doc.Sections.Count,
            "tables": doc.Tables.Count,
            "comments": doc.Comments.Count,
            "track_revisions": doc.TrackRevisions,
            "saved": doc.Saved,
        }

        # Built-in properties (best effort)
        try:
            props = doc.BuiltInDocumentProperties
            info["author"] = str(props("Author").Value) if props("Author").Value else ""
            info["title"] = str(props("Title").Value) if props("Title").Value else ""
            info["subject"] = str(props("Subject").Value) if props("Subject").Value else ""
        except Exception as exc:
            logger.debug("Built-in document properties are unavailable: %s", exc)

        return json.dumps({"success": True, **info}, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Find Text", domain="inspection", change="read")
async def word_live_find_text(
    filename: str | None = None,
    search_text: str = "",
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    context_chars: int = 60,
    max_results: int = 50,
) -> str:
    """Find text in an open Word document.

    Supports Word special characters when use_wildcards=True:
    ^m (manual page break), ^t (tab), ^p (paragraph mark), ^s (non-breaking space), and Word wildcard syntax.
    Note: whole_word is ignored when use_wildcards is True (Word limitation).

    Args:
        filename: Document name or path (None = active document).
        search_text: Text to search for. With use_wildcards=True, supports ^m, ^t, ^p, ^s and Word wildcards.
        match_case: Case-sensitive search.
        whole_word: Match whole words only (ignored when use_wildcards=True).
        use_wildcards: Enable Word wildcards and special characters (^m, ^t, ^p, ^s, etc.).
        context_chars: Characters of context before/after each match (default 60).
        max_results: Maximum number of matches to return.

    Returns:
        JSON with list of matches (position, context).
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if not search_text:
        return json.dumps({"error": "search_text is required"})

    # Same control-byte hazard as replace_text: \x07 and other control
    # bytes corrupt Word's Find engine. Reject before issuing Find.Execute.
    from word_mcp_codemode_live.utils.text_safety import reject_control_chars

    try:
        reject_control_chars("search_text", search_text)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        # COM-marshalling can fail intermittently after MCP reconnect on
        # property access (rng.Text, doc.Name). Wrap each access so a
        # single hiccup yields a partial result rather than aborting.
        def _safe_attr(obj, attr, default=None):
            try:
                return getattr(obj, attr)
            except Exception:
                return default

        matches = []
        partial_errors = []
        rng = doc.Content.Duplicate
        rng.Find.ClearFormatting()

        while len(matches) < max_results:
            try:
                found = rng.Find.Execute(
                    FindText=search_text,
                    MatchCase=match_case,
                    MatchWholeWord=whole_word if not use_wildcards else False,
                    MatchWildcards=use_wildcards,
                    Forward=True,
                    Wrap=0,  # wdFindStop
                )
            except Exception as e:
                partial_errors.append(f"Find.Execute failed: {e}")
                break
            if not found:
                break

            match_start = _safe_attr(rng, "Start", -1)
            match_end = _safe_attr(rng, "End", -1)

            try:
                context_rng = rng.Duplicate
                content_end = _safe_attr(doc.Content, "End", match_end)
                context_start = max(0, match_start - context_chars) if match_start >= 0 else 0
                context_end = (
                    min(content_end, match_end + context_chars) if match_end >= 0 else context_chars
                )
                context_rng.SetRange(context_start, context_end)
                context_text = _safe_attr(context_rng, "Text", "<unreadable>")
            except Exception as e:
                context_text = f"<context unavailable: {e}>"

            matches.append(
                {
                    "start": match_start,
                    "end": match_end,
                    "text": _safe_attr(rng, "Text", "<unreadable>"),
                    "context": context_text,
                }
            )

            # Move past current match — guard against transient COM failures
            try:
                rng.SetRange(match_end if match_end >= 0 else rng.End, doc.Content.End)
            except Exception as e:
                partial_errors.append(f"advance past match failed: {e}")
                break

        result = {
            "success": True,
            "document": _safe_attr(doc, "Name", "<unknown>"),
            "search_text": search_text,
            "match_count": len(matches),
            "matches": matches,
        }
        if partial_errors:
            result["partial_errors"] = partial_errors
        return json.dumps(result, ensure_ascii=False)

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Get Page Text", domain="inspection", change="read")
async def word_live_get_page_text(
    filename: str | None = None,
    page: int = 1,
    end_page: int | None = None,
) -> str:
    """[Windows only] Get text from specific page(s) of an open Word document.

    Returns paragraphs on the requested page(s) with char_start/char_end offsets
    that can be passed directly to word_live_format_text, word_live_delete_text, etc.

    Uses Word's GoTo API to find page boundaries. For long legal documents, this
    is much more efficient than reading all paragraphs.

    Args:
        filename: Document name or path (None = active document).
        page: Page number to read (1-indexed, required).
        end_page: Last page to read (inclusive). If None, reads only `page`.

    Returns:
        JSON with paragraphs list, each containing index, text, char_start, char_end.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if page < 1:
        return json.dumps({"error": "page must be >= 1"})

    if end_page is not None and end_page < page:
        return json.dumps({"error": "end_page must be >= page"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        # wdStatisticPages = 2
        total_pages = doc.ComputeStatistics(2)

        if page > total_pages:
            return json.dumps(
                {"error": f"Page {page} out of range (document has {total_pages} pages)"}
            )

        if end_page is None:
            end_page = page

        if end_page > total_pages:
            end_page = total_pages

        # wdGoToPage=1, wdGoToAbsolute=1
        # Get start of requested page
        page_start_range = doc.GoTo(What=1, Which=1, Count=page)
        range_start = page_start_range.Start

        # Get start of page after end_page (or end of doc)
        if end_page < total_pages:
            next_page_range = doc.GoTo(What=1, Which=1, Count=end_page + 1)
            range_end = next_page_range.Start
        else:
            range_end = doc.Content.End

        # Collect paragraphs within the page range
        paragraphs = []
        for i in range(1, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            p_start = para.Range.Start
            p_end = para.Range.End

            # Skip paragraphs entirely before our range
            if p_end <= range_start:
                continue
            # Stop once we pass our range
            if p_start >= range_end:
                break

            text = para.Range.Text.rstrip("\r\x07")
            paragraphs.append(
                {
                    "index": i,
                    "text": text,
                    "char_start": p_start,
                    "char_end": p_end,
                }
            )

        page_label = f"{page}" if page == end_page else f"{page}-{end_page}"
        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "pages": page_label,
                "total_pages": total_pages,
                "paragraph_count": len(paragraphs),
                "range_start": range_start,
                "range_end": range_end,
                "paragraphs": paragraphs,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})
