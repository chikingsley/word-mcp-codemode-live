"""Inspect highlighted text across the populated stories of an open Word document."""

import sys
from typing import Any

from word_mcp_codemode_live.tools.metadata import word_tool

_STORY_NAMES = {
    1: "main_text",
    2: "footnotes",
    3: "endnotes",
    4: "comments",
    5: "text_frames",
    6: "even_pages_header",
    7: "primary_header",
    8: "even_pages_footer",
    9: "primary_footer",
    10: "first_page_header",
    11: "first_page_footer",
    12: "footnote_separator",
    13: "footnote_continuation_separator",
    14: "footnote_continuation_notice",
    15: "endnote_separator",
    16: "endnote_continuation_separator",
    17: "endnote_continuation_notice",
}
_COLOR_NAMES = {
    1: "black",
    2: "blue",
    3: "turquoise",
    4: "bright_green",
    5: "pink",
    6: "red",
    7: "yellow",
    8: "white",
    9: "dark_blue",
    10: "teal",
    11: "green",
    12: "violet",
    13: "dark_red",
    14: "dark_yellow",
    15: "gray_50",
    16: "gray_25",
}
_WD_FIND_STOP = 0
_WD_UNDEFINED = 9999999


def _require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Live highlight tools are only available on Windows")


def _story_ranges(document: Any) -> tuple[list[tuple[int, int, Any]], list[str]]:
    stories: list[tuple[int, int, Any]] = []
    skipped: list[str] = []
    for story_type, story_name in _STORY_NAMES.items():
        try:
            story_range = document.StoryRanges(story_type)
        except Exception:
            skipped.append(story_name)
            continue
        instance_index = 1
        while story_range is not None:
            stories.append((story_type, instance_index, story_range))
            instance_index += 1
            if instance_index > 10_000:
                raise RuntimeError(f"Word returned a cyclic story chain for story {story_type}")
            try:
                story_range = story_range.NextStoryRange
            except Exception:
                story_range = None
    return stories, skipped


def _highlight_ranges(story_range: Any, limit: int) -> list[dict[str, Any]]:
    # Word's formatting-only Find does not match when the entire story has one
    # highlight color. Handle that native COM edge case directly first.
    try:
        uniform_color = int(story_range.HighlightColorIndex)
    except Exception:
        uniform_color = _WD_UNDEFINED
    if uniform_color not in {0, _WD_UNDEFINED}:
        return [
            {
                "start_offset": int(story_range.Start),
                "end_offset": int(story_range.End),
                "text": str(story_range.Text).rstrip("\r\x07"),
                "color_index": uniform_color,
                "color": _COLOR_NAMES.get(uniform_color, "unknown"),
            }
        ]

    matches: list[dict[str, Any]] = []
    search_range = story_range.Duplicate
    story_end = int(story_range.End)
    while int(search_range.Start) < story_end:
        find = search_range.Find
        find.ClearFormatting()
        find.Text = ""
        find.Forward = True
        find.Wrap = _WD_FIND_STOP
        find.Format = True
        find.Highlight = True
        if not bool(find.Execute()):
            break
        start = int(search_range.Start)
        end = int(search_range.End)
        if end <= start:
            raise RuntimeError("Word returned an empty highlighted range and could not advance")
        color_index = int(search_range.HighlightColorIndex)
        matches.append(
            {
                "start_offset": start,
                "end_offset": end,
                "text": str(search_range.Text).rstrip("\r\x07"),
                "color_index": color_index,
                "color": _COLOR_NAMES.get(color_index, "mixed_or_unknown"),
            }
        )
        if len(matches) >= limit:
            break
        search_range.SetRange(end, story_end)
    return matches


@word_tool(title="Word Live Inspect Highlighted Text", domain="inspection", change="read")
async def word_live_inspect_highlighted_text(
    filename: str | None = None,
    max_results: int = 500,
) -> dict[str, Any]:
    """Find highlighted ranges across every populated Word story.

    Results include main text, notes, comments, text frames, headers, footers, and
    separator stories when Word exposes them. Offsets are zero-based within each
    story; story instance indexes and result indexes are one-based.
    """
    _require_windows()
    if max_results < 1:
        raise ValueError("max_results must be at least 1")
    from word_mcp_codemode_live.core.word_com import find_document, get_word_app

    document = find_document(get_word_app(), filename)
    stories, skipped_stories = _story_ranges(document)
    results: list[dict[str, Any]] = []
    searched_stories: list[dict[str, Any]] = []
    truncated = False
    for story_type, story_instance_index, story_range in stories:
        searched_stories.append(
            {
                "story": _STORY_NAMES[story_type],
                "story_type_id": story_type,
                "story_instance_index": story_instance_index,
            }
        )
        # Read one beyond the public cap so ``truncated`` means an additional
        # match was actually observed, not merely that the cap was reached.
        remaining_with_sentinel = max_results + 1 - len(results)
        for match in _highlight_ranges(story_range, remaining_with_sentinel):
            results.append(
                {
                    "index": len(results) + 1,
                    "story": _STORY_NAMES[story_type],
                    "story_type_id": story_type,
                    "story_instance_index": story_instance_index,
                    **match,
                }
            )
            if len(results) > max_results:
                truncated = True
                break
        if truncated:
            break

    return {
        "success": True,
        "document": str(document.Name),
        "highlight_count": min(len(results), max_results),
        "truncated": truncated,
        "highlights": results[:max_results],
        "searched_stories": searched_stories,
        "absent_story_types": skipped_stories,
        "limitations": [
            "Offsets are local to each Word story, not global document offsets.",
            "Drawing-layer text is included only when Word exposes it through a text-frame story.",
            "Results reflect the revision text representation exposed by Word's Range.Find.",
        ],
    }
