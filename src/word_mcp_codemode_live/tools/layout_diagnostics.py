"""Objective pagination and page-geometry inspection for open Word documents."""

import sys
from typing import Any

from word_mcp_codemode_live.tools.metadata import word_tool

_SECTION_START_NAMES = {
    0: "continuous",
    1: "new_column",
    2: "new_page",
    3: "even_page",
    4: "odd_page",
}
_GUTTER_POSITION_NAMES = {0: "left", 1: "top", 2: "right"}


def _bool_word(value: Any) -> bool:
    """Interpret Word's VBA True (-1) without treating mixed values as true."""
    return value is True or value == -1


def _page_number(word_range: Any) -> int | None:
    try:
        return int(word_range.Information(3))  # wdActiveEndPageNumber
    except Exception:
        return None


def _manual_page_break_offsets(document: Any) -> list[int]:
    """Find manual page breaks using Word offsets, excluding section breaks."""
    search_range = document.Content.Duplicate
    story_end = int(search_range.End)
    offsets: list[int] = []
    while int(search_range.Start) < story_end:
        search_range.Find.ClearFormatting()
        found = search_range.Find.Execute(FindText="^m", Forward=True, Wrap=0)
        if not found:
            break
        start = int(search_range.Start)
        end = int(search_range.End)
        if end <= start:
            raise RuntimeError("Word returned an empty page-break range and could not advance")
        offsets.append(start)
        search_range.SetRange(end, story_end)
    return offsets


@word_tool(title="Word Live Inspect Layout", domain="layout", change="read")
async def word_live_inspect_layout(
    filename: str | None = None,
    max_controlled_paragraphs: int = 200,
) -> dict[str, Any]:
    """Inspect pagination controls and section geometry without changing Word.

    This reports objective state rather than guessing at visual defects: page and
    section counts, usable page geometry, manual page-break offsets, and paragraphs
    carrying pagination controls. ``max_controlled_paragraphs`` limits only the
    detailed paragraph rows, not the counts.
    """
    if sys.platform != "win32":
        raise RuntimeError("Live layout tools are only available on Windows")
    if max_controlled_paragraphs < 1 or max_controlled_paragraphs > 2_000:
        raise ValueError("max_controlled_paragraphs must be between 1 and 2000")

    from word_mcp_codemode_live.core.word_com import find_document, get_word_app

    document = find_document(get_word_app(), filename)
    page_count = int(document.ComputeStatistics(2))  # wdStatisticPages
    sections: list[dict[str, Any]] = []
    findings: list[dict[str, Any]] = []

    for index in range(1, int(document.Sections.Count) + 1):
        section = document.Sections(index)
        section_range = section.Range
        setup = section.PageSetup
        page_width = float(setup.PageWidth)
        page_height = float(setup.PageHeight)
        left = float(setup.LeftMargin)
        right = float(setup.RightMargin)
        top = float(setup.TopMargin)
        bottom = float(setup.BottomMargin)
        gutter = float(setup.Gutter)
        gutter_position_id = int(setup.GutterPos)
        gutter_position = _GUTTER_POSITION_NAMES.get(gutter_position_id, "unknown")
        usable_width = page_width - left - right
        usable_height = page_height - top - bottom
        if gutter_position == "top":
            usable_height -= gutter
        else:
            usable_width -= gutter
        if usable_width <= 0 or usable_height <= 0:
            findings.append(
                {
                    "kind": "non_positive_usable_page_area",
                    "section_index": index,
                    "usable_width_points": usable_width,
                    "usable_height_points": usable_height,
                }
            )
        start_type_id = int(setup.SectionStart)
        section_start_range = section_range.Duplicate
        section_start_range.SetRange(int(section_range.Start), int(section_range.Start))
        sections.append(
            {
                "index": index,
                "start_offset": int(section_range.Start),
                "end_offset": int(section_range.End),
                "start_page": _page_number(section_start_range),
                "start_type": _SECTION_START_NAMES.get(start_type_id, "unknown"),
                "start_type_id": start_type_id,
                "orientation": "landscape" if int(setup.Orientation) == 1 else "portrait",
                "page_width_points": page_width,
                "page_height_points": page_height,
                "margins_points": {
                    "top": top,
                    "bottom": bottom,
                    "left": left,
                    "right": right,
                    "gutter": gutter,
                    "gutter_position": gutter_position,
                    "gutter_position_id": gutter_position_id,
                },
                "usable_width_points": usable_width,
                "usable_height_points": usable_height,
            }
        )

    manual_page_break_offsets = _manual_page_break_offsets(document)

    controlled: list[dict[str, Any]] = []
    controlled_count = 0
    for index in range(1, int(document.Paragraphs.Count) + 1):
        paragraph = document.Paragraphs(index)
        paragraph_range = paragraph.Range
        paragraph_format = paragraph.Format
        controls = {
            "page_break_before": _bool_word(paragraph_format.PageBreakBefore),
            "keep_with_next": _bool_word(paragraph_format.KeepWithNext),
            "keep_together": _bool_word(paragraph_format.KeepTogether),
            "widow_control": _bool_word(paragraph_format.WidowControl),
        }
        paragraph_start = int(paragraph_range.Start)
        paragraph_end = int(paragraph_range.End)
        has_manual_break = any(
            paragraph_start <= offset < paragraph_end for offset in manual_page_break_offsets
        )
        if has_manual_break or any(controls.values()):
            controlled_count += 1
            if len(controlled) < max_controlled_paragraphs:
                controlled.append(
                    {
                        "paragraph_index": index,
                        "start_offset": int(paragraph_range.Start),
                        "end_offset": int(paragraph_range.End),
                        "page": _page_number(paragraph_range),
                        "text": str(paragraph_range.Text).rstrip("\r\x07")[:200],
                        "contains_manual_page_break": has_manual_break,
                        **controls,
                    }
                )

    return {
        "success": True,
        "document": str(document.Name),
        "page_count": page_count,
        "section_count": int(document.Sections.Count),
        "paragraph_count": int(document.Paragraphs.Count),
        "sections": sections,
        "manual_page_break_count": len(manual_page_break_offsets),
        "manual_page_break_offsets": manual_page_break_offsets,
        "controlled_paragraph_count": controlled_count,
        "controlled_paragraphs": controlled,
        "controlled_paragraphs_truncated": controlled_count > len(controlled),
        "findings": findings,
    }
