"""Navigate Word's UI selection without changing document content."""

import sys
from typing import Annotated, Any

from pydantic import Field

from word_mcp_codemode_live.tools.metadata import word_tool

_WD_ACTIVE_END_PAGE_NUMBER = 3
_WD_GO_TO_PAGE = 1
_WD_GO_TO_ABSOLUTE = 1
_WD_NUMBER_OF_PAGES_IN_DOCUMENT = 4
_WD_COLLAPSE_END = 0
_WD_COLLAPSE_START = 1


def _require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Live navigation tools are only available on Windows")


def _page_at(word_range: Any, collapse: int) -> int:
    probe = word_range.Duplicate
    probe.Collapse(collapse)
    return int(probe.Information(_WD_ACTIVE_END_PAGE_NUMBER))


def _selected_end_page(word_range: Any) -> int:
    if int(word_range.End) <= int(word_range.Start):
        return _page_at(word_range, _WD_COLLAPSE_END)
    probe = word_range.Duplicate
    probe.SetRange(int(word_range.End) - 1, int(word_range.End) - 1)
    return int(probe.Information(_WD_ACTIVE_END_PAGE_NUMBER))


@word_tool(title="Word Live Navigate", domain="lifecycle", change="safe_write")
async def word_live_navigate(
    filename: str | None = None,
    page: Annotated[int, Field(ge=1, description="One-based rendered page number.")] | None = None,
    char_start: Annotated[
        int,
        Field(ge=0, description="Zero-based UTF-16 offset in the main document story."),
    ]
    | None = None,
    char_end: Annotated[
        int,
        Field(ge=0, description="End-exclusive UTF-16 main-story offset."),
    ]
    | None = None,
) -> dict[str, Any]:
    """Activate an open document and select a page position or main-story range.

    Use exactly one target mode: ``page`` selects a collapsed position at the
    start of that rendered page; ``char_start`` with optional ``char_end``
    selects a zero-based, end-exclusive range in the document's main story.
    These positions are Word-native UTF-16 code-unit offsets, matching offsets
    returned by the other live inspection tools. ``char_end`` defaults to
    ``char_start`` for a caret position.

    The tool does not alter text, formatting, the document's saved/dirty state,
    or Word's application visibility. When Word is visible it scrolls the target
    into view; in a hidden automation instance it updates the same selection and
    reports that no visible UI was shown.
    """
    _require_windows()
    page_mode = page is not None
    range_mode = char_start is not None or char_end is not None
    if page_mode == range_mode:
        raise ValueError("provide exactly one target: page or char_start/char_end")
    if page_mode and page is not None and page < 1:
        raise ValueError("page must be at least 1")
    if range_mode and char_start is None:
        raise ValueError("char_start is required when char_end is provided")

    from word_mcp_codemode_live.core.word_com import find_document, get_word_app

    application = get_word_app()
    document = find_document(application, filename)
    content_end = int(document.Content.End)
    saved_before = bool(document.Saved)

    if page_mode:
        # Range.Information(4) reports current rendered pages without the
        # saved-state mutation Word's Document.ComputeStatistics(2) can cause.
        total_pages = int(document.Content.Information(_WD_NUMBER_OF_PAGES_IN_DOCUMENT))
        assert page is not None
        if page > total_pages:
            raise ValueError(f"page {page} is out of range; document has {total_pages} pages")
        target = document.GoTo(What=_WD_GO_TO_PAGE, Which=_WD_GO_TO_ABSOLUTE, Count=page)
        target.SetRange(int(target.Start), int(target.Start))
        requested: dict[str, Any] = {"kind": "page", "page": page}
    else:
        assert char_start is not None
        requested_end = char_start if char_end is None else char_end
        if char_start < 0:
            raise ValueError("char_start must be at least 0")
        if requested_end < char_start:
            raise ValueError("char_end must be greater than or equal to char_start")
        if requested_end > content_end:
            raise ValueError(
                f"range end {requested_end} is out of bounds; main story ends at {content_end}"
            )
        target = document.Range(Start=char_start, End=requested_end)
        requested = {
            "kind": "range",
            "char_start": char_start,
            "char_end": requested_end,
        }
        total_pages = int(document.Content.Information(_WD_NUMBER_OF_PAGES_IN_DOCUMENT))

    document.Activate()
    target.Select()
    selected = application.Selection.Range
    actual_start = int(selected.Start)
    actual_end = int(selected.End)
    expected_start = int(target.Start)
    expected_end = int(target.End)
    if (actual_start, actual_end) != (expected_start, expected_end):
        raise RuntimeError(
            "Word did not retain the requested selection: "
            f"expected {expected_start}:{expected_end}, got {actual_start}:{actual_end}"
        )

    # ScrollIntoView operates on window state only. Re-select afterwards so the
    # returned positions describe Word's final UI state, not merely the target.
    document.ActiveWindow.ScrollIntoView(selected, True)
    selected = application.Selection.Range
    final_start = int(selected.Start)
    final_end = int(selected.End)
    if (final_start, final_end) != (expected_start, expected_end):
        raise RuntimeError(
            "Word changed the selection while scrolling it into view: "
            f"expected {expected_start}:{expected_end}, got {final_start}:{final_end}"
        )

    saved_after = bool(document.Saved)
    if saved_after != saved_before:
        raise RuntimeError(
            "Word changed the document saved state during navigation; inspect the document"
        )

    application_visible = bool(application.Visible)
    return {
        "success": True,
        "document": str(document.Name),
        "requested": requested,
        "selection": {
            "char_start": final_start,
            "char_end": final_end,
            "collapsed": final_start == final_end,
            "start_page": _page_at(selected, _WD_COLLAPSE_START),
            "end_page": _selected_end_page(selected),
            "active_end_page": int(application.Selection.Information(_WD_ACTIVE_END_PAGE_NUMBER)),
        },
        "total_pages": total_pages,
        "main_story_end": content_end,
        "saved_before": saved_before,
        "saved_after": saved_after,
        "saved_state_unchanged": True,
        "application_visible": application_visible,
        "ui_effect": (
            "The target document was activated, selected, and scrolled into view."
            if application_visible
            else "The hidden Word selection was updated; application visibility was not changed."
        ),
        "limitations": [
            "Character offsets are Word UTF-16 positions in the main document story only.",
            "Page numbers are Word's current rendered pagination and can change after layout edits.",
        ],
    }
