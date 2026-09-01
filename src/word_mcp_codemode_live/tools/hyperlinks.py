"""Inspect and edit native hyperlinks in open Microsoft Word documents."""

import logging
from typing import Any

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session

logger = logging.getLogger(__name__)


def _as_text(value: Any) -> str:
    """Return a stable string for optional COM string properties."""
    return "" if value is None else str(value)


def _display_range(hyperlink: Any) -> Any:
    """Return the visible result range instead of Word's hidden field-code range."""
    hyperlink_range = hyperlink.Range
    try:
        if hyperlink_range.Fields.Count:
            return hyperlink_range.Fields(1).Result
    except Exception as exc:
        logger.debug("Could not resolve hyperlink field result: %s", exc)
    return hyperlink_range


def _hyperlink_data(hyperlink: Any, index: int) -> dict[str, Any]:
    """Serialize a hyperlink without retaining volatile COM range objects."""
    hyperlink_range = _display_range(hyperlink)
    address = _as_text(hyperlink.Address)
    subaddress = _as_text(hyperlink.SubAddress)
    return {
        "index": index,
        "address": address,
        "subaddress": subaddress,
        "target_kind": (
            "external_and_internal"
            if address and subaddress
            else "external"
            if address
            else "internal"
        ),
        "display_text": _as_text(hyperlink_range.Text),
        "range": {
            "start": int(hyperlink_range.Start),
            "end": int(hyperlink_range.End),
        },
    }


def _get_hyperlink(document: Any, hyperlink_index: int) -> Any:
    if hyperlink_index < 1 or hyperlink_index > document.Hyperlinks.Count:
        raise ValueError(
            f"hyperlink_index {hyperlink_index} out of range (1-{document.Hyperlinks.Count})"
        )
    return document.Hyperlinks(hyperlink_index)


def _find_hyperlink_index(document: Any, hyperlink: Any) -> int:
    """Resolve Word's returned Hyperlink object to its one-based collection index."""
    expected = _hyperlink_data(hyperlink, 0)
    for index in range(1, document.Hyperlinks.Count + 1):
        candidate = _hyperlink_data(document.Hyperlinks(index), 0)
        if candidate == expected:
            return index
    raise RuntimeError("Word added the hyperlink, but it could not be found in the document")


def _resolve_anchor(
    document: Any,
    *,
    start: int | None,
    end: int | None,
    display_text: str | None,
) -> Any:
    if (start is None) != (end is None):
        raise ValueError("start and end must be provided together")

    content_start = int(document.Content.Start)
    content_end = int(document.Content.End) - 1
    if start is None:
        if display_text is None:
            raise ValueError("Provide start/end to link existing text, or provide display_text")
        return document.Range(content_end, content_end)

    assert end is not None
    if start < content_start or end < start or end > content_end:
        raise ValueError(f"Range must satisfy {content_start} <= start <= end <= {content_end}")
    if start == end and display_text is None:
        raise ValueError("display_text is required for a collapsed insertion range")
    return document.Range(start, end)


def _replace_anchor_text(document: Any, anchor: Any, display_text: str) -> Any:
    """Replace exactly the anchor text and return a range covering the replacement."""
    start = int(anchor.Start)
    anchor.Text = display_text
    return document.Range(start, start + len(display_text))


@word_tool(title="Word Live List Hyperlinks", domain="references", change="read")
async def word_live_list_hyperlinks(filename: str | None = None) -> dict[str, Any]:
    """List native hyperlinks in an open Word document.

    Hyperlink indexes are one-based. Range offsets are Word's zero-based character
    positions and the end offset is exclusive.

    Args:
        filename: Open document name or full path (None selects the active document).
    """
    word_session.require_windows("Live hyperlink tools")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    hyperlinks = [
        _hyperlink_data(document.Hyperlinks(index), index)
        for index in range(1, document.Hyperlinks.Count + 1)
    ]
    return {
        "success": True,
        "document": str(document.Name),
        "hyperlink_count": len(hyperlinks),
        "hyperlinks": hyperlinks,
    }


@word_tool(title="Word Live Add Hyperlink", domain="references", change="edit", batchable=True)
async def word_live_add_hyperlink(
    filename: str | None = None,
    address: str = "",
    subaddress: str = "",
    display_text: str | None = None,
    start: int | None = None,
    end: int | None = None,
) -> dict[str, Any]:
    """Add an external or internal native hyperlink to an open document.

    With ``start`` and ``end``, the existing range text is retained unless
    ``display_text`` is supplied. Without a range, ``display_text`` is inserted at
    the end of the document. Range offsets use Word's zero-based, end-exclusive
    character positions.

    Args:
        filename: Open document name or full path (None selects the active document).
        address: External target such as an HTTPS or mailto URL.
        subaddress: Internal target such as a bookmark name or heading location.
        display_text: Optional replacement or newly inserted display text.
        start: Optional zero-based start character position.
        end: Optional zero-based exclusive end character position.
    """
    word_session.require_windows("Live hyperlink tools")
    if not address.strip() and not subaddress.strip():
        raise ValueError("Provide address, subaddress, or both")
    if display_text == "":
        raise ValueError("display_text cannot be empty")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    anchor = _resolve_anchor(
        document,
        start=start,
        end=end,
        display_text=display_text,
    )

    with word_session.undo_record(app, "MCP: Add Hyperlink"):
        if display_text is not None:
            anchor = _replace_anchor_text(document, anchor, display_text)
        hyperlink = document.Hyperlinks.Add(
            Anchor=anchor,
            Address=address,
            SubAddress=subaddress,
        )

    index = _find_hyperlink_index(document, hyperlink)
    return {
        "success": True,
        "document": str(document.Name),
        "hyperlink": _hyperlink_data(hyperlink, index),
    }


@word_tool(
    title="Word Live Update Hyperlink",
    domain="references",
    change="edit",
    batchable=True,
)
async def word_live_update_hyperlink(
    filename: str | None = None,
    hyperlink_index: int = 1,
    address: str | None = None,
    subaddress: str | None = None,
    display_text: str | None = None,
) -> dict[str, Any]:
    """Update one native hyperlink selected by its one-based index.

    ``None`` leaves a property unchanged. An empty address or subaddress clears that
    target component, provided the resulting hyperlink still has a target.

    Args:
        filename: Open document name or full path (None selects the active document).
        hyperlink_index: One-based index returned by ``word_live_list_hyperlinks``.
        address: New external target, an empty string to clear it, or None to retain it.
        subaddress: New internal target, an empty string to clear it, or None to retain it.
        display_text: New visible text or None to retain it.
    """
    word_session.require_windows("Live hyperlink tools")
    if address is None and subaddress is None and display_text is None:
        raise ValueError("Provide at least one property to update")
    if display_text == "":
        raise ValueError("display_text cannot be empty")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    hyperlink = _get_hyperlink(document, hyperlink_index)
    resulting_address = _as_text(hyperlink.Address) if address is None else address
    resulting_subaddress = _as_text(hyperlink.SubAddress) if subaddress is None else subaddress
    if not resulting_address.strip() and not resulting_subaddress.strip():
        raise ValueError("The updated hyperlink must retain an address or subaddress")

    # Word ignores assigning an empty string to some Hyperlink target properties.
    # Rebuild in place when clearing either component so conversion between external
    # and internal targets is reliable and the visible text remains untouched.
    rebuild = (address == "" and _as_text(hyperlink.Address)) or (
        subaddress == "" and _as_text(hyperlink.SubAddress)
    )
    with word_session.undo_record(app, "MCP: Update Hyperlink"):
        if rebuild:
            # A duplicate tracks Word's range adjustment when deleting the hidden
            # HYPERLINK field code; numeric endpoints captured before Delete do not.
            anchor = _display_range(hyperlink).Duplicate
            hyperlink.Delete()
            if display_text is not None:
                anchor = _replace_anchor_text(document, anchor, display_text)
            hyperlink = document.Hyperlinks.Add(
                Anchor=anchor,
                Address=resulting_address,
                SubAddress=resulting_subaddress,
            )
            hyperlink_index = _find_hyperlink_index(document, hyperlink)
        else:
            if address is not None:
                hyperlink.Address = address
            if subaddress is not None:
                hyperlink.SubAddress = subaddress
            if display_text is not None:
                hyperlink.TextToDisplay = display_text

    # Text changes can replace the underlying Word range; reacquire it for stable output.
    hyperlink = _get_hyperlink(document, hyperlink_index)
    return {
        "success": True,
        "document": str(document.Name),
        "hyperlink": _hyperlink_data(hyperlink, hyperlink_index),
    }


@word_tool(
    title="Word Live Remove Hyperlink",
    domain="references",
    change="edit",
    batchable=True,
)
async def word_live_remove_hyperlink(
    filename: str | None = None,
    hyperlink_index: int = 1,
) -> dict[str, Any]:
    """Remove one native hyperlink while preserving its visible document text.

    Args:
        filename: Open document name or full path (None selects the active document).
        hyperlink_index: One-based index returned by ``word_live_list_hyperlinks``.
    """
    word_session.require_windows("Live hyperlink tools")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    hyperlink = _get_hyperlink(document, hyperlink_index)
    removed = _hyperlink_data(hyperlink, hyperlink_index)

    with word_session.undo_record(app, "MCP: Remove Hyperlink"):
        hyperlink.Delete()

    return {
        "success": True,
        "document": str(document.Name),
        "removed_hyperlink": removed,
        "remaining_hyperlinks": int(document.Hyperlinks.Count),
    }
