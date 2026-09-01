"""Create, inspect, update, and delete custom styles in open Word documents."""

import sys
from typing import Any, Literal

from word_mcp_codemode_live.tools.metadata import word_tool

StyleKind = Literal["paragraph", "character"]

_STYLE_TYPES = {"paragraph": 1, "character": 2}
_STYLE_TYPE_NAMES = {
    1: "paragraph",
    2: "character",
    3: "table",
    4: "list",
    5: "paragraph_only",
    6: "linked",
}
_ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "justify": 3, "distribute": 4}
_ALIGNMENT_NAMES = {value: name for name, value in _ALIGNMENTS.items()}
_MIXED = 9999999


def _require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Live style tools are only available on Windows")


def _style_by_name(document: Any, style_name: str) -> Any:
    try:
        return document.Styles(style_name)
    except Exception as exc:
        raise ValueError(f"Word style not found: {style_name}") from exc


def _optional_value(value: Any) -> Any:
    return None if value == _MIXED else value


def _style_name(style_or_name: Any) -> str | None:
    if style_or_name is None:
        return None
    try:
        return str(style_or_name.NameLocal)
    except Exception:
        text = str(style_or_name)
        return text if text else None


def _style_entry(style: Any, index: int | None = None) -> dict[str, Any]:
    font = style.Font
    style_type = int(style.Type)
    entry: dict[str, Any] = {
        "name": str(style.NameLocal),
        "built_in": bool(style.BuiltIn),
        "type": _STYLE_TYPE_NAMES.get(style_type, "unknown"),
        "type_id": style_type,
        "base_style": _style_name(style.BaseStyle),
        "automatically_update": bool(style.AutomaticallyUpdate) if style_type in {1, 5} else None,
        "font": {
            "name": _optional_value(font.Name),
            "size_pt": _optional_value(font.Size),
            "bold": None if font.Bold == _MIXED else bool(font.Bold),
            "italic": None if font.Italic == _MIXED else bool(font.Italic),
            "underline": _optional_value(font.Underline),
            "color_bgr": _optional_value(font.Color),
        },
    }
    if index is not None:
        entry["index"] = index
    if style_type in {1, 5}:
        paragraph_format = style.ParagraphFormat
        entry["paragraph"] = {
            "alignment": _ALIGNMENT_NAMES.get(
                int(paragraph_format.Alignment), str(paragraph_format.Alignment)
            ),
            "space_before_pt": _optional_value(paragraph_format.SpaceBefore),
            "space_after_pt": _optional_value(paragraph_format.SpaceAfter),
            "left_indent_pt": _optional_value(paragraph_format.LeftIndent),
            "right_indent_pt": _optional_value(paragraph_format.RightIndent),
            "first_line_indent_pt": _optional_value(paragraph_format.FirstLineIndent),
            "keep_with_next": None
            if paragraph_format.KeepWithNext == _MIXED
            else bool(paragraph_format.KeepWithNext),
            "keep_together": None
            if paragraph_format.KeepTogether == _MIXED
            else bool(paragraph_format.KeepTogether),
            "page_break_before": None
            if paragraph_format.PageBreakBefore == _MIXED
            else bool(paragraph_format.PageBreakBefore),
            "outline_level": _optional_value(paragraph_format.OutlineLevel),
        }
    return entry


def _color_value(color: str) -> int:
    from word_mcp_codemode_live.core.word_values import rgb_hex_to_word

    return rgb_hex_to_word(color, field_name="font_color")


def _set_optional_properties(target: Any, properties: tuple[tuple[str, Any], ...]) -> None:
    for name, value in properties:
        if value is not None:
            setattr(target, name, value)


def _apply_style_properties(
    style: Any,
    *,
    base_style: Any | None,
    automatically_update: bool | None,
    font_name: str | None,
    font_size: float | None,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    font_color_value: int | None,
    alignment_value: int | None,
    space_before: float | None,
    space_after: float | None,
    left_indent: float | None,
    right_indent: float | None,
    first_line_indent: float | None,
    keep_with_next: bool | None,
    keep_together: bool | None,
    page_break_before: bool | None,
    outline_level: int | None,
) -> None:
    _set_optional_properties(
        style,
        (("BaseStyle", base_style), ("AutomaticallyUpdate", automatically_update)),
    )
    _set_optional_properties(
        style.Font,
        (
            ("Name", font_name),
            ("Size", font_size),
            ("Bold", bold),
            ("Italic", italic),
            ("Underline", None if underline is None else (1 if underline else 0)),
            ("Color", font_color_value),
        ),
    )

    paragraph_values = (
        alignment_value,
        space_before,
        space_after,
        left_indent,
        right_indent,
        first_line_indent,
        keep_with_next,
        keep_together,
        page_break_before,
        outline_level,
    )
    if any(value is not None for value in paragraph_values):
        if int(style.Type) not in {1, 5}:
            raise ValueError("paragraph formatting options require a paragraph style")
        paragraph = style.ParagraphFormat
        _set_optional_properties(
            paragraph,
            (
                ("Alignment", alignment_value),
                ("SpaceBefore", space_before),
                ("SpaceAfter", space_after),
                ("LeftIndent", left_indent),
                ("RightIndent", right_indent),
                ("FirstLineIndent", first_line_indent),
                ("KeepWithNext", keep_with_next),
                ("KeepTogether", keep_together),
                ("PageBreakBefore", page_break_before),
                ("OutlineLevel", outline_level),
            ),
        )


def _validate_properties(
    *,
    font_size: float | None,
    alignment: str | None,
    spacing_values: tuple[float | None, ...],
    outline_level: int | None,
) -> int | None:
    if font_size is not None and font_size <= 0:
        raise ValueError("font_size must be greater than zero")
    if any(value is not None and value < 0 for value in spacing_values[:2]):
        raise ValueError("space_before and space_after cannot be negative")
    if outline_level is not None and not 1 <= outline_level <= 10:
        raise ValueError("outline_level must be between 1 and 10 (10 means body text)")
    if alignment is None:
        return None
    try:
        return _ALIGNMENTS[alignment.casefold()]
    except KeyError as exc:
        raise ValueError(f"alignment must be one of {sorted(_ALIGNMENTS)}") from exc


def _uses_paragraph_properties(
    alignment: str | None,
    space_before: float | None,
    space_after: float | None,
    left_indent: float | None,
    right_indent: float | None,
    first_line_indent: float | None,
    keep_with_next: bool | None,
    keep_together: bool | None,
    page_break_before: bool | None,
    outline_level: int | None,
) -> bool:
    return any(
        value is not None
        for value in (
            alignment,
            space_before,
            space_after,
            left_indent,
            right_indent,
            first_line_indent,
            keep_with_next,
            keep_together,
            page_break_before,
            outline_level,
        )
    )


@word_tool(title="Word Live List Custom Styles", domain="styles", change="read")
async def word_live_list_custom_styles(filename: str | None = None) -> dict[str, Any]:
    """List user-defined Word styles. Returned indexes are one-based."""
    _require_windows()
    from word_mcp_codemode_live.core.word_com import find_document, get_word_app

    document = find_document(get_word_app(), filename)
    styles = []
    for collection_index in range(1, int(document.Styles.Count) + 1):
        style = document.Styles(collection_index)
        if not bool(style.BuiltIn):
            styles.append(_style_entry(style, len(styles) + 1))
            styles[-1]["collection_index"] = collection_index
    return {
        "success": True,
        "document": str(document.Name),
        "custom_style_count": len(styles),
        "styles": styles,
    }


@word_tool(title="Word Live Create Custom Style", domain="styles", change="edit", batchable=True)
async def word_live_create_custom_style(
    filename: str | None = None,
    style_name: str = "",
    style_type: StyleKind = "paragraph",
    base_style: str | None = None,
    automatically_update: bool | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    font_color: str | None = None,
    alignment: str | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    left_indent: float | None = None,
    right_indent: float | None = None,
    first_line_indent: float | None = None,
    keep_with_next: bool | None = None,
    keep_together: bool | None = None,
    page_break_before: bool | None = None,
    outline_level: int | None = None,
) -> dict[str, Any]:
    """Create a custom paragraph or character style in an open Word document."""
    _require_windows()
    style_name = style_name.strip()
    if not style_name:
        raise ValueError("style_name is required")
    if style_type not in _STYLE_TYPES:
        raise ValueError("style_type must be paragraph or character")
    if style_type == "character" and _uses_paragraph_properties(
        alignment,
        space_before,
        space_after,
        left_indent,
        right_indent,
        first_line_indent,
        keep_with_next,
        keep_together,
        page_break_before,
        outline_level,
    ):
        raise ValueError("paragraph formatting options require a paragraph style")
    if style_type == "character" and automatically_update is not None:
        raise ValueError("automatically_update is not supported for character styles")
    alignment_value = _validate_properties(
        font_size=font_size,
        alignment=alignment,
        spacing_values=(space_before, space_after, left_indent, right_indent, first_line_indent),
        outline_level=outline_level,
    )
    color_value = _color_value(font_color) if font_color is not None else None

    from word_mcp_codemode_live.core.word_com import (
        find_document,
        get_word_app,
        undo_transaction,
    )

    app = get_word_app()
    document = find_document(app, filename)
    style_exists = True
    try:
        document.Styles(style_name)
    except Exception:
        style_exists = False
    if style_exists:
        raise ValueError(f"A Word style named {style_name!r} already exists")
    base = _style_by_name(document, base_style) if base_style is not None else None

    with undo_transaction(app, document, "MCP: Create Custom Style"):
        style = document.Styles.Add(Name=style_name, Type=_STYLE_TYPES[style_type])
        _apply_style_properties(
            style,
            base_style=base,
            automatically_update=automatically_update,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            font_color_value=color_value,
            alignment_value=alignment_value,
            space_before=space_before,
            space_after=space_after,
            left_indent=left_indent,
            right_indent=right_indent,
            first_line_indent=first_line_indent,
            keep_with_next=keep_with_next,
            keep_together=keep_together,
            page_break_before=page_break_before,
            outline_level=outline_level,
        )
        created = _style_entry(style)
    return {"success": True, "document": str(document.Name), "style": created}


@word_tool(title="Word Live Update Custom Style", domain="styles", change="edit", batchable=True)
async def word_live_update_custom_style(
    filename: str | None = None,
    style_name: str = "",
    base_style: str | None = None,
    automatically_update: bool | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    font_color: str | None = None,
    alignment: str | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    left_indent: float | None = None,
    right_indent: float | None = None,
    first_line_indent: float | None = None,
    keep_with_next: bool | None = None,
    keep_together: bool | None = None,
    page_break_before: bool | None = None,
    outline_level: int | None = None,
) -> dict[str, Any]:
    """Update formatting or inheritance of an existing custom Word style."""
    _require_windows()
    style_name = style_name.strip()
    if not style_name:
        raise ValueError("style_name is required")
    alignment_value = _validate_properties(
        font_size=font_size,
        alignment=alignment,
        spacing_values=(space_before, space_after, left_indent, right_indent, first_line_indent),
        outline_level=outline_level,
    )
    color_value = _color_value(font_color) if font_color is not None else None

    from word_mcp_codemode_live.core.word_com import (
        find_document,
        get_word_app,
        undo_transaction,
    )

    app = get_word_app()
    document = find_document(app, filename)
    style = _style_by_name(document, style_name)
    if bool(style.BuiltIn):
        raise ValueError("Only custom styles can be updated by this tool")
    if int(style.Type) not in {1, 5} and _uses_paragraph_properties(
        alignment,
        space_before,
        space_after,
        left_indent,
        right_indent,
        first_line_indent,
        keep_with_next,
        keep_together,
        page_break_before,
        outline_level,
    ):
        raise ValueError("paragraph formatting options require a paragraph style")
    if int(style.Type) == 2 and automatically_update is not None:
        raise ValueError("automatically_update is not supported for character styles")
    base = _style_by_name(document, base_style) if base_style is not None else None
    with undo_transaction(app, document, "MCP: Update Custom Style"):
        _apply_style_properties(
            style,
            base_style=base,
            automatically_update=automatically_update,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            font_color_value=color_value,
            alignment_value=alignment_value,
            space_before=space_before,
            space_after=space_after,
            left_indent=left_indent,
            right_indent=right_indent,
            first_line_indent=first_line_indent,
            keep_with_next=keep_with_next,
            keep_together=keep_together,
            page_break_before=page_break_before,
            outline_level=outline_level,
        )
        updated = _style_entry(style)
    return {"success": True, "document": str(document.Name), "style": updated}


@word_tool(title="Word Live Delete Custom Style", domain="styles", change="edit", batchable=True)
async def word_live_delete_custom_style(
    filename: str | None = None,
    style_name: str = "",
) -> dict[str, Any]:
    """Delete one user-defined Word style; built-in styles are rejected."""
    _require_windows()
    style_name = style_name.strip()
    if not style_name:
        raise ValueError("style_name is required")
    from word_mcp_codemode_live.core.word_com import (
        find_document,
        get_word_app,
        undo_transaction,
    )

    app = get_word_app()
    document = find_document(app, filename)
    style = _style_by_name(document, style_name)
    if bool(style.BuiltIn):
        raise ValueError("Built-in Word styles cannot be deleted by this tool")
    deleted = _style_entry(style)
    with undo_transaction(app, document, "MCP: Delete Custom Style"):
        style.Delete()
    return {
        "success": True,
        "document": str(document.Name),
        "deleted_style": deleted,
    }
