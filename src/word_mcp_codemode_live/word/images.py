"""Validate and apply native Word image insertion options."""

import os
from typing import Any

from word_mcp_codemode_live.word.values import rgb_hex_to_word

WRAP_STYLES = {
    "inline": None,
    "square": 0,
    "tight": 1,
    "behind": 3,
    "infront": 4,
    "topbottom": 2,
}
_BORDER_STYLES = {"none": 0, "single": 1, "double": 7, "dotted": 3, "dashed": 2, "thick": 6}
_ALIGNMENTS = {"left": 0, "center": 1, "right": 2}


def validate_options(options: dict[str, Any]) -> tuple[str, int]:
    path = os.path.abspath(options["image_path"])
    if not os.path.isfile(path):
        raise ValueError(f"Image file not found: {path}")
    choices = (
        ("wrapping", WRAP_STYLES),
        ("border_style", _BORDER_STYLES),
        ("alignment", _ALIGNMENTS),
    )
    for name, allowed in choices:
        value = options[name]
        if value is not None and value.casefold() not in allowed:
            raise ValueError(f"Unknown {name}: {value}")
    dimensions = ("width_inches", "height_inches", "width_pt", "height_pt", "border_width_pt")
    if any(options[name] is not None and options[name] <= 0 for name in dimensions):
        raise ValueError("Image dimensions and border_width_pt must be positive")
    color = (
        rgb_hex_to_word(options["border_color"], field_name="border_color")
        if options["border_color"]
        else 0
    )
    return path, color


def insertion_range(document: Any, paragraph_index: int | None, position: str) -> Any:
    if paragraph_index is not None:
        if not 1 <= paragraph_index <= document.Paragraphs.Count:
            raise ValueError(
                f"paragraph_index {paragraph_index} out of range (1-{document.Paragraphs.Count})"
            )
        word_range = document.Paragraphs(paragraph_index).Range
        word_range.Collapse(1)
        return word_range
    if position == "start":
        return document.Range(0, 0)
    if position == "end":
        word_range = document.Range()
        word_range.Collapse(0)
        return word_range
    try:
        offset = int(position)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid position: {position}") from exc
    if not 0 <= offset < document.Content.End:
        raise ValueError(f"position must be between 0 and {document.Content.End - 1}")
    return document.Range(offset, offset)


def dimensions(options: dict[str, Any]) -> tuple[float | None, float | None]:
    width = options["width_pt"]
    if width is None and options["width_inches"] is not None:
        width = float(options["width_inches"]) * 72.0
    height = options["height_pt"]
    if height is None and options["height_inches"] is not None:
        height = float(options["height_inches"]) * 72.0
    return width, height


def resize(shape: Any, width: float | None, height: float | None) -> None:
    if width is not None and height is not None:
        shape.Width, shape.Height = width, height
    elif width is not None:
        ratio = shape.Height / shape.Width
        shape.Width, shape.Height = width, width * ratio
    elif height is not None:
        ratio = shape.Width / shape.Height
        shape.Height, shape.Width = height, height * ratio


def apply_floating(shape: Any, document: Any, options: dict[str, Any], color: int) -> None:
    shape.WrapFormat.Type = WRAP_STYLES[options["wrapping"].casefold()]
    _apply_shape_line(shape.Line, options, color)
    alignment = options["alignment"]
    if alignment is None:
        return
    shape.RelativeHorizontalPosition = 0
    shape.RelativeVerticalPosition = 2
    setup = document.PageSetup
    text_width = setup.PageWidth - setup.LeftMargin - setup.RightMargin
    positions = {
        "left": 0,
        "right": max(0, text_width - shape.Width),
        "center": max(0, (text_width - shape.Width) / 2),
    }
    shape.Left = positions[alignment.casefold()]


def _apply_shape_line(line: Any, options: dict[str, Any], color: int) -> None:
    style = options["border_style"]
    if style is None:
        return
    if style.casefold() == "none":
        line.Visible = False
        return
    line.Visible = True
    line.DashStyle = {"dotted": 3, "dashed": 4}.get(style.casefold(), 1)
    line.Weight = float(options["border_width_pt"] or 1.0)
    line.ForeColor.RGB = color
    if style.casefold() == "double":
        line.Style = 3


def apply_inline(shape: Any, options: dict[str, Any], color: int) -> list[str]:
    failures: list[str] = []
    style = options["border_style"]
    if style is not None:
        for border_id in (-1, -2, -3, -4):
            try:
                border = shape.Borders(border_id)
                border.LineStyle = _BORDER_STYLES[style.casefold()]
                if style.casefold() != "none":
                    border.LineWidth = float(options["border_width_pt"] or 1.0)
                    border.Color = color
            except Exception as exc:
                failures.append(f"border {border_id}: {exc}")
    if options["alignment"] is not None:
        shape.Range.ParagraphFormat.Alignment = _ALIGNMENTS[options["alignment"].casefold()]
    return failures
