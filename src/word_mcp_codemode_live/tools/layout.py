"""COM-based layout tools for Microsoft Word.

These tools operate on documents currently open in Word via COM automation.
They provide layout, spacing, bookmark, watermark, and section management for
files that are open (and locked) in Word. Header/footer editing lives in the
dedicated ``headers_footers`` module.
"""

from typing import Annotated, Any, Literal

from pydantic import Field

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word.values import rgb_hex_to_word

# 1 inch = 72 points (avoid app.InchesToPoints which can fail on some COM setups)
_PTS_PER_INCH = 72.0
_LINE_SPACING_RULES = {
    "single": 0,
    "1.5_lines": 1,
    "double": 2,
    "at_least": 3,
    "exactly": 4,
    "multiple": 5,
}
_PARAGRAPH_ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "justify": 3}


def _apply_page_setup(page_setup, values: dict[str, Any]) -> list[str]:
    changes: list[str] = []
    orientation = values["orientation"]
    if orientation is not None:
        page_setup.Orientation = 1 if str(orientation).casefold() == "landscape" else 0
        changes.append(f"orientation={str(orientation).casefold()}")
    assignments = (
        ("page_width_inches", "PageWidth", "width"),
        ("page_height_inches", "PageHeight", "height"),
        ("margin_top_inches", "TopMargin", "margin_top"),
        ("margin_bottom_inches", "BottomMargin", "margin_bottom"),
        ("margin_left_inches", "LeftMargin", "margin_left"),
        ("margin_right_inches", "RightMargin", "margin_right"),
    )
    for name, attribute, label in assignments:
        value = values[name]
        if value is not None:
            setattr(page_setup, attribute, float(value) * _PTS_PER_INCH)
            changes.append(f"{label}={value}in")
    return changes


def _paragraph_indices(document, paragraph_index, start_paragraph, end_paragraph):
    total = document.Paragraphs.Count
    if start_paragraph is not None and end_paragraph is not None:
        if not 1 <= start_paragraph <= end_paragraph <= total:
            raise ValueError(f"Paragraph range must be within 1-{total} and ordered")
        return range(start_paragraph, end_paragraph + 1)
    if paragraph_index is not None:
        if not 1 <= paragraph_index <= total:
            raise ValueError(f"paragraph_index {paragraph_index} out of range (1-{total})")
        return (paragraph_index,)
    return range(1, total + 1)


def _apply_paragraph_spacing(paragraph_format, values: dict[str, Any]) -> None:
    assignments = (
        ("SpaceBefore", values["space_before_pt"]),
        ("SpaceAfter", values["space_after_pt"]),
        ("LineSpacing", values["line_spacing"]),
        ("KeepWithNext", values["keep_with_next"]),
        ("KeepTogether", values["keep_together"]),
    )
    for attribute, value in assignments:
        if value is not None:
            setattr(paragraph_format, attribute, value)
    if values["line_spacing_rule"] is not None:
        paragraph_format.LineSpacingRule = _LINE_SPACING_RULES[str(values["line_spacing_rule"])]
    if values["alignment"] is not None:
        paragraph_format.Alignment = _PARAGRAPH_ALIGNMENTS[str(values["alignment"])]


@word_tool(title="Word Live Set Page Layout", domain="layout", change="edit", batchable=True)
async def word_live_set_page_layout(
    filename: str | None = None,
    section_index: Annotated[int, Field(ge=1)] = 1,
    orientation: Literal["portrait", "landscape"] | None = None,
    page_width_inches: float | None = None,
    page_height_inches: float | None = None,
    margin_top_inches: float | None = None,
    margin_bottom_inches: float | None = None,
    margin_left_inches: float | None = None,
    margin_right_inches: float | None = None,
) -> dict[str, Any]:
    """Set page layout for a section in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        section_index: Section number (1-indexed, COM style). Default 1.
        orientation: "portrait" or "landscape".
        page_width_inches: Page width in inches.
        page_height_inches: Page height in inches.
        margin_top_inches: Top margin in inches.
        margin_bottom_inches: Bottom margin in inches.
        margin_left_inches: Left margin in inches.
        margin_right_inches: Right margin in inches.

    Returns:
        JSON with result info.
    """

    word_session.require_windows("Live layout tools")

    if orientation is not None and orientation.casefold() not in {"portrait", "landscape"}:
        raise ValueError("orientation must be portrait or landscape")
    dimensions = {
        "page_width_inches": page_width_inches,
        "page_height_inches": page_height_inches,
    }
    if any(value is not None and value <= 0 for value in dimensions.values()):
        raise ValueError("Page width and height must be greater than zero")
    margins = {
        "margin_top_inches": margin_top_inches,
        "margin_bottom_inches": margin_bottom_inches,
        "margin_left_inches": margin_left_inches,
        "margin_right_inches": margin_right_inches,
    }
    if any(value is not None and value < 0 for value in margins.values()):
        raise ValueError("Margins cannot be negative")

    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    if section_index < 1 or section_index > doc.Sections.Count:
        raise ValueError(f"Section {section_index} out of range (1-{doc.Sections.Count})")

    with word_session.undo_record(app, "MCP: Set Page Layout"):
        changes = _apply_page_setup(
            doc.Sections(section_index).PageSetup,
            {
                "orientation": orientation,
                "page_width_inches": page_width_inches,
                "page_height_inches": page_height_inches,
                **margins,
            },
        )

    return {
        "success": True,
        "document": str(doc.Name),
        "section": section_index,
        "changes": changes,
    }


@word_tool(title="Word Live Add Section Break", domain="layout", change="edit", batchable=True)
async def word_live_add_section_break(
    filename: str | None = None,
    break_type: Literal["new_page", "continuous", "even_page", "odd_page"] = "new_page",
    position: Annotated[int, Field(ge=0)] | None = None,
    paragraph_index: Annotated[int, Field(ge=1)] | None = None,
) -> dict[str, Any]:
    """Add a section break to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        break_type: "new_page", "continuous", "even_page", "odd_page".
        position: Optional character offset. Defaults to the document end.
        paragraph_index: Optional one-based paragraph; inserts before it.

    Returns:
        JSON with result info.
    """

    word_session.require_windows("Live layout tools")
    if position is not None and paragraph_index is not None:
        raise ValueError("Provide position or paragraph_index, not both")

    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    # wdSectionBreakNextPage=2, Continuous=3, EvenPage=4, OddPage=5
    type_map = {"new_page": 2, "continuous": 3, "even_page": 4, "odd_page": 5}

    if break_type not in type_map:
        raise ValueError(f"Invalid break_type: {break_type}. Use: {list(type_map.keys())}")

    if paragraph_index is not None:
        if not 1 <= paragraph_index <= doc.Paragraphs.Count:
            raise ValueError(f"paragraph_index must be between 1 and {doc.Paragraphs.Count}")
        insertion_position = int(doc.Paragraphs(paragraph_index).Range.Start)
    elif position is not None:
        if not 0 <= position < doc.Content.End:
            raise ValueError(f"position must be between 0 and {doc.Content.End - 1}")
        insertion_position = position
    else:
        insertion_position = int(doc.Content.End - 1)

    with word_session.undo_record(app, "MCP: Add Section Break"):
        rng = doc.Range(insertion_position, insertion_position)
        rng.InsertBreak(Type=type_map[break_type])

    return {
        "success": True,
        "document": str(doc.Name),
        "break_type": break_type,
        "position": insertion_position,
        "total_sections": int(doc.Sections.Count),
    }


@word_tool(
    title="Word Live Set Paragraph Spacing",
    domain="formatting",
    change="edit",
    batchable=True,
)
async def word_live_set_paragraph_spacing(
    filename: str | None = None,
    paragraph_index: Annotated[int, Field(ge=1)] | None = None,
    start_paragraph: Annotated[int, Field(ge=1)] | None = None,
    end_paragraph: Annotated[int, Field(ge=1)] | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    line_spacing: float | None = None,
    line_spacing_rule: Literal["single", "1.5_lines", "double", "at_least", "exactly", "multiple"]
    | None = None,
    keep_with_next: bool | None = None,
    keep_together: bool | None = None,
    alignment: Literal["left", "center", "right", "justify"] | None = None,
) -> dict[str, Any]:
    """Set paragraph spacing and layout properties in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Single paragraph (1-indexed). Ignored if start/end given.
        start_paragraph: Start of range (1-indexed, inclusive).
        end_paragraph: End of range (1-indexed, inclusive).
        space_before_pt: Space before paragraph in points.
        space_after_pt: Space after paragraph in points.
        line_spacing: Line spacing value IN POINTS (depends on rule).
            IMPORTANT: For "multiple" rule, value is in points, NOT a multiplier.
            Single spacing (1.0) = 12pt. So: 1.15 lines = 13.8pt, 1.5 lines = 18pt,
            2.0 lines = 24pt. Formula: desired_lines * 12 = points_value.
        line_spacing_rule: "single"(0), "1.5_lines"(1), "double"(2),
                           "at_least"(3), "exactly"(4), "multiple"(5).
        keep_with_next: Keep paragraph with next paragraph on same page (True/False).
        keep_together: Keep all lines of paragraph on same page (True/False).
        alignment: Paragraph alignment - "left"(0), "center"(1), "right"(2), "justify"(3).

    Returns:
        JSON with count of affected paragraphs.
    """

    word_session.require_windows("Live layout tools")

    if line_spacing_rule is not None and line_spacing_rule not in _LINE_SPACING_RULES:
        raise ValueError(f"Invalid line_spacing_rule: {line_spacing_rule}")
    if alignment is not None and alignment not in _PARAGRAPH_ALIGNMENTS:
        raise ValueError(f"Invalid alignment: {alignment}")
    if (start_paragraph is None) != (end_paragraph is None):
        raise ValueError("start_paragraph and end_paragraph must be provided together")
    if any(value is not None and value < 0 for value in (space_before_pt, space_after_pt)):
        raise ValueError("Paragraph spacing cannot be negative")
    if line_spacing is not None and line_spacing <= 0:
        raise ValueError("line_spacing must be greater than zero")

    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    indices = _paragraph_indices(doc, paragraph_index, start_paragraph, end_paragraph)
    values = {
        "space_before_pt": space_before_pt,
        "space_after_pt": space_after_pt,
        "line_spacing": line_spacing,
        "line_spacing_rule": line_spacing_rule,
        "keep_with_next": keep_with_next,
        "keep_together": keep_together,
        "alignment": alignment,
    }

    with word_session.undo_record(app, "MCP: Set Paragraph Spacing"):
        count = 0
        for index in indices:
            _apply_paragraph_spacing(doc.Paragraphs(index).Format, values)
            count += 1

    return {
        "success": True,
        "document": str(doc.Name),
        "paragraphs_affected": count,
    }


@word_tool(title="Word Live Add Bookmark", domain="layout", change="edit", batchable=True)
async def word_live_add_bookmark(
    filename: str | None = None,
    paragraph_index: Annotated[int, Field(ge=1)] = 1,
    bookmark_name: str = "",
) -> dict[str, Any]:
    """Add a named bookmark at a paragraph in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        paragraph_index: Paragraph to bookmark (1-indexed).
        bookmark_name: Bookmark name (alphanumeric + underscore, no spaces).

    Returns:
        JSON with result info.
    """

    word_session.require_windows("Live layout tools")

    if not bookmark_name:
        raise ValueError("bookmark_name is required")

    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
        raise ValueError(
            f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
        )

    with word_session.undo_record(app, "MCP: Add Bookmark"):
        rng = doc.Paragraphs(paragraph_index).Range
        doc.Bookmarks.Add(bookmark_name, rng)

    return {
        "success": True,
        "document": str(doc.Name),
        "bookmark_name": bookmark_name,
        "paragraph_index": paragraph_index,
    }


@word_tool(title="Word Live Add Watermark", domain="layout", change="edit", batchable=True)
async def word_live_add_watermark(
    filename: str | None = None,
    text: str = "TASLAK",
    font_size: Annotated[int, Field(gt=0)] = 72,
    font_color: str = "C0C0C0",
    rotation: int = -45,
    section_index: Annotated[int, Field(ge=1)] = 1,
) -> dict[str, Any]:
    """Add a diagonal text watermark to an open Word document.

    Args:
        filename: Document name or path (None = active document).
        text: Watermark text (e.g. "DRAFT" or "CONFIDENTIAL").
        font_size: Font size in points.
        font_color: Hex color without # (e.g. "C0C0C0").
        rotation: Rotation angle in degrees (e.g. -45).
        section_index: Section number (1-indexed).

    Returns:
        JSON with result info.
    """

    word_session.require_windows("Live layout tools")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    if section_index < 1 or section_index > doc.Sections.Count:
        raise ValueError(f"Section {section_index} out of range (1-{doc.Sections.Count})")

    with word_session.undo_record(app, "MCP: Add Watermark"):
        section = doc.Sections(section_index)
        header = section.Headers(1)  # wdHeaderFooterPrimary

        # Parse color
        rgb_color = rgb_hex_to_word(font_color, field_name="font_color")

        # AddTextEffect(PresetTextEffect, Text, FontName, FontSize,
        #               FontBold, FontItalic, Left, Top)
        # COM requires positional args
        shape = header.Shapes.AddTextEffect(0, text, "Calibri", font_size, False, False, 0, 0)

        # Configure shape
        shape.Fill.ForeColor.RGB = rgb_color
        shape.Fill.Transparency = 0.5
        shape.Line.Visible = False  # msoFalse
        shape.Rotation = rotation
        shape.LockAspectRatio = False

        # Position relative to page center
        # msoRelativeHorizontalPositionMargin = 0
        # msoRelativeVerticalPositionMargin = 0
        shape.RelativeHorizontalPosition = 0
        shape.RelativeVerticalPosition = 0
        shape.Left = -999995  # wdShapeCenter
        shape.Top = -999995

        # Send behind text
        shape.WrapFormat.Type = 3  # wdWrapBehind
        shape.WrapFormat.AllowOverlap = True

    return {
        "success": True,
        "document": str(doc.Name),
        "text": text,
        "font_size": font_size,
        "color": font_color,
        "rotation": rotation,
        "section": section_index,
    }
