"""
Style-related functions for Word Document Server.
"""

import logging

from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor

logger = logging.getLogger(__name__)


_NAMED_COLORS = {
    "red": RGBColor(255, 0, 0),
    "blue": RGBColor(0, 0, 255),
    "green": RGBColor(0, 128, 0),
    "yellow": RGBColor(255, 255, 0),
    "black": RGBColor(0, 0, 0),
    "gray": RGBColor(128, 128, 128),
    "white": RGBColor(255, 255, 255),
    "purple": RGBColor(128, 0, 128),
    "orange": RGBColor(255, 165, 0),
}


def _set_font_properties(style, properties) -> None:
    font = style.font
    direct_properties = {"bold": "bold", "italic": "italic", "name": "name"}
    for source, target in direct_properties.items():
        if source in properties:
            setattr(font, target, properties[source])
    if "size" in properties:
        font.size = Pt(properties["size"])
    if "color" in properties:
        font.color.rgb = _docx_color(properties["color"])


def _docx_color(value):
    if isinstance(value, str) and value.lower() in _NAMED_COLORS:
        return _NAMED_COLORS[value.lower()]
    if hasattr(value, "rgb"):
        return value
    if isinstance(value, str):
        try:
            return RGBColor.from_string(value)
        except ValueError:
            return _NAMED_COLORS["black"]
    return value


def _set_paragraph_properties(style, properties) -> None:
    paragraph = style.paragraph_format
    if "alignment" in properties:
        paragraph.alignment = properties["alignment"]
    if "spacing" in properties:
        paragraph.line_spacing = properties["spacing"]


def ensure_heading_style(doc):
    """
    Ensure Heading styles exist in the document.

    Args:
        doc: Document object
    """
    for i in range(1, 10):  # Create Heading 1 through Heading 9
        style_name = f"Heading {i}"
        try:
            # Try to access the style to see if it exists
            style = doc.styles[style_name]
        except KeyError:
            # Create the style if it doesn't exist
            try:
                style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
                if i == 1:
                    style.font.size = Pt(16)
                    style.font.bold = True
                elif i == 2:
                    style.font.size = Pt(14)
                    style.font.bold = True
                else:
                    style.font.size = Pt(12)
                    style.font.bold = True
            except Exception as exc:
                # If style creation fails, we'll just use default formatting
                logger.warning("Could not create %s; using default formatting: %s", style_name, exc)


def ensure_table_style(doc):
    """
    Ensure Table Grid style exists in the document.

    Args:
        doc: Document object
    """
    try:
        # Try to access the style to see if it exists
        doc.styles["Table Grid"]
    except KeyError:
        # If style doesn't exist, we'll handle it at usage time
        pass


def create_style(
    doc, style_name, style_type, base_style=None, font_properties=None, paragraph_properties=None
):
    """
    Create a new style in the document.

    Args:
        doc: Document object
        style_name: Name for the new style
        style_type: Type of style (WD_STYLE_TYPE)
        base_style: Optional base style to inherit from
        font_properties: Dictionary of font properties (bold, italic, size, name, color)
        paragraph_properties: Dictionary of paragraph properties (alignment, spacing)

    Returns:
        The created style
    """
    try:
        # Check if style already exists
        style = doc.styles.get_by_id(style_name, WD_STYLE_TYPE.PARAGRAPH)
        return style
    except Exception:
        # Create new style
        new_style = doc.styles.add_style(style_name, style_type)

        # Set base style if specified
        if base_style:
            new_style.base_style = doc.styles[base_style]

        if font_properties:
            _set_font_properties(new_style, font_properties)

        # Set paragraph properties
        if paragraph_properties:
            _set_paragraph_properties(new_style, paragraph_properties)

        return new_style
