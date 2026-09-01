"""Inspect and edit formatting in live Word documents."""

from typing import Annotated, Any, Literal

from pydantic import BaseModel, Field

from word_mcp_codemode_live.defaults import DEFAULT_AUTHOR
from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word.ranges import character_or_paragraph_range
from word_mcp_codemode_live.word.values import rgb_hex_to_word

_ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "justify": 3}
_ALIGNMENT_NAMES = {0: "left", 1: "center", 2: "right", 3: "justify", 4: "distribute"}
_SPACING_RULE_NAMES = {
    0: "single",
    1: "1.5_lines",
    2: "double",
    3: "at_least",
    4: "exactly",
    5: "multiple",
}


class ParagraphFormatResult(BaseModel):
    """Structured result for paragraph-format inspection."""

    success: Literal[True] = True
    document: str
    paragraphs: list[dict[str, Any]]


class FormatTextResult(BaseModel):
    """Structured result for a formatting edit."""

    success: Literal[True] = True
    document: str
    range: str
    text_preview: str
    tracked: bool


def _word_runs(word_range: Any) -> list[dict[str, Any]]:
    runs: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for index in range(1, word_range.Words.Count + 1):
        word = word_range.Words(index)
        values = {
            "bold": bool(word.Font.Bold) if word.Font.Bold != 9999999 else "mixed",
            "italic": bool(word.Font.Italic) if word.Font.Italic != 9999999 else "mixed",
            "font_name": str(word.Font.Name) if word.Font.Name != 9999999 else "mixed",
            "font_size": word.Font.Size if word.Font.Size != 9999999 else "mixed",
        }
        if current is not None and current["_format"] == values:
            current["text"] += word.Text
            continue
        if current is not None:
            current.pop("_format")
            runs.append(current)
        current = {"text": word.Text, **values, "_format": values}
    if current is not None:
        current.pop("_format")
        runs.append(current)
    return runs


def _paragraph_format_info(paragraph: Any, index: int, include_runs: bool) -> dict[str, Any]:
    word_range = paragraph.Range
    paragraph_format = paragraph.Format
    text = word_range.Text.rstrip("\r\x07")
    info = {
        "index": index,
        "text_preview": text[:80] + ("..." if len(text) > 80 else ""),
        "char_start": word_range.Start,
        "char_end": word_range.End,
        "style": str(word_range.Style) if word_range.Style else "",
        "font_name": str(word_range.Font.Name) if word_range.Font.Name else "",
        "font_size": word_range.Font.Size if word_range.Font.Size else None,
        "bold": bool(word_range.Font.Bold) if word_range.Font.Bold != 9999999 else "mixed",
        "italic": bool(word_range.Font.Italic) if word_range.Font.Italic != 9999999 else "mixed",
        "alignment": _ALIGNMENT_NAMES.get(
            paragraph_format.Alignment, str(paragraph_format.Alignment)
        ),
        "space_before_pt": paragraph_format.SpaceBefore,
        "space_after_pt": paragraph_format.SpaceAfter,
        "line_spacing": paragraph_format.LineSpacing,
        "line_spacing_rule": _SPACING_RULE_NAMES.get(
            paragraph_format.LineSpacingRule, str(paragraph_format.LineSpacingRule)
        ),
        "left_indent_pt": paragraph_format.LeftIndent,
        "right_indent_pt": paragraph_format.RightIndent,
        "first_line_indent_pt": paragraph_format.FirstLineIndent,
        "page_break_before": bool(paragraph_format.PageBreakBefore),
        "keep_with_next": bool(paragraph_format.KeepWithNext),
        "keep_together": bool(paragraph_format.KeepTogether),
    }
    _add_optional_paragraph_info(info, word_range, include_runs)
    return info


def _add_optional_paragraph_info(info: dict[str, Any], word_range: Any, include_runs: bool) -> None:
    warnings: list[str] = []
    try:
        info["horizontal_position_pt"] = word_range.Information(5)
    except Exception as exc:
        warnings.append(f"horizontal position unavailable: {exc}")
    try:
        list_format = word_range.ListFormat
        if list_format.ListType > 0:
            names = {
                1: "bullet",
                2: "simple_number",
                3: "upper_roman",
                4: "lower_roman",
                5: "upper_letter",
                6: "lower_letter",
            }
            info["list_type"] = names.get(list_format.ListType, f"type_{list_format.ListType}")
            info["list_level"] = list_format.ListLevelNumber
            info["list_string"] = list_format.ListString
    except Exception as exc:
        warnings.append(f"list formatting unavailable: {exc}")
    try:
        info["highlight_color"] = word_range.HighlightColorIndex
    except Exception as exc:
        warnings.append(f"highlight unavailable: {exc}")
    if include_runs:
        try:
            info["runs"] = _word_runs(word_range)
        except Exception as exc:
            info["runs_error"] = f"Could not read word-level formatting: {exc}"
    if warnings:
        info["warnings"] = warnings


def _format_range(
    document: Any,
    start: int | None,
    end: int | None,
    start_paragraph: int | None,
    end_paragraph: int | None,
) -> tuple[Any, str]:
    resolved = character_or_paragraph_range(
        document,
        start=start,
        end=end,
        start_paragraph=start_paragraph,
        end_paragraph=end_paragraph,
    )
    return resolved.com_range, resolved.label


def _validate_format_options(
    document: Any,
    *,
    paragraph_alignment: str | None,
    font_color: str | None,
    font_size: float | None,
    highlight_color: int | None,
    style_name: str | None,
) -> tuple[int | None, int | None]:
    alignment = None
    if paragraph_alignment is not None:
        try:
            alignment = _ALIGNMENTS[paragraph_alignment.casefold()]
        except KeyError as exc:
            raise ValueError(f"Invalid alignment: {paragraph_alignment}") from exc
    color = rgb_hex_to_word(font_color, field_name="font_color") if font_color else None
    if font_size is not None and font_size <= 0:
        raise ValueError("font_size must be greater than zero")
    if highlight_color is not None and not 0 <= highlight_color <= 16:
        raise ValueError("highlight_color must be between 0 and 16")
    if style_name is not None:
        try:
            document.Styles(style_name)
        except Exception as exc:
            raise ValueError(f"Word style not found: {style_name}") from exc
    return alignment, color


def _saved_direct_formats(word_range: Any) -> list[dict[str, Any]]:
    saved = []
    for paragraph in word_range.Paragraphs:
        font = paragraph.Range.Font
        formatting = paragraph.Format
        saved.append(
            {
                "paragraph": paragraph,
                "font_name": str(font.Name) if font.Name and font.Name != 9999999 else None,
                "font_size": font.Size if font.Size and font.Size != 9999999 else None,
                "bold": font.Bold if font.Bold != 9999999 else None,
                "italic": font.Italic if font.Italic != 9999999 else None,
                "strikethrough": font.StrikeThrough if font.StrikeThrough != 9999999 else None,
                "alignment": formatting.Alignment,
                "space_before": formatting.SpaceBefore,
                "space_after": formatting.SpaceAfter,
                "line_spacing": formatting.LineSpacing,
                "line_spacing_rule": formatting.LineSpacingRule,
            }
        )
    return saved


def _restore_direct_format(saved: dict[str, Any], style: Any) -> None:
    paragraph = saved["paragraph"]
    paragraph.Style = style
    optional_font = (
        ("Name", saved["font_name"]),
        ("Size", saved["font_size"]),
        ("Bold", saved["bold"]),
        ("Italic", saved["italic"]),
        ("StrikeThrough", saved["strikethrough"]),
    )
    for name, value in optional_font:
        if value is not None:
            setattr(paragraph.Range.Font, name, value)
    formatting = paragraph.Format
    for name, key in (
        ("Alignment", "alignment"),
        ("SpaceBefore", "space_before"),
        ("SpaceAfter", "space_after"),
        ("LineSpacingRule", "line_spacing_rule"),
        ("LineSpacing", "line_spacing"),
    ):
        setattr(formatting, name, saved[key])


def _apply_text_format(word_range: Any, options: dict[str, Any]) -> None:
    style_name = options["style_name"]
    if style_name is not None and options["preserve_direct_formatting"]:
        style = options["document"].Styles(style_name)
        for saved in _saved_direct_formats(word_range):
            _restore_direct_format(saved, style)
    elif style_name is not None:
        word_range.Style = style_name
    font_values = (
        ("Bold", options["bold"]),
        ("Italic", options["italic"]),
        ("Underline", None if options["underline"] is None else int(options["underline"])),
        ("StrikeThrough", options["strikethrough"]),
        ("Name", options["font_name"]),
        ("Size", options["font_size"]),
        ("Color", options["font_color_value"]),
    )
    for name, value in font_values:
        if value is not None:
            setattr(word_range.Font, name, value)
    if options["highlight_color"] is not None:
        word_range.HighlightColorIndex = options["highlight_color"]
    _apply_paragraph_format(word_range, options)


def _apply_paragraph_format(word_range: Any, options: dict[str, Any]) -> None:
    for paragraph in word_range.Paragraphs:
        if options["alignment_value"] is not None:
            paragraph.Format.Alignment = options["alignment_value"]
        if options["page_break_before"] is not None:
            paragraph.Format.PageBreakBefore = options["page_break_before"]


@word_tool(title="Word Live Get Paragraph Format", domain="formatting", change="read")
async def word_live_get_paragraph_format(
    filename: str | None = None,
    start_paragraph: Annotated[int, Field(ge=1)] | None = None,
    end_paragraph: Annotated[int, Field(ge=1)] | None = None,
    include_runs: bool = False,
) -> ParagraphFormatResult:
    """[Windows only] Inspect paragraph formatting properties for diagnostics.

    Returns detailed formatting info for each paragraph in the range. Essential for
    debugging layout issues like unexpected page breaks (caused by keep_with_next chains),
    broken list formatting, wrong styles, or inconsistent fonts.

    Per-paragraph fields returned: index, text_preview (first 80 chars), char_start, char_end,
    style, font_name, font_size, bold, italic, alignment, space_before_pt, space_after_pt,
    line_spacing, line_spacing_rule, left/right/first-line indents, horizontal position,
    page_break_before, keep_with_next, keep_together.
    Also: list_type, list_level, list_string (if paragraph is in a list), highlight_color.

    When include_runs=True, each paragraph also includes a "runs" array with per-run
    formatting: text, bold, italic, font_name, font_size. Consecutive words with identical
    formatting are grouped into a single run. Useful for detecting which specific words
    are bold/italic (e.g., bold sub-clause numbers in otherwise normal text).

    Args:
        filename: Document name or path (None = active document).
        start_paragraph: First paragraph (1-indexed, required).
        end_paragraph: Last paragraph (1-indexed, defaults to start_paragraph).
        include_runs: Include per-run (word-level) formatting detail (default False).

    Returns:
        Structured formatting details per paragraph.
    """
    word_session.require_windows("Live formatting tools")

    if start_paragraph is None:
        raise ValueError("start_paragraph is required (1-indexed)")

    if end_paragraph is None:
        end_paragraph = start_paragraph

    document = word_session.find_document(word_session.get_word_app(), filename)
    total_paragraphs = int(document.Paragraphs.Count)
    if end_paragraph > total_paragraphs:
        raise ValueError(
            f"Range {start_paragraph}-{end_paragraph} out of bounds "
            f"(document has {total_paragraphs} paragraphs)"
        )

    paragraphs = [
        _paragraph_format_info(document.Paragraphs(index), index, include_runs)
        for index in range(start_paragraph, end_paragraph + 1)
    ]
    return ParagraphFormatResult(document=str(document.Name), paragraphs=paragraphs)


@word_tool(title="Word Live Format Text", domain="formatting", change="edit", batchable=True)
async def word_live_format_text(
    filename: str | None = None,
    start: Annotated[int, Field(ge=0)] | None = None,
    end: Annotated[int, Field(ge=0)] | None = None,
    start_paragraph: Annotated[int, Field(ge=1)] | None = None,
    end_paragraph: Annotated[int, Field(ge=1)] | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strikethrough: bool | None = None,
    font_name: str | None = None,
    font_size: Annotated[float, Field(gt=0)] | None = None,
    font_color: Annotated[str, Field(pattern=r"^#?[0-9A-Fa-f]{6}$")] | None = None,
    highlight_color: Annotated[int, Field(ge=0, le=16)] | None = None,
    style_name: str | None = None,
    paragraph_alignment: Literal["left", "center", "right", "justify"] | None = None,
    page_break_before: bool | None = None,
    preserve_direct_formatting: bool = False,
    track_changes: bool = False,
) -> FormatTextResult:
    """[Windows only] Format text in an open Word document: font, color, highlight, style, alignment, page breaks.
    Use this tool for any visual/formatting change that does NOT alter the text content itself.

    Two addressing modes (provide one):
    - start/end: Character positions (from word_live_find_text or word_live_get_page_text).
    - start_paragraph/end_paragraph: 1-indexed paragraph range (from word_live_get_text etc.).

    Args:
        filename: Document name or path (None = active document).
        start: Start character position.
        end: End character position.
        start_paragraph: First paragraph index (1-indexed). Alternative to start/end.
        end_paragraph: Last paragraph index (1-indexed, defaults to start_paragraph).
        bold: Set bold (True/False).
        italic: Set italic (True/False).
        underline: Set underline (True/False).
        strikethrough: Set strikethrough (True/False).
        font_name: Font family (e.g., "Arial", "Times New Roman").
        font_size: Font size in points (e.g., 12).
        font_color: Text color as "#RRGGBB" hex (e.g., "#FF0000" for red).
        highlight_color: Text highlight background color index.
            0 = remove highlight, 1 = black, 2 = blue, 3 = turquoise,
            4 = bright green, 5 = pink, 6 = red, 7 = yellow,
            8 = white, 9 = dark blue, 10 = teal, 11 = green,
            12 = violet, 13 = dark red, 14 = dark yellow, 15 = gray, 16 = light gray.
            Common: 7=yellow (add), 0=none (remove).
        style_name: Apply a named Word style (e.g., "Heading 1", "Normal").
        paragraph_alignment: Paragraph alignment — "left" (0), "center" (1), "right" (2), "justify" (3).
            Applies to ALL paragraphs in the selected range.
        page_break_before: Set or clear PageBreakBefore on paragraphs in range (True/False).
        preserve_direct_formatting: When True and style_name is set, saves font/size/bold/italic/
            alignment/spacing before applying the style and restores them after. Useful for changing
            a paragraph's style (e.g., Heading 5 → Normal) without losing its visual formatting.
        track_changes: Track formatting changes as revisions.

    Returns:
        Structured information about the formatted range.
    """
    word_session.require_windows("Live formatting tools")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    word_range, range_label = _format_range(document, start, end, start_paragraph, end_paragraph)
    alignment_value, rgb_color = _validate_format_options(
        document,
        paragraph_alignment=paragraph_alignment,
        font_color=font_color,
        font_size=font_size,
        highlight_color=highlight_color,
        style_name=style_name,
    )
    options = {
        "document": document,
        "bold": bold,
        "italic": italic,
        "underline": underline,
        "strikethrough": strikethrough,
        "font_name": font_name,
        "font_size": font_size,
        "font_color_value": rgb_color,
        "highlight_color": highlight_color,
        "style_name": style_name,
        "alignment_value": alignment_value,
        "page_break_before": page_break_before,
        "preserve_direct_formatting": preserve_direct_formatting,
    }

    with word_session.undo_record(app, "MCP: Format Text"):
        with word_session.revision_tracking(app, document, track_changes, DEFAULT_AUTHOR):
            _apply_text_format(word_range, options)

    preview = str(word_range.Text)
    if len(preview) > 50:
        preview = preview[:50] + "..."

    return FormatTextResult(
        document=str(document.Name),
        range=range_label,
        text_preview=preview,
        tracked=track_changes,
    )
