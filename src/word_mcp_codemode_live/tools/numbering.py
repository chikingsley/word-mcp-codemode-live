"""Apply list and numbering formats in live Word documents."""

import re
from typing import Annotated, Any, Literal

from pydantic import Field

from word_mcp_codemode_live.defaults import DEFAULT_AUTHOR
from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word.values import rgb_hex_to_word

_LIST_TYPE_NAMES = {
    0: "none",
    1: "list_num_field_only",
    2: "bullet",
    3: "simple_numbering",
    4: "outline_numbering",
    5: "mixed_numbering",
    6: "picture_bullet",
}
_NUMBER_STYLES = {
    "arabic": 0,
    "uppercase_roman": 1,
    "lowercase_roman": 2,
    "uppercase_letter": 3,
    "lowercase_letter": 4,
}
_NUMBER_STYLE_NAMES = {value: key for key, value in _NUMBER_STYLES.items()}
_TRAILING_CHARACTER_NAMES = {0: "tab", 1: "space", 2: "nothing"}
_ALIGNMENT_NAMES = {0: "left", 1: "center", 2: "right"}


def _heading_style(document: Any, level: int) -> Any:
    """Return a built-in heading style without relying on its localized name."""
    return document.Styles(-1 - level)  # wdStyleHeading1 (-2) through Heading9 (-10)


def _list_level_details(list_level: Any) -> dict[str, Any]:
    number_style = int(list_level.NumberStyle)
    trailing_character = int(list_level.TrailingCharacter)
    alignment = int(list_level.Alignment)
    return {
        "level": int(list_level.Index),
        "number_format": str(list_level.NumberFormat),
        "number_style": _NUMBER_STYLE_NAMES.get(number_style, f"word_value_{number_style}"),
        "number_style_id": number_style,
        "start_at": int(list_level.StartAt),
        "reset_on_higher": int(list_level.ResetOnHigher),
        "linked_style": str(list_level.LinkedStyle),
        "number_position_points": float(list_level.NumberPosition),
        "text_position_points": float(list_level.TextPosition),
        "tab_position_points": float(list_level.TabPosition),
        "trailing_character": _TRAILING_CHARACTER_NAMES.get(
            trailing_character, f"word_value_{trailing_character}"
        ),
        "trailing_character_id": trailing_character,
        "alignment": _ALIGNMENT_NAMES.get(alignment, f"word_value_{alignment}"),
        "alignment_id": alignment,
    }


def _style_numbering_details(document: Any, level: int) -> dict[str, Any]:
    style = _heading_style(document, level)
    row: dict[str, Any] = {
        "heading_level": level,
        "style_name": str(style.NameLocal),
        "linked": False,
    }
    template = style.ListTemplate
    if template is None:
        return row
    style_level = int(style.ListLevelNumber)
    row.update(
        {
            "linked": style_level > 0,
            "list_level": style_level,
            "template_name": str(template.Name),
            "template_outline_numbered": bool(template.OutlineNumbered),
        }
    )
    if style_level > 0:
        row["level_definition"] = _list_level_details(template.ListLevels(style_level))
    return row


def _heading_numbering_snapshot(document: Any) -> dict[str, Any]:
    styles = [_style_numbering_details(document, level) for level in range(1, 10)]
    headings: list[dict[str, Any]] = []
    for paragraph_index in range(1, int(document.Paragraphs.Count) + 1):
        paragraph = document.Paragraphs(paragraph_index)
        paragraph_range = paragraph.Range
        outline_level = int(paragraph.OutlineLevel)
        if not 1 <= outline_level <= 9:
            continue

        list_format = paragraph_range.ListFormat
        list_type = int(list_format.ListType)
        row: dict[str, Any] = {
            "paragraph_index": paragraph_index,
            "heading_level": outline_level,
            "style_name": str(paragraph_range.Style),
            "start_offset": int(paragraph_range.Start),
            "end_offset": int(paragraph_range.End),
            "text": str(paragraph_range.Text).rstrip("\r\x07"),
            "numbered": list_type != 0,
            "list_type": _LIST_TYPE_NAMES.get(list_type, f"word_value_{list_type}"),
            "list_type_id": list_type,
        }
        if list_type != 0:
            list_level = int(list_format.ListLevelNumber)
            row.update(
                {
                    "list_level": list_level,
                    "list_value": int(list_format.ListValue),
                    "list_string": str(list_format.ListString),
                }
            )
            try:
                paragraph_template = list_format.ListTemplate
                row["template_name"] = str(paragraph_template.Name)
                paragraph_level = paragraph_template.ListLevels(list_level)
                row["template_linked_style"] = str(paragraph_level.LinkedStyle)
                row["level_definition"] = _list_level_details(paragraph_level)
            except Exception:
                row["template_name"] = None
                row["template_linked_style"] = None
        headings.append(row)

    return {
        "heading_styles": styles,
        "headings": headings,
        "heading_count": len(headings),
        "numbered_heading_count": sum(bool(row["numbered"]) for row in headings),
    }


def _normalize_level_mappings(
    mappings: list[tuple[str, dict[int, Any]]],
) -> dict[str, dict[int, Any]]:
    normalized: dict[str, dict[int, Any]] = {}
    for name, mapping in mappings:
        try:
            values = {int(level): value for level, value in mapping.items()}
        except (AttributeError, TypeError, ValueError) as exc:
            raise ValueError(f"{name} must be a mapping keyed by levels 1-9") from exc
        invalid = sorted(level for level in values if not 1 <= level <= 9)
        if invalid:
            raise ValueError(f"{name} contains levels outside 1-9: {invalid}")
        normalized[name] = values
    _validate_heading_options(normalized)
    return normalized


def _validate_heading_options(options: dict[str, dict[int, Any]]) -> None:
    for level, value in options["number_formats"].items():
        if not isinstance(value, str) or not value:
            raise ValueError(f"number_formats[{level}] must be a non-empty string")
        deeper = [int(token) for token in re.findall(r"%(\d+)", value) if int(token) > level]
        if deeper:
            raise ValueError(f"number_formats[{level}] cannot reference deeper levels: {deeper}")
    invalid_styles = sorted(
        {str(value) for value in options["number_styles"].values()} - set(_NUMBER_STYLES)
    )
    if invalid_styles:
        raise ValueError(f"number_styles contains unsupported values: {invalid_styles}")
    starts = options["start_at"].values()
    if any(not isinstance(value, int) or isinstance(value, bool) or value < 1 for value in starts):
        raise ValueError("start_at values must be positive integers")
    for name in ("number_position_points", "text_position_points"):
        if any(
            not isinstance(value, (int, float)) or isinstance(value, bool) or value < 0
            for value in options[name].values()
        ):
            raise ValueError(f"{name} values must be non-negative numbers")


def _existing_heading_numbering(snapshot: dict[str, Any]):
    style_names = {row["heading_level"]: row["style_name"] for row in snapshot["heading_styles"]}
    linked = [row["heading_level"] for row in snapshot["heading_styles"] if row["linked"]]
    numbered = [
        row["paragraph_index"]
        for row in snapshot["headings"]
        if row["numbered"] and row["style_name"] == style_names[row["heading_level"]]
    ]
    return style_names, linked, numbered


def _build_heading_template(document: Any, options: dict[str, dict[int, Any]]) -> Any:
    template = document.ListTemplates.Add(OutlineNumbered=True)
    for level in range(1, 10):
        list_level = template.ListLevels(level)
        list_level.NumberStyle = _NUMBER_STYLES[str(options["number_styles"].get(level, "arabic"))]
        default_format = ".".join(f"%{ancestor}" for ancestor in range(1, level + 1)) + "."
        list_level.NumberFormat = str(options["number_formats"].get(level, default_format))
        list_level.StartAt = int(options["start_at"].get(level, 1))
        list_level.ResetOnHigher = 0 if level == 1 else level - 1
        list_level.Alignment = 0
        list_level.NumberPosition = float(
            options["number_position_points"].get(level, (level - 1) * 18.0)
        )
        text_position = float(options["text_position_points"].get(level, level * 18.0))
        list_level.TextPosition = text_position
        list_level.TabPosition = text_position
        list_level.TrailingCharacter = 0
        _heading_style(document, level).LinkToListTemplate(
            ListTemplate=template, ListLevelNumber=level
        )
    return template


def _apply_heading_template(
    document: Any,
    template: Any,
    snapshot: dict[str, Any],
    style_names: dict[int, str],
) -> None:
    applied = 0
    for row in snapshot["headings"]:
        level = row["heading_level"]
        if row["style_name"] != style_names[level]:
            continue
        document.Paragraphs(row["paragraph_index"]).Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=template,
            ContinuePreviousList=applied > 0,
            ApplyTo=2,
            DefaultListBehavior=1,
            ApplyLevel=level,
        )
        applied += 1


def _invalid_numbered_headings(snapshot: dict[str, Any], style_names: dict[int, str]):
    invalid = []
    keys = ("number_format", "number_style_id", "start_at", "reset_on_higher")
    for row in snapshot["headings"]:
        level = row["heading_level"]
        if row["style_name"] != style_names[level]:
            continue
        expected = snapshot["heading_styles"][level - 1].get("level_definition", {})
        actual = row.get("level_definition", {})
        mismatches = {
            key: {"actual": actual.get(key), "expected": expected.get(key)}
            for key in keys
            if actual.get(key) != expected.get(key)
        }
        if not row["numbered"] or row.get("list_level") != level or mismatches:
            invalid.append(
                {
                    "paragraph_index": row["paragraph_index"],
                    "numbered": row["numbered"],
                    "list_level": row.get("list_level"),
                    "expected_level": level,
                    "definition_mismatches": mismatches,
                }
            )
    return invalid


def _parse_list_options(
    number_format: dict | None,
    number_style: dict | str | None,
    start_at: dict | None,
    level_map: dict | None,
    font_color: str | None,
) -> dict[str, Any]:
    formats = {int(key): value for key, value in (number_format or {1: "%1.", 2: "%1.%2."}).items()}
    starts = {int(key): int(value) for key, value in (start_at or {}).items()}
    levels = {int(key): int(value) for key, value in (level_map or {}).items()}
    if any(not 1 <= key <= 9 for key in formats | starts):
        raise ValueError("number-format and start-at levels must be between 1 and 9")
    if any(not 1 <= value <= 9 for value in levels.values()):
        raise ValueError("level_map values must be between 1 and 9")
    styles = _parse_list_styles(number_style, formats)
    color = rgb_hex_to_word(font_color, field_name="font_color") if font_color else None
    return {
        "formats": formats,
        "starts": starts,
        "levels": levels,
        "styles": styles,
        "color": color,
    }


def _parse_list_styles(number_style: dict | str | None, formats: dict[int, Any]) -> dict[int, int]:
    if isinstance(number_style, dict):
        unknown = set(number_style.values()) - set(_NUMBER_STYLES)
        if unknown:
            raise ValueError(f"Unknown number_style values: {sorted(unknown)}")
        return {int(key): _NUMBER_STYLES[value] for key, value in number_style.items()}
    if isinstance(number_style, str):
        if number_style not in _NUMBER_STYLES:
            raise ValueError(f"Unknown number_style: {number_style}")
        return {level: _NUMBER_STYLES[number_style] for level in formats}
    return {}


def _apply_multilevel_list(
    document: Any,
    start: int,
    end: int,
    level: int,
    continue_previous: bool,
    options: dict[str, Any],
) -> int:
    template = document.ListTemplates.Add(OutlineNumbered=True)
    for level_number, number_format in options["formats"].items():
        list_level = template.ListLevels(level_number)
        list_level.NumberFormat = number_format
        list_level.NumberStyle = options["styles"].get(level_number, 0)
        list_level.StartAt = options["starts"].get(level_number, 1)
        list_level.Alignment = 0
        list_level.NumberPosition = 0
        list_level.TextPosition = 28
        list_level.TabPosition = 28
    word_range = document.Range(
        document.Paragraphs(start).Range.Start, document.Paragraphs(end).Range.End
    )
    word_range.ListFormat.ApplyListTemplateWithLevel(
        ListTemplate=template,
        ContinuePreviousList=continue_previous,
        ApplyTo=2,
        DefaultListBehavior=0,
    )
    default_level = level + 1 if level > 0 else 1
    for index in range(start, end + 1):
        target_level = options["levels"].get(index, default_level)
        if target_level != 1:
            document.Paragraphs(index).Range.ListFormat.ListLevelNumber = target_level
    return end - start + 1


def _apply_simple_list(
    document: Any, start: int, end: int, list_type: str, level: int, continue_previous: bool
) -> int:
    gallery = document.Application.ListGalleries({"bullet": 1, "number": 2}[list_type])
    template = gallery.ListTemplates(1)
    for index in range(start, end + 1):
        paragraph = document.Paragraphs(index)
        paragraph.Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=template,
            ContinuePreviousList=index > start or continue_previous,
            DefaultListBehavior=1,
        )
        if level > 0:
            paragraph.Range.ListFormat.ListLevelNumber = level + 1
    return end - start + 1


def _apply_list_operation(
    document: Any,
    start: int,
    end: int,
    list_type: str,
    level: int,
    remove: bool,
    continue_previous: bool,
    options: dict[str, Any],
) -> int:
    if remove:
        for index in range(start, end + 1):
            document.Paragraphs(index).Range.ListFormat.RemoveNumbers()
        return end - start + 1
    if list_type == "multilevel":
        return _apply_multilevel_list(document, start, end, level, continue_previous, options)
    return _apply_simple_list(document, start, end, list_type, level, continue_previous)


@word_tool(
    title="Word Live Inspect Heading Numbering",
    domain="numbering",
    change="read",
)
async def word_live_inspect_heading_numbering(
    filename: str | None = None,
) -> dict[str, Any]:
    """Inspect built-in Heading 1-9 style links and native numbering on headings.

    This is an objective view of Word's native list state. Paragraph indexes are
    one-based; character offsets are Word's zero-based range positions.

    Args:
        filename: Open document name or full path (None = active document).
    """
    word_session.require_windows("Live numbering tools")

    document = word_session.find_document(word_session.get_word_app(), filename)
    return {
        "success": True,
        "document": str(document.Name),
        **_heading_numbering_snapshot(document),
    }


@word_tool(
    title="Word Live Setup Heading Numbering",
    domain="numbering",
    change="edit",
    batchable=True,
)
async def word_live_setup_heading_numbering(
    filename: str | None = None,
    number_formats: dict[int, str] | None = None,
    number_styles: dict[int, str] | None = None,
    start_at: dict[int, int] | None = None,
    number_position_points: dict[int, float] | None = None,
    text_position_points: dict[int, float] | None = None,
    replace_existing: bool = False,
) -> dict[str, Any]:
    """Configure one native nine-level list linked to built-in Heading 1-9.

    By default, level N renders all ancestor numbers (for example, ``1.2.3.``),
    uses Arabic numerals, starts at one, and restarts beneath its parent level.
    Passing a mapping overrides only the named one-based levels.

    The tool refuses to change a document that already has native numbering on
    a heading style or heading paragraph. Inspect first, then pass
    ``replace_existing=True`` only when replacing that scheme is intended.

    Args:
        filename: Open document name or full path (None = active document).
        number_formats: Level-to-format mapping using Word placeholders ``%1``-``%9``.
        number_styles: Level-to-style mapping: arabic, uppercase_roman,
            lowercase_roman, uppercase_letter, or lowercase_letter.
        start_at: Level-to-positive-start-number mapping.
        number_position_points: Level-to-number-position mapping in points.
        text_position_points: Level-to-text-position mapping in points.
        replace_existing: Explicitly allow replacement of existing native heading numbering.

    Returns:
        The post-edit native style definitions and numbered heading state.
    """
    word_session.require_windows("Live numbering tools")

    mappings: list[tuple[str, dict[int, Any]]] = [
        ("number_formats", number_formats or {}),
        ("number_styles", number_styles or {}),
        ("start_at", start_at or {}),
        ("number_position_points", number_position_points or {}),
        ("text_position_points", text_position_points or {}),
    ]
    normalized = _normalize_level_mappings(mappings)

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    before = _heading_numbering_snapshot(document)
    built_in_style_names, linked_styles, numbered_paragraphs = _existing_heading_numbering(before)
    if not replace_existing and (linked_styles or numbered_paragraphs):
        raise ValueError(
            "The document already has native heading numbering; no changes were made. "
            f"Linked heading levels: {linked_styles}; numbered heading paragraphs: "
            f"{numbered_paragraphs}. Inspect the scheme and pass replace_existing=true "
            "only to replace it."
        )

    with word_session.undo_transaction(app, document, "MCP: Setup Heading Numbering"):
        try:
            if replace_existing:
                for paragraph_index in numbered_paragraphs:
                    document.Paragraphs(paragraph_index).Range.ListFormat.RemoveNumbers()
            template = _build_heading_template(document, normalized)
            _apply_heading_template(document, template, before, built_in_style_names)
        except Exception as exc:
            raise RuntimeError(f"Word could not configure heading numbering: {exc}") from exc
        after = _heading_numbering_snapshot(document)
        invalid_headings = _invalid_numbered_headings(after, built_in_style_names)
        if invalid_headings:
            raise RuntimeError(
                "Word did not apply the new linked heading scheme to paragraphs: "
                f"{invalid_headings}"
            )
    return {
        "success": True,
        "document": str(document.Name),
        "replaced_existing": bool(linked_styles or numbered_paragraphs),
        "previous_native_numbering": {
            "linked_heading_levels": linked_styles,
            "numbered_heading_paragraphs": numbered_paragraphs,
        },
        **after,
    }


@word_tool(title="Word Live Apply List", domain="numbering", change="edit", batchable=True)
async def word_live_apply_list(
    filename: str | None = None,
    start_paragraph: Annotated[int, Field(ge=1)] | None = None,
    end_paragraph: Annotated[int, Field(ge=1)] | None = None,
    list_type: Literal["bullet", "number", "multilevel"] = "bullet",
    level: Annotated[int, Field(ge=0, le=8)] = 0,
    remove: bool = False,
    continue_previous: bool = False,
    number_format: dict | None = None,
    number_style: dict | None = None,
    start_at: dict | None = None,
    level_map: dict | None = None,
    track_changes: bool = False,
    font_color: str | None = None,
) -> dict[str, Any]:
    """[Windows only] Apply or remove bullet/numbered/multilevel list formatting on paragraphs.

    Args:
        filename: Document name or path (None = active document).
        start_paragraph: First paragraph to format (1-indexed, required).
        end_paragraph: Last paragraph to format (1-indexed, defaults to start_paragraph).
        list_type: "bullet", "number", or "multilevel" (outline numbered).
        level: Indentation level (0 = first level, 1 = second level, etc.).
            For multilevel, this sets the default list level for all paragraphs.
            Use level_map instead for per-paragraph level control.
        remove: If True, removes list formatting from the range.
        continue_previous: If True, continues numbering from a previous list above.
        number_format: (multilevel only) Dict mapping level (int) to format string.
            Example: {1: "4.%1.", 2: "(%2)", 3: "(%3)"} → "4.1.", "(a)", "(i)"
            Keys are 1-indexed levels. If not provided, defaults to {1: "%1.", 2: "%1.%2."}.
        number_style: (multilevel only) Dict mapping level (int) to numbering style string.
            Styles: "arabic" (1,2,3), "lowercase_letter" (a,b,c), "uppercase_letter" (A,B,C),
            "lowercase_roman" (i,ii,iii), "uppercase_roman" (I,II,III).
            Example: {1: "arabic", 2: "lowercase_letter", 3: "lowercase_roman"}
            If a string is given instead of dict, applies same style to all levels.
            Default: "arabic" for all levels.
        start_at: (multilevel only) Dict mapping level (int) to starting number.
            Example: {1: 5} → numbering starts at 5.
            If not provided, starts at 1.
        level_map: (multilevel only) Dict mapping one-based paragraph indices to levels.
            JSON keys may be numeric strings, for example {"12": 2}.
            Paragraphs not in the map stay at level 1 (or value of `level + 1`).
        track_changes: Track changes as revisions.
        font_color: Optional six-digit RGB hex color applied after list formatting.

    Returns:
        JSON with result info.
    """

    word_session.require_windows("Live Word editing")

    if start_paragraph is None:
        raise ValueError("start_paragraph is required (1-indexed)")

    if end_paragraph is None:
        end_paragraph = start_paragraph

    if list_type not in {"bullet", "number", "multilevel"}:
        raise ValueError("list_type must be bullet, number, or multilevel")
    if level < 0 or level > 8:
        raise ValueError("level must be between 0 and 8")

    options = _parse_list_options(number_format, number_style, start_at, level_map, font_color)
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    total_paras = int(doc.Paragraphs.Count)
    if start_paragraph < 1 or end_paragraph > total_paras:
        raise ValueError(
            f"Paragraph range {start_paragraph}-{end_paragraph} out of bounds "
            f"(doc has {total_paras} paragraphs)"
        )

    with word_session.undo_record(app, "MCP: Apply List"):
        with word_session.revision_tracking(app, doc, track_changes, DEFAULT_AUTHOR):
            formatted = _apply_list_operation(
                doc,
                start_paragraph,
                end_paragraph,
                list_type,
                level,
                remove,
                continue_previous,
                options,
            )
            if options["color"] is not None:
                for index in range(start_paragraph, end_paragraph + 1):
                    doc.Paragraphs(index).Range.Font.Color = options["color"]

    action = "removed" if remove else f"applied {list_type}"
    return {
        "success": True,
        "document": str(doc.Name),
        "action": action,
        "paragraphs": f"{start_paragraph}-{end_paragraph}",
        "count": formatted,
        "level": level,
        "tracked": track_changes,
    }
