"""Inspect and edit native Microsoft Word headers, footers, and page fields."""

import re
from typing import Any, Literal

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word.values import rgb_hex_to_word

StoryKind = Literal["header", "footer"]
StoryVariant = Literal["primary", "first", "even"]
EditOperation = Literal["set", "clear"]
WriteMode = Literal["replace", "append"]
Alignment = Literal["left", "center", "right", "justify"]
PageNumberStyle = Literal[
    "arabic",
    "upper_roman",
    "lower_roman",
    "upper_letter",
    "lower_letter",
]

_VARIANT_IDS = {"primary": 1, "first": 2, "even": 3}
_STORY_KINDS: tuple[StoryKind, ...] = ("header", "footer")
_VARIANTS: tuple[StoryVariant, ...] = ("primary", "first", "even")
_ALIGNMENT_IDS = {"left": 0, "center": 1, "right": 2, "justify": 3}
_PAGE_STYLE_IDS = {
    "arabic": 0,
    "upper_roman": 1,
    "lower_roman": 2,
    "upper_letter": 3,
    "lower_letter": 4,
}
_PAGE_STYLE_NAMES = {value: key for key, value in _PAGE_STYLE_IDS.items()}
_FIELD_IDS = {"page": 33, "pages": 26, "section_pages": 66}
_FIELD_NAMES = {value: key for key, value in _FIELD_IDS.items()}
_TEMPLATE_TOKEN = re.compile(r"(\{(?:page|pages|section_pages)\})")


def _section_indices(document: Any, section: int | Literal["all"]) -> list[int]:
    count = int(document.Sections.Count)
    if section == "all":
        return list(range(1, count + 1))
    if isinstance(section, bool) or not isinstance(section, int) or not 1 <= section <= count:
        raise ValueError(f"section must be 'all' or an integer between 1 and {count}")
    return [section]


def _story(section: Any, story_kind: StoryKind, variant: StoryVariant) -> Any:
    collection = section.Headers if story_kind == "header" else section.Footers
    return collection(_VARIANT_IDS[variant])


def _body_range(header_footer: Any) -> Any:
    word_range = header_footer.Range.Duplicate
    if word_range.End > word_range.Start:
        word_range.End -= 1  # Preserve Word's required final paragraph mark.
    return word_range


def _clean_text(value: Any) -> str:
    return str(value).rstrip("\r\x07")


def _safe_value(getter: Any) -> Any:
    try:
        value = getter()
        if isinstance(value, (str, int, float, bool)) or value is None:
            return value
        return str(value)
    except Exception:
        return None


def _rgb_hex(value: Any) -> str | None:
    try:
        color = int(value)
    except (TypeError, ValueError):
        return None
    if color < 0:
        return None
    red = color & 0xFF
    green = (color >> 8) & 0xFF
    blue = (color >> 16) & 0xFF
    return f"{red:02X}{green:02X}{blue:02X}"


def _parse_rgb(value: str) -> int:
    return rgb_hex_to_word(value, field_name="font_color")


def _field_entries(header_footer: Any) -> list[dict[str, Any]]:
    fields: list[dict[str, Any]] = []
    collection = header_footer.Range.Fields
    for index in range(1, int(collection.Count) + 1):
        field = collection(index)
        field_type = int(field.Type)
        fields.append(
            {
                "index": index,
                "type": _FIELD_NAMES.get(field_type, field_type),
                "code": _clean_text(field.Code.Text).strip(),
                "result": _clean_text(field.Result.Text),
            }
        )
    return fields


def _story_state(header_footer: Any) -> dict[str, Any]:
    word_range = _body_range(header_footer)
    font = word_range.Font
    paragraph_format = word_range.ParagraphFormat
    page_numbers = header_footer.PageNumbers
    number_style = _safe_value(lambda: int(page_numbers.NumberStyle))
    return {
        "exists": bool(header_footer.Exists),
        "linked_to_previous": bool(header_footer.LinkToPrevious),
        "text": _clean_text(word_range.Text),
        "fields": _field_entries(header_footer),
        "alignment": _safe_value(lambda: int(paragraph_format.Alignment)),
        "font": {
            "name": _safe_value(lambda: font.Name),
            "size": _safe_value(lambda: float(font.Size)),
            "bold": _safe_value(lambda: int(font.Bold)),
            "italic": _safe_value(lambda: int(font.Italic)),
            "color": _rgb_hex(_safe_value(lambda: int(font.Color))),
        },
        "page_numbering": {
            "count": int(page_numbers.Count),
            "style": _PAGE_STYLE_NAMES.get(number_style, number_style),
            "restart_at_section": _safe_value(lambda: bool(page_numbers.RestartNumberingAtSection)),
            "starting_number": _safe_value(lambda: int(page_numbers.StartingNumber)),
        },
    }


def _insert_template(document: Any, header_footer: Any, content: str, write_mode: WriteMode) -> Any:
    body = _body_range(header_footer)
    if write_mode == "replace":
        body.Text = ""
    insertion_start = int(_body_range(header_footer).End)

    for part in _TEMPLATE_TOKEN.split(content):
        if not part:
            continue
        cursor = _body_range(header_footer)
        cursor.Collapse(0)  # wdCollapseEnd
        if part.startswith("{"):
            document.Fields.Add(Range=cursor, Type=_FIELD_IDS[part[1:-1]])
        else:
            cursor.InsertAfter(part)

    inserted = header_footer.Range.Duplicate
    inserted.SetRange(insertion_start, int(_body_range(header_footer).End))
    return inserted


def _linked_sections(
    document: Any,
    indices: list[int],
    story_kind: StoryKind,
    variant: StoryVariant,
) -> list[int]:
    return [
        index
        for index in indices
        if index > 1 and bool(_story(document.Sections(index), story_kind, variant).LinkToPrevious)
    ]


def _configure_section_variant(
    section: Any, variant: StoryVariant, options: dict[str, Any]
) -> None:
    first_page = options["different_first_page"]
    odd_even = options["different_odd_even"]
    if first_page is not None or variant == "first":
        section.PageSetup.DifferentFirstPageHeaderFooter = (
            first_page if first_page is not None else True
        )
    if odd_even is not None or variant == "even":
        section.PageSetup.OddAndEvenPagesHeaderFooter = odd_even if odd_even is not None else True


def _apply_story_format(header_footer: Any, changed_range: Any, options: dict[str, Any]) -> None:
    if options["alignment"] is not None:
        header_footer.Range.ParagraphFormat.Alignment = _ALIGNMENT_IDS[options["alignment"]]
    assignments = (
        ("Name", options["font_name"]),
        ("Size", options["font_size"]),
        ("Bold", options["bold"]),
        ("Italic", options["italic"]),
        ("Color", options["rgb_color"]),
    )
    for attribute, value in assignments:
        if value is not None:
            setattr(changed_range.Font, attribute, value)


def _apply_page_numbering(header_footer: Any, options: dict[str, Any]) -> None:
    page_numbers = header_footer.PageNumbers
    if options["page_number_style"] is not None:
        page_numbers.NumberStyle = _PAGE_STYLE_IDS[options["page_number_style"]]
    if options["restart_page_numbering"] is not None:
        page_numbers.RestartNumberingAtSection = options["restart_page_numbering"]
    if options["start_at"] is not None:
        page_numbers.StartingNumber = options["start_at"]


def _edit_story(
    document: Any,
    index: int,
    story_kind: StoryKind,
    variant: StoryVariant,
    operation: EditOperation,
    content: str,
    write_mode: WriteMode,
    options: dict[str, Any],
) -> dict[str, Any]:
    section = document.Sections(index)
    _configure_section_variant(section, variant, options)
    header_footer = _story(section, story_kind, variant)
    if options["link_to_previous"] is not None and index > 1:
        header_footer.LinkToPrevious = options["link_to_previous"]
    if operation == "clear":
        _body_range(header_footer).Text = ""
        changed_range = _body_range(header_footer)
    else:
        changed_range = _insert_template(document, header_footer, content, write_mode)
    _apply_story_format(header_footer, changed_range, options)
    _apply_page_numbering(header_footer, options)
    return {"section": index, "state": _story_state(header_footer)}


@word_tool(
    title="Word Live Get Headers and Footers",
    domain="headers_footers",
    change="read",
)
async def word_live_get_headers_footers(
    filename: str | None = None,
    section: int | Literal["all"] = "all",
) -> dict[str, Any]:
    """Inspect native Word headers, footers, page fields, linkage, and formatting.

    Args:
        filename: Document name or path (None = active document).
        section: One-based section number, or "all".

    Returns:
        JSON for primary, first-page, and even-page headers and footers.
    """
    word_session.require_windows("Live header/footer tools")
    document = word_session.find_document(word_session.get_word_app(), filename)
    sections: list[dict[str, Any]] = []
    for index in _section_indices(document, section):
        word_section = document.Sections(index)
        section_state: dict[str, Any] = {
            "section": index,
            "different_first_page": bool(word_section.PageSetup.DifferentFirstPageHeaderFooter),
            "different_odd_even": bool(word_section.PageSetup.OddAndEvenPagesHeaderFooter),
            "headers": {},
            "footers": {},
        }
        for story_kind in _STORY_KINDS:
            target = section_state[f"{story_kind}s"]
            for variant in _VARIANTS:
                target[variant] = _story_state(_story(word_section, story_kind, variant))
        sections.append(section_state)
    return {"success": True, "document": str(document.Name), "sections": sections}


@word_tool(
    title="Word Live Edit Headers and Footers",
    domain="headers_footers",
    change="edit",
    batchable=True,
)
async def word_live_edit_headers_footers(
    filename: str | None = None,
    section: int | Literal["all"] = 1,
    story_kind: StoryKind = "footer",
    variant: StoryVariant = "primary",
    operation: EditOperation = "set",
    content: str = "{page}",
    write_mode: WriteMode = "replace",
    alignment: Alignment | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_color: str | None = None,
    link_to_previous: bool | None = None,
    different_first_page: bool | None = None,
    different_odd_even: bool | None = None,
    page_number_style: PageNumberStyle | None = None,
    restart_page_numbering: bool | None = None,
    start_at: int | None = None,
) -> dict[str, Any]:
    """Set or clear one native Word header/footer variant without touching other stories.

    ``content`` may contain ``{page}``, ``{pages}``, and ``{section_pages}`` fields.
    Use ``variant='first'`` or ``variant='even'`` for different first/even pages.
    Section 2+ linked stories require an explicit ``link_to_previous`` choice before
    content can be changed, preventing an accidental edit to an earlier section.

    Args:
        filename: Document name or path (None = active document).
        section: One-based section number, or "all".
        story_kind: "header" or "footer".
        variant: "primary", "first", or "even".
        operation: "set" or "clear".
        content: Literal text plus optional page-field tokens.
        write_mode: "replace" or "append" for operation="set".
        alignment: Optional left, center, right, or justify alignment.
        font_name: Optional Word font name.
        font_size: Optional font size in points.
        bold: Optional bold override.
        italic: Optional italic override.
        font_color: Optional six-digit RGB hex color.
        link_to_previous: Optional Word section linkage setting.
        different_first_page: Optional section first-page toggle.
        different_odd_even: Optional section odd/even-page toggle.
        page_number_style: Arabic, Roman, or letter numbering style.
        restart_page_numbering: True to restart in each edited section; False to continue.
        start_at: First number when restarting (requires restart_page_numbering=True).

    Returns:
        JSON with the actual post-edit state of each affected story.
    """
    word_session.require_windows("Live header/footer tools")
    if operation == "set" and not content:
        raise ValueError("content is required for operation='set'; use operation='clear'")
    if font_size is not None and font_size <= 0:
        raise ValueError("font_size must be greater than zero")
    if start_at is not None and start_at < 0:
        raise ValueError("start_at must be zero or greater")
    if start_at is not None and restart_page_numbering is not True:
        raise ValueError("start_at requires restart_page_numbering=true")

    rgb_color = _parse_rgb(font_color) if font_color is not None else None
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    indices = _section_indices(document, section)

    if operation in {"set", "clear"} and link_to_previous is None:
        linked = _linked_sections(document, indices, story_kind, variant)
        if linked:
            raise RuntimeError(
                "The selected header/footer is linked to the previous section in section(s) "
                f"{linked}. Set link_to_previous=false to detach it, or true to intentionally "
                "edit the shared story."
            )

    options = {
        "alignment": alignment,
        "font_name": font_name,
        "font_size": font_size,
        "bold": bold,
        "italic": italic,
        "rgb_color": rgb_color,
        "link_to_previous": link_to_previous,
        "different_first_page": different_first_page,
        "different_odd_even": different_odd_even,
        "page_number_style": page_number_style,
        "restart_page_numbering": restart_page_numbering,
        "start_at": start_at,
    }
    with word_session.undo_record(app, "MCP: Edit Headers/Footers"):
        results = [
            _edit_story(
                document, index, story_kind, variant, operation, content, write_mode, options
            )
            for index in indices
        ]

    return {
        "success": True,
        "document": str(document.Name),
        "story_kind": story_kind,
        "variant": variant,
        "operation": operation,
        "results": results,
    }
