"""Inspect and edit native Microsoft Word footnotes and endnotes."""

from typing import Annotated, Any, Literal

from pydantic import BaseModel, Field

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session

NoteType = Literal["footnote", "endnote", "all"]
MutationNoteType = Literal["footnote", "endnote"]
NoteOperation = Literal["add", "delete", "convert"]
PositiveIndex = Annotated[int, Field(ge=1)]

_NUMBERING_RULES = {"continuous": 0, "restart_each_section": 1, "restart_each_page": 2}
_NUMBER_STYLES = {
    "arabic": 0,
    "uppercase_roman": 1,
    "lowercase_roman": 2,
    "uppercase_letter": 3,
    "lowercase_letter": 4,
    "symbol": 9,
}
_FOOTNOTE_LOCATIONS = {"bottom_of_page": 0, "beneath_text": 1}
_ENDNOTE_LOCATIONS = {"end_of_section": 0, "end_of_document": 1}
_SEPARATOR_STORIES = {
    "footnote": {"separator": 12, "continuation_separator": 13},
    "endnote": {"separator": 15, "continuation_separator": 16},
}


class NoteEntry(BaseModel):
    index: int
    type: MutationNoteType
    text: str
    reference_start: int
    page: int | None


class NoteListResult(BaseModel):
    success: Literal[True] = True
    document: str
    footnote_count: int
    endnote_count: int
    notes: list[NoteEntry]


class NoteEditResult(BaseModel):
    success: Literal[True] = True
    document: str
    operation: NoteOperation
    note_type: MutationNoteType
    before: dict[str, int]
    after: dict[str, int]
    note_index: int | None = None
    reference_start: int | None = None
    text: str | None = None
    deleted_text: str | None = None
    converted: int | None = None
    from_type: MutationNoteType | None = Field(default=None, alias="from")
    to: MutationNoteType | None = None


def _enum_name(value: int, values: dict[str, int]) -> str:
    return next((name for name, enum_value in values.items() if enum_value == value), "unknown")


def _story_text(document: Any, story_type: int) -> str | None:
    try:
        return str(document.StoryRanges(story_type).Text).rstrip("\r\x07")
    except Exception:
        return None


def _note_configuration(document: Any, note_type: MutationNoteType) -> dict[str, Any]:
    collection = _note_collection(document, note_type)
    locations = _FOOTNOTE_LOCATIONS if note_type == "footnote" else _ENDNOTE_LOCATIONS
    stories = _SEPARATOR_STORIES[note_type]
    return {
        "type": note_type,
        "count": int(collection.Count),
        "starting_number": int(collection.StartingNumber),
        "numbering_rule": _enum_name(int(collection.NumberingRule), _NUMBERING_RULES),
        "numbering_rule_id": int(collection.NumberingRule),
        "number_style": _enum_name(int(collection.NumberStyle), _NUMBER_STYLES),
        "number_style_id": int(collection.NumberStyle),
        "location": _enum_name(int(collection.Location), locations),
        "location_id": int(collection.Location),
        "separator_text": _story_text(document, stories["separator"]),
        "continuation_separator_text": _story_text(document, stories["continuation_separator"]),
    }


def _separator_range(document: Any, story_type: int) -> Any:
    try:
        return document.StoryRanges(story_type)
    except Exception as exc:
        raise RuntimeError(
            "Word has not created this note separator story; add at least one note of that type first"
        ) from exc


def _set_separator_text(story_range: Any, text: str) -> None:
    story_range.Text = f"{text.rstrip(chr(13))}\r"


def _note_collection(document: Any, note_type: MutationNoteType) -> Any:
    return document.Footnotes if note_type == "footnote" else document.Endnotes


def _note_entries(document: Any, note_type: MutationNoteType) -> list[NoteEntry]:
    collection = _note_collection(document, note_type)
    entries: list[NoteEntry] = []
    for index in range(1, collection.Count + 1):
        note = collection(index)
        reference = note.Reference
        try:
            page = int(reference.Information(3))  # wdActiveEndPageNumber
        except Exception:
            page = None
        entries.append(
            NoteEntry(
                index=index,
                type=note_type,
                text=str(note.Range.Text).rstrip("\r\x07"),
                reference_start=int(reference.Start),
                page=page,
            )
        )
    return entries


def _find_text_range(document: Any, target_text: str, occurrence: int) -> Any | None:
    search_range = document.Content.Duplicate
    for current_occurrence in range(1, occurrence + 1):
        search_range.Find.ClearFormatting()
        found = search_range.Find.Execute(
            FindText=target_text,
            Forward=True,
            MatchCase=False,
            MatchWholeWord=False,
            Wrap=0,
        )
        if not found:
            return None
        if current_occurrence == occurrence:
            return search_range.Duplicate
        next_start = int(search_range.End)
        search_range.SetRange(next_start, int(document.Content.End))
    return None


def _target_range(
    document: Any,
    *,
    target_text: str | None,
    paragraph_index: int | None,
    occurrence: int,
    position: Literal["before", "after"],
) -> Any:
    if bool(target_text) == (paragraph_index is not None):
        raise ValueError("Provide exactly one of target_text or paragraph_index")
    if occurrence < 1:
        raise ValueError("occurrence must be at least 1")

    if target_text:
        word_range = _find_text_range(document, target_text, occurrence)
        if word_range is None:
            raise ValueError(f"Occurrence {occurrence} of {target_text!r} was not found")
    else:
        if paragraph_index is None or not 1 <= paragraph_index <= document.Paragraphs.Count:
            raise ValueError(f"paragraph_index must be between 1 and {document.Paragraphs.Count}")
        word_range = document.Paragraphs(paragraph_index).Range.Duplicate
        if word_range.End > word_range.Start:
            word_range.End -= 1  # Exclude the paragraph mark.

    word_range.Collapse(1 if position == "before" else 0)  # wdCollapseStart / wdCollapseEnd
    return word_range


@word_tool(title="Word Live List Footnotes and Endnotes", domain="notes", change="read")
async def word_live_list_footnotes_endnotes(
    filename: str | None = None,
    note_type: NoteType = "all",
) -> NoteListResult:
    """List genuine Word footnotes and/or endnotes in an open document.

    Args:
        filename: Document name or path (None = active document).
        note_type: "footnote", "endnote", or "all".

    Returns:
        JSON containing note text, one-based index, reference position, and page.
    """
    word_session.require_windows("Live note tools")
    if note_type not in {"footnote", "endnote", "all"}:
        raise ValueError("note_type must be footnote, endnote, or all")

    document = word_session.find_document(word_session.get_word_app(), filename)
    notes: list[NoteEntry] = []
    if note_type in {"footnote", "all"}:
        notes.extend(_note_entries(document, "footnote"))
    if note_type in {"endnote", "all"}:
        notes.extend(_note_entries(document, "endnote"))
    return NoteListResult(
        document=str(document.Name),
        footnote_count=int(document.Footnotes.Count),
        endnote_count=int(document.Endnotes.Count),
        notes=notes,
    )


@word_tool(title="Word Live Get Note Configuration", domain="notes", change="read")
async def word_live_get_note_configuration(filename: str | None = None) -> dict[str, Any]:
    """Inspect native footnote/endnote numbering, placement, and separators."""
    word_session.require_windows("Live note tools")

    document = word_session.find_document(word_session.get_word_app(), filename)
    return {
        "success": True,
        "document": str(document.Name),
        "footnotes": _note_configuration(document, "footnote"),
        "endnotes": _note_configuration(document, "endnote"),
    }


def _validate_note_configuration(
    note_type: MutationNoteType,
    starting_number: int | None,
    numbering_rule: str | None,
    number_style: str | None,
    location: str | None,
    changes: tuple[Any, ...],
) -> dict[str, int]:
    if note_type not in {"footnote", "endnote"}:
        raise ValueError("note_type must be footnote or endnote")
    if starting_number is not None and starting_number < 1:
        raise ValueError("starting_number must be at least 1")
    if numbering_rule is not None and numbering_rule not in _NUMBERING_RULES:
        raise ValueError(f"numbering_rule must be one of {list(_NUMBERING_RULES)}")
    if note_type == "endnote" and numbering_rule == "restart_each_page":
        raise ValueError("endnote numbering_rule cannot be restart_each_page")
    if number_style is not None and number_style not in _NUMBER_STYLES:
        raise ValueError(f"number_style must be one of {list(_NUMBER_STYLES)}")
    locations = _FOOTNOTE_LOCATIONS if note_type == "footnote" else _ENDNOTE_LOCATIONS
    if location is not None and location not in locations:
        raise ValueError(f"location for {note_type} must be one of {list(locations)}")
    if all(value is None for value in changes):
        raise ValueError("Provide at least one note configuration change")
    return locations


def _apply_note_numbering(
    collection: Any,
    *,
    effective_rule: int,
    starting_number: int | None,
    numbering_rule: str | None,
    changed: list[str],
) -> None:
    assignments = (
        (
            ("StartingNumber", starting_number, "starting_number"),
            ("NumberingRule", numbering_rule, "numbering_rule"),
        )
        if effective_rule != _NUMBERING_RULES["continuous"]
        else (
            ("NumberingRule", numbering_rule, "numbering_rule"),
            ("StartingNumber", starting_number, "starting_number"),
        )
    )
    for attribute, requested, label in assignments:
        if requested is not None:
            value = effective_rule if attribute == "NumberingRule" else requested
            setattr(collection, attribute, value)
            changed.append(label)


def _apply_note_options(
    collection: Any,
    *,
    number_style: str | None,
    location: str | None,
    locations: dict[str, int],
    separator_range: Any,
    separator_text: str | None,
    continuation_range: Any,
    continuation_text: str | None,
    changed: list[str],
) -> None:
    if number_style is not None:
        collection.NumberStyle = _NUMBER_STYLES[number_style]
        changed.append("number_style")
    if location is not None:
        collection.Location = locations[location]
        changed.append("location")
    if separator_text is not None:
        _set_separator_text(separator_range, separator_text)
        changed.append("separator_text")
    if continuation_text is not None:
        _set_separator_text(continuation_range, continuation_text)
        changed.append("continuation_separator_text")


@word_tool(
    title="Word Live Set Note Configuration",
    domain="notes",
    change="edit",
    batchable=True,
)
async def word_live_set_note_configuration(
    filename: str | None = None,
    note_type: MutationNoteType = "footnote",
    starting_number: int | None = None,
    numbering_rule: Literal["continuous", "restart_each_section", "restart_each_page"]
    | None = None,
    number_style: Literal[
        "arabic",
        "uppercase_roman",
        "lowercase_roman",
        "uppercase_letter",
        "lowercase_letter",
        "symbol",
    ]
    | None = None,
    location: str | None = None,
    separator_text: str | None = None,
    continuation_separator_text: str | None = None,
) -> dict[str, Any]:
    """Set native note numbering, number style, placement, or separator text.

    ``location`` accepts ``bottom_of_page`` or ``beneath_text`` for footnotes,
    and ``end_of_section`` or ``end_of_document`` for endnotes. Separator stories
    exist only after Word has created at least one note of the selected type.
    """
    word_session.require_windows("Live note tools")
    changes = (
        starting_number,
        numbering_rule,
        number_style,
        location,
        separator_text,
        continuation_separator_text,
    )
    locations = _validate_note_configuration(
        note_type, starting_number, numbering_rule, number_style, location, changes
    )

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    collection = _note_collection(document, note_type)
    before = _note_configuration(document, note_type)
    effective_rule = (
        _NUMBERING_RULES[numbering_rule]
        if numbering_rule is not None
        else int(collection.NumberingRule)
    )
    effective_start = (
        starting_number if starting_number is not None else int(collection.StartingNumber)
    )
    if effective_rule != _NUMBERING_RULES["continuous"] and effective_start != 1:
        raise ValueError(
            "Word requires starting_number=1 when numbering_rule restarts each section or page"
        )
    changed: list[str] = []
    stories = _SEPARATOR_STORIES[note_type]
    separator_range = (
        _separator_range(document, stories["separator"]) if separator_text is not None else None
    )
    continuation_separator_range = (
        _separator_range(document, stories["continuation_separator"])
        if continuation_separator_text is not None
        else None
    )

    with word_session.undo_transaction(app, document, f"MCP: Configure {note_type.title()}s"):
        _apply_note_numbering(
            collection,
            effective_rule=effective_rule,
            starting_number=starting_number,
            numbering_rule=numbering_rule,
            changed=changed,
        )
        _apply_note_options(
            collection,
            number_style=number_style,
            location=location,
            locations=locations,
            separator_range=separator_range,
            separator_text=separator_text,
            continuation_range=continuation_separator_range,
            continuation_text=continuation_separator_text,
            changed=changed,
        )
        after = _note_configuration(document, note_type)

    return {
        "success": True,
        "document": str(document.Name),
        "note_type": note_type,
        "changed": changed,
        "before": before,
        "after": after,
    }


@word_tool(
    title="Word Live Edit Footnotes and Endnotes",
    domain="notes",
    change="edit",
    batchable=True,
)
async def word_live_edit_footnotes_endnotes(
    filename: str | None = None,
    operation: NoteOperation = "add",
    note_type: MutationNoteType = "footnote",
    target_text: str | None = None,
    paragraph_index: PositiveIndex | None = None,
    occurrence: PositiveIndex = 1,
    position: Literal["before", "after"] = "after",
    text: str = "",
    note_index: PositiveIndex | None = None,
) -> NoteEditResult:
    """Add, delete, or convert genuine Word footnotes and endnotes.

    For ``add``, provide exactly one target: ``target_text`` or a one-based
    ``paragraph_index``. For ``delete``, provide the one-based ``note_index``.
    For ``convert``, every note of ``note_type`` is converted to the other type.

    Args:
        filename: Document name or path (None = active document).
        operation: "add", "delete", or "convert".
        note_type: "footnote" or "endnote".
        target_text: Text whose selected occurrence anchors a new note.
        paragraph_index: One-based paragraph whose start/end anchors a new note.
        occurrence: One-based occurrence of target_text.
        position: Place a new note before or after the target.
        text: New note text for operation="add".
        note_index: One-based note index for operation="delete".

    Returns:
        JSON with native Word note counts and mutation details.
    """
    word_session.require_windows("Live note tools")
    if operation not in {"add", "delete", "convert"}:
        raise ValueError("operation must be add, delete, or convert")
    if note_type not in {"footnote", "endnote"}:
        raise ValueError("note_type must be footnote or endnote")
    if position not in {"before", "after"}:
        raise ValueError("position must be before or after")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    collection = _note_collection(document, note_type)
    before = {
        "footnotes": int(document.Footnotes.Count),
        "endnotes": int(document.Endnotes.Count),
    }
    result: dict[str, Any]

    with word_session.undo_record(app, f"MCP: {operation.title()} {note_type.title()}"):
        if operation == "add":
            if not text:
                raise ValueError("text is required for operation='add'")
            word_range = _target_range(
                document,
                target_text=target_text,
                paragraph_index=paragraph_index,
                occurrence=occurrence,
                position=position,
            )
            # Dynamic COM dispatch can silently drop Text when Word's optional
            # Reference argument is omitted. Create the native note first, then
            # assign its body explicitly through the returned Word Range.
            note = collection.Add(Range=word_range)
            note.Range.Text = text
            result = {
                "note_index": int(note.Index),
                "reference_start": int(note.Reference.Start),
                "text": text,
            }
        elif operation == "delete":
            if note_index is None or not 1 <= note_index <= collection.Count:
                raise ValueError(f"note_index must be between 1 and {collection.Count}")
            note = collection(note_index)
            deleted_text = str(note.Range.Text).rstrip("\r\x07")
            note.Delete()
            result = {"note_index": note_index, "deleted_text": deleted_text}
        else:
            converted = int(collection.Count)
            collection.Convert()
            result = {
                "converted": converted,
                "from": note_type,
                "to": "endnote" if note_type == "footnote" else "footnote",
            }

    after = {
        "footnotes": int(document.Footnotes.Count),
        "endnotes": int(document.Endnotes.Count),
    }
    return NoteEditResult(
        document=str(document.Name),
        operation=operation,
        note_type=note_type,
        before=before,
        after=after,
        **result,
    )
