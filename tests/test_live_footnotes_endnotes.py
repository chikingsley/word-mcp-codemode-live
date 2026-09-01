from contextlib import contextmanager
from typing import Any

import pytest

from word_mcp_codemode_live.tools import notes as live_footnote_endnote_tools
from word_mcp_codemode_live.word import session as word_com


class FakeRange:
    def __init__(self, start: int, end: int, text: str = "") -> None:
        self.Start = start
        self.End = end
        self.Text = text

    @property
    def Duplicate(self) -> "FakeRange":
        return FakeRange(self.Start, self.End, self.Text)

    def Collapse(self, direction: int) -> None:
        if direction == 1:
            self.End = self.Start
        else:
            self.Start = self.End

    def Information(self, _kind: int) -> int:
        return 2


class FakeNote:
    def __init__(
        self, collection: "FakeNotes", index: int, reference: FakeRange, text: str
    ) -> None:
        self._collection = collection
        self.Index = index
        self.Reference = reference
        self.Range = FakeRange(0, len(text), text)

    def Delete(self) -> None:
        self._collection.items.remove(self)
        for index, note in enumerate(self._collection.items, 1):
            note.Index = index


class FakeNotes:
    def __init__(self) -> None:
        self.items: list[FakeNote] = []
        self.StartingNumber = 1
        self.NumberingRule = 0
        self.NumberStyle = 0
        self.Location = 0

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> FakeNote:
        return self.items[index - 1]

    def Add(self, *, Range: FakeRange) -> FakeNote:
        note = FakeNote(self, self.Count + 1, Range.Duplicate, "")
        self.items.append(note)
        return note

    def Convert(self) -> None:
        raise AssertionError("Conversion is covered by the real Word integration check")


class FakeParagraph:
    def __init__(self, start: int, end: int) -> None:
        self.Range = FakeRange(start, end)


class FakeParagraphs:
    def __init__(self) -> None:
        self.items = [FakeParagraph(10, 20)]

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> FakeParagraph:
        return self.items[index - 1]


class FakeDocument:
    def __init__(self) -> None:
        self.Name = "notes.docx"
        self.Footnotes = FakeNotes()
        self.Endnotes = FakeNotes()
        self.Paragraphs = FakeParagraphs()
        self._stories = {
            12: FakeRange(0, 1, "\x03\r"),
            13: FakeRange(0, 1, "\x03\r"),
            15: FakeRange(0, 1, "\x03\r"),
            16: FakeRange(0, 1, "\x03\r"),
        }

    def StoryRanges(self, story_type: int) -> FakeRange:
        return self._stories[story_type]


@contextmanager
def fake_undo_record(_app: Any, _name: str):  # type: ignore[no-untyped-def]
    yield


@pytest.mark.asyncio
async def test_native_note_tools_add_list_and_delete(monkeypatch: pytest.MonkeyPatch) -> None:
    document = FakeDocument()
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    monkeypatch.setattr(word_com, "undo_record", fake_undo_record)

    added = await live_footnote_endnote_tools.word_live_edit_footnotes_endnotes(
        operation="add",
        note_type="footnote",
        paragraph_index=1,
        text="Native note",
    )
    listed = await live_footnote_endnote_tools.word_live_list_footnotes_endnotes(note_type="all")

    assert added.success is True
    assert added.reference_start == 19
    assert added.before == {"footnotes": 0, "endnotes": 0}
    assert added.after == {"footnotes": 1, "endnotes": 0}
    assert [note.model_dump() for note in listed.notes] == [
        {
            "index": 1,
            "type": "footnote",
            "text": "Native note",
            "reference_start": 19,
            "page": 2,
        }
    ]

    deleted = await live_footnote_endnote_tools.word_live_edit_footnotes_endnotes(
        operation="delete",
        note_type="footnote",
        note_index=1,
    )

    assert deleted.success is True
    assert deleted.deleted_text == "Native note"
    assert deleted.after == {"footnotes": 0, "endnotes": 0}


@pytest.mark.asyncio
async def test_get_and_set_native_note_configuration(monkeypatch: pytest.MonkeyPatch) -> None:
    document = FakeDocument()
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    monkeypatch.setattr(word_com, "undo_record", fake_undo_record)

    result = await live_footnote_endnote_tools.word_live_set_note_configuration(
        note_type="footnote",
        starting_number=1,
        numbering_rule="restart_each_section",
        number_style="lowercase_roman",
        location="beneath_text",
        separator_text="---",
        continuation_separator_text="continued",
    )

    assert result["changed"] == [
        "starting_number",
        "numbering_rule",
        "number_style",
        "location",
        "separator_text",
        "continuation_separator_text",
    ]
    assert result["after"] == {
        "type": "footnote",
        "count": 0,
        "starting_number": 1,
        "numbering_rule": "restart_each_section",
        "numbering_rule_id": 1,
        "number_style": "lowercase_roman",
        "number_style_id": 2,
        "location": "beneath_text",
        "location_id": 1,
        "separator_text": "---",
        "continuation_separator_text": "continued",
    }

    inspected = await live_footnote_endnote_tools.word_live_get_note_configuration()
    assert inspected["footnotes"] == result["after"]
    assert inspected["endnotes"]["location"] == "end_of_section"


@pytest.mark.asyncio
async def test_note_configuration_validates_type_specific_location(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(word_com.sys, "platform", "win32")

    with pytest.raises(ValueError, match="location for endnote"):
        await live_footnote_endnote_tools.word_live_set_note_configuration(
            note_type="endnote", location="bottom_of_page"
        )


@pytest.mark.asyncio
async def test_note_configuration_rejects_custom_start_with_restart_rule(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    document = FakeDocument()
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)

    with pytest.raises(ValueError, match="requires starting_number=1"):
        await live_footnote_endnote_tools.word_live_set_note_configuration(
            note_type="footnote",
            starting_number=3,
            numbering_rule="restart_each_section",
        )

    with pytest.raises(ValueError, match="endnote numbering_rule"):
        await live_footnote_endnote_tools.word_live_set_note_configuration(
            note_type="endnote", numbering_rule="restart_each_page"
        )
