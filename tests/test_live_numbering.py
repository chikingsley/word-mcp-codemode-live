from contextlib import contextmanager
from typing import Any

import pytest

from word_mcp_codemode_live.tools import numbering as live_numbering
from word_mcp_codemode_live.word import session as word_com


class FakeListLevel:
    def __init__(self, index: int) -> None:
        self.Index = index
        self.NumberFormat = f"%{index}."
        self.NumberStyle = 0
        self.StartAt = 1
        self.ResetOnHigher = 0
        self.LinkedStyle = ""
        self.Alignment = 0
        self.NumberPosition = 0.0
        self.TextPosition = 18.0
        self.TabPosition = 18.0
        self.TrailingCharacter = 0


class FakeListLevels:
    def __init__(self) -> None:
        self.items = {index: FakeListLevel(index) for index in range(1, 10)}

    def __call__(self, index: int) -> FakeListLevel:
        return self.items[index]


class FakeListTemplate:
    def __init__(self, index: int, outline_numbered: bool) -> None:
        self.Name = f"Template {index}"
        self.OutlineNumbered = outline_numbered
        self.ListLevels = FakeListLevels()


class FakeListTemplates:
    def __init__(self) -> None:
        self.items: list[FakeListTemplate] = []

    def Add(self, *, OutlineNumbered: bool) -> FakeListTemplate:
        template = FakeListTemplate(len(self.items) + 1, OutlineNumbered)
        self.items.append(template)
        return template


class FakeListFormat:
    def __init__(self) -> None:
        self.ListType = 0
        self.ListLevelNumber = 0
        self.ListValue = 0
        self.ListString = ""
        self.ListTemplate: FakeListTemplate | None = None
        self.remove_count = 0

    def RemoveNumbers(self) -> None:
        self.remove_count += 1
        self.ListType = 0
        self.ListLevelNumber = 0
        self.ListValue = 0
        self.ListString = ""
        self.ListTemplate = None

    def ApplyListTemplateWithLevel(
        self,
        *,
        ListTemplate: FakeListTemplate,
        ApplyLevel: int,
        **_kwargs: Any,
    ) -> None:
        self.ListType = 4
        self.ListLevelNumber = ApplyLevel
        self.ListValue = 1
        self.ListString = "1." * ApplyLevel
        self.ListTemplate = ListTemplate


class FakeRange:
    def __init__(self, start: int, text: str, style_name: str) -> None:
        self.Start = start
        self.End = start + len(text) + 1
        self.Text = text + "\r"
        self.Style = style_name
        self.ListFormat = FakeListFormat()


class FakeParagraph:
    def __init__(self, index: int, outline_level: int, text: str) -> None:
        self.OutlineLevel = outline_level
        style_name = f"Heading {outline_level}" if outline_level <= 9 else "Normal"
        self.Range = FakeRange(index * 10, text, style_name)


class FakeParagraphs:
    def __init__(self, paragraphs: list[FakeParagraph]) -> None:
        self.items = paragraphs

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> FakeParagraph:
        return self.items[index - 1]


class FakeStyle:
    def __init__(self, level: int, document: "FakeDocument") -> None:
        self.level = level
        self.document = document
        self.NameLocal = f"Heading {level}"
        self.ListTemplate: FakeListTemplate | None = None
        self.ListLevelNumber = 0

    def LinkToListTemplate(self, *, ListTemplate: FakeListTemplate, ListLevelNumber: int) -> None:
        self.ListTemplate = ListTemplate
        self.ListLevelNumber = ListLevelNumber
        ListTemplate.ListLevels(ListLevelNumber).LinkedStyle = self.NameLocal
        for paragraph in self.document.Paragraphs.items:
            if paragraph.OutlineLevel == self.level:
                paragraph.Range.ListFormat.ListType = 4
                paragraph.Range.ListFormat.ListLevelNumber = self.level
                paragraph.Range.ListFormat.ListValue = 1
                paragraph.Range.ListFormat.ListString = "1." * self.level
                paragraph.Range.ListFormat.ListTemplate = ListTemplate


class FakeStyles:
    def __init__(self, document: "FakeDocument") -> None:
        self.items = {-1 - level: FakeStyle(level, document) for level in range(1, 10)}

    def __call__(self, style_id: int) -> FakeStyle:
        return self.items[style_id]


class FakeDocument:
    def __init__(self) -> None:
        self.Name = "numbering.docx"
        self.Paragraphs = FakeParagraphs(
            [
                FakeParagraph(1, 1, "First heading"),
                FakeParagraph(2, 10, "Body"),
                FakeParagraph(3, 2, "Nested heading"),
            ]
        )
        self.Styles = FakeStyles(self)
        self.ListTemplates = FakeListTemplates()


@contextmanager
def fake_undo_record(_app: Any, _name: str):  # type: ignore[no-untyped-def]
    yield


@pytest.fixture
def fake_document(monkeypatch: pytest.MonkeyPatch) -> FakeDocument:
    document = FakeDocument()
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    monkeypatch.setattr(word_com, "undo_record", fake_undo_record)
    return document


@pytest.mark.asyncio
async def test_setup_heading_numbering_links_all_levels_and_returns_post_state(
    fake_document: FakeDocument,
) -> None:
    result = await live_numbering.word_live_setup_heading_numbering(
        number_formats={1: "Article %1", 2: "%1-%2"},
        number_styles={1: "uppercase_roman", 2: "lowercase_letter"},
        start_at={1: 3},
        number_position_points={2: 24.0},
        text_position_points={2: 42.0},
    )

    assert result["success"] is True
    assert result["replaced_existing"] is False
    assert result["numbered_heading_count"] == 2
    assert result["headings"][0]["paragraph_index"] == 1
    assert result["headings"][1]["paragraph_index"] == 3
    assert result["headings"][1]["list_level"] == 2
    assert result["heading_styles"][0]["level_definition"]["number_format"] == "Article %1"
    assert result["heading_styles"][0]["level_definition"]["number_style"] == "uppercase_roman"
    assert result["heading_styles"][0]["level_definition"]["start_at"] == 3
    assert result["heading_styles"][1]["level_definition"]["number_format"] == "%1-%2"
    assert result["heading_styles"][1]["level_definition"]["number_style"] == "lowercase_letter"
    assert result["heading_styles"][1]["level_definition"]["reset_on_higher"] == 1
    assert result["heading_styles"][1]["level_definition"]["number_position_points"] == 24.0
    assert result["heading_styles"][1]["level_definition"]["text_position_points"] == 42.0


@pytest.mark.asyncio
async def test_setup_refuses_existing_native_heading_numbering_without_opt_in(
    fake_document: FakeDocument,
) -> None:
    await live_numbering.word_live_setup_heading_numbering()

    with pytest.raises(ValueError, match="already has native heading numbering"):
        await live_numbering.word_live_setup_heading_numbering()

    replacement = await live_numbering.word_live_setup_heading_numbering(replace_existing=True)
    assert replacement["replaced_existing"] is True
    assert replacement["previous_native_numbering"]["linked_heading_levels"] == list(range(1, 10))
    assert replacement["previous_native_numbering"]["numbered_heading_paragraphs"] == [1, 3]


@pytest.mark.asyncio
async def test_replace_existing_removes_direct_heading_numbering_first(
    fake_document: FakeDocument,
) -> None:
    list_format = fake_document.Paragraphs(1).Range.ListFormat
    direct_template = FakeListTemplate(99, False)
    list_format.ListType = 3
    list_format.ListLevelNumber = 1
    list_format.ListValue = 1
    list_format.ListString = "1."
    list_format.ListTemplate = direct_template

    result = await live_numbering.word_live_setup_heading_numbering(replace_existing=True)

    assert list_format.remove_count == 1
    assert result["headings"][0]["level_definition"]["number_format"] == "%1."


@pytest.mark.asyncio
async def test_inspect_heading_numbering_reports_unlinked_and_numbered_states(
    fake_document: FakeDocument,
) -> None:
    before = await live_numbering.word_live_inspect_heading_numbering()
    assert before["heading_count"] == 2
    assert before["numbered_heading_count"] == 0
    assert all(not row["linked"] for row in before["heading_styles"])

    await live_numbering.word_live_setup_heading_numbering()
    after = await live_numbering.word_live_inspect_heading_numbering()
    assert after["numbered_heading_count"] == 2
    assert all(row["linked"] for row in after["heading_styles"])


@pytest.mark.asyncio
async def test_setup_validates_level_maps_before_opening_word(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(word_com.sys, "platform", "win32")

    with pytest.raises(ValueError, match="outside 1-9"):
        await live_numbering.word_live_setup_heading_numbering(number_formats={10: "%10"})
    with pytest.raises(ValueError, match="deeper levels"):
        await live_numbering.word_live_setup_heading_numbering(number_formats={1: "%2."})
    with pytest.raises(ValueError, match="unsupported values"):
        await live_numbering.word_live_setup_heading_numbering(number_styles={1: "decimal"})
    with pytest.raises(ValueError, match="positive integers"):
        await live_numbering.word_live_setup_heading_numbering(start_at={1: 0})
    with pytest.raises(ValueError, match="non-negative numbers"):
        await live_numbering.word_live_setup_heading_numbering(text_position_points={1: -1})
