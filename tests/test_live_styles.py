from contextlib import contextmanager
from typing import Any

import pytest

from word_mcp_codemode_live.tools import styles as live_styles
from word_mcp_codemode_live.word import session as word_com


class FakeFont:
    def __init__(self) -> None:
        self.Name = "Calibri"
        self.Size = 11.0
        self.Bold = False
        self.Italic = False
        self.Underline = 0
        self.Color = 0


class FakeParagraphFormat:
    def __init__(self) -> None:
        self.Alignment = 0
        self.SpaceBefore = 0.0
        self.SpaceAfter = 0.0
        self.LeftIndent = 0.0
        self.RightIndent = 0.0
        self.FirstLineIndent = 0.0
        self.KeepWithNext = False
        self.KeepTogether = False
        self.PageBreakBefore = False
        self.OutlineLevel = 10


class FakeStyle:
    def __init__(self, owner: "FakeStyles", name: str, style_type: int, built_in: bool) -> None:
        self.owner = owner
        self.NameLocal = name
        self.Type = style_type
        self.BuiltIn = built_in
        self.BaseStyle: Any = None
        self.AutomaticallyUpdate = False
        self.Font = FakeFont()
        self.ParagraphFormat = FakeParagraphFormat()

    def Delete(self) -> None:
        self.owner.items.remove(self)


class FakeStyles:
    def __init__(self) -> None:
        self.items = [FakeStyle(self, "Normal", 1, True)]

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, key: int | str) -> FakeStyle:
        if isinstance(key, int):
            return self.items[key - 1]
        for style in self.items:
            if style.NameLocal.casefold() == key.casefold():
                return style
        raise RuntimeError("style missing")

    def Add(self, *, Name: str, Type: int) -> FakeStyle:
        style = FakeStyle(self, Name, Type, False)
        self.items.append(style)
        return style


class FakeDocument:
    def __init__(self) -> None:
        self.Name = "styles.docx"
        self.Styles = FakeStyles()


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
async def test_create_update_list_and_delete_custom_style(fake_document: FakeDocument) -> None:
    created = await live_styles.word_live_create_custom_style(
        style_name="Contract Body",
        base_style="Normal",
        font_name="Arial",
        font_size=10.5,
        bold=True,
        font_color="#123456",
        alignment="justify",
        space_after=6,
        keep_together=True,
        outline_level=3,
    )

    assert created["style"]["name"] == "Contract Body"
    assert created["style"]["base_style"] == "Normal"
    assert created["style"]["font"]["color_bgr"] == 0x563412
    assert created["style"]["paragraph"]["alignment"] == "justify"
    assert created["style"]["paragraph"]["outline_level"] == 3

    listed = await live_styles.word_live_list_custom_styles()
    assert listed["custom_style_count"] == 1
    assert listed["styles"][0]["index"] == 1
    assert listed["styles"][0]["collection_index"] == 2

    updated = await live_styles.word_live_update_custom_style(
        style_name="Contract Body", bold=False, italic=True, page_break_before=True
    )
    assert updated["style"]["font"]["bold"] is False
    assert updated["style"]["font"]["italic"] is True
    assert updated["style"]["paragraph"]["page_break_before"] is True

    deleted = await live_styles.word_live_delete_custom_style(style_name="Contract Body")
    assert deleted["deleted_style"]["name"] == "Contract Body"
    assert fake_document.Styles.Count == 1


@pytest.mark.asyncio
async def test_custom_style_tools_reject_unsafe_or_invalid_requests(
    fake_document: FakeDocument,
) -> None:
    with pytest.raises(ValueError, match="already exists"):
        await live_styles.word_live_create_custom_style(style_name="Normal")
    with pytest.raises(ValueError, match="six-digit"):
        await live_styles.word_live_create_custom_style(style_name="Bad", font_color="red")
    with pytest.raises(ValueError, match="custom styles"):
        await live_styles.word_live_update_custom_style(style_name="Normal", bold=True)
    with pytest.raises(ValueError, match="Built-in"):
        await live_styles.word_live_delete_custom_style(style_name="Normal")
    with pytest.raises(ValueError, match="paragraph formatting"):
        await live_styles.word_live_create_custom_style(
            style_name="Character Accent", style_type="character", alignment="left"
        )
    with pytest.raises(ValueError, match="automatically_update"):
        await live_styles.word_live_create_custom_style(
            style_name="Character Accent",
            style_type="character",
            automatically_update=False,
        )


@pytest.mark.asyncio
async def test_character_style_round_trip_does_not_access_paragraph_only_properties(
    fake_document: FakeDocument,
) -> None:
    created = await live_styles.word_live_create_custom_style(
        style_name="Character Accent", style_type="character", bold=True
    )
    assert created["style"]["type"] == "character"
    assert created["style"]["automatically_update"] is None
    assert "paragraph" not in created["style"]

    listed = await live_styles.word_live_list_custom_styles()
    assert listed["styles"][0]["name"] == "Character Accent"

    updated = await live_styles.word_live_update_custom_style(
        style_name="Character Accent", italic=True
    )
    assert updated["style"]["font"]["italic"] is True

    deleted = await live_styles.word_live_delete_custom_style(style_name="Character Accent")
    assert deleted["deleted_style"]["type"] == "character"
