import gc
import sys
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from word_mcp_codemode_live.tools import navigation
from word_mcp_codemode_live.word import session as word_com


class FakeRange:
    def __init__(self, document: "FakeDocument", start: int, end: int) -> None:
        self.document = document
        self.Start = start
        self.End = end

    @property
    def Duplicate(self) -> "FakeRange":
        return FakeRange(self.document, self.Start, self.End)

    def SetRange(self, start: int, end: int) -> None:
        self.Start = start
        self.End = end

    def Collapse(self, direction: int) -> None:
        if direction == 0:
            self.Start = self.End
        else:
            self.End = self.Start

    def Information(self, kind: int) -> int:
        assert kind == 3
        return 1 if self.Start < 20 else 2

    def Select(self) -> None:
        self.document.application.Selection.Range = self.Duplicate


class FakeWindow:
    def __init__(self) -> None:
        self.scrolled: tuple[int, int] | None = None

    def ScrollIntoView(self, word_range: FakeRange, start: bool) -> None:
        assert start is True
        self.scrolled = (word_range.Start, word_range.End)


class FakeSelection:
    def __init__(self) -> None:
        self.Range: FakeRange | None = None

    def Information(self, kind: int) -> int:
        assert self.Range is not None
        return self.Range.Information(kind)


class FakeDocument:
    def __init__(self) -> None:
        self.Name = "navigation.docx"
        self.Saved = True
        self.Content = SimpleNamespace(End=40, Information=lambda kind: 2 if kind == 4 else None)
        self.ActiveWindow = FakeWindow()
        self.application: Any = None
        self.activated = False

    def Activate(self) -> None:
        self.activated = True

    def GoTo(self, *, What: int, Which: int, Count: int) -> FakeRange:
        assert (What, Which) == (1, 1)
        return FakeRange(self, 0 if Count == 1 else 20, 0 if Count == 1 else 20)

    def Range(self, *, Start: int, End: int) -> FakeRange:
        return FakeRange(self, Start, End)


@pytest.fixture
def fake_word(monkeypatch: pytest.MonkeyPatch) -> tuple[Any, FakeDocument]:
    document = FakeDocument()
    application = SimpleNamespace(Visible=True, Selection=FakeSelection())
    document.application = application
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: application)
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    return application, document


@pytest.mark.asyncio
async def test_navigate_to_page_reports_actual_selection_and_visibility(
    fake_word: tuple[Any, FakeDocument],
) -> None:
    _application, document = fake_word

    result = await navigation.word_live_navigate(page=2)

    assert document.activated is True
    assert document.ActiveWindow.scrolled == (20, 20)
    assert result["requested"] == {"kind": "page", "page": 2}
    assert result["selection"] == {
        "char_start": 20,
        "char_end": 20,
        "collapsed": True,
        "start_page": 2,
        "end_page": 2,
        "active_end_page": 2,
    }
    assert result["application_visible"] is True
    assert result["saved_state_unchanged"] is True


@pytest.mark.asyncio
async def test_navigate_to_range_preserves_dirty_state(
    fake_word: tuple[Any, FakeDocument],
) -> None:
    _application, document = fake_word
    document.Saved = False

    result = await navigation.word_live_navigate(char_start=4, char_end=25)

    assert result["selection"]["char_start"] == 4
    assert result["selection"]["char_end"] == 25
    assert result["selection"]["start_page"] == 1
    assert result["selection"]["end_page"] == 2
    assert result["saved_before"] is False
    assert result["saved_after"] is False


@pytest.mark.asyncio
async def test_range_end_page_uses_last_selected_character_not_excluded_boundary(
    fake_word: tuple[Any, FakeDocument],
) -> None:
    result = await navigation.word_live_navigate(char_start=0, char_end=20)

    assert result["selection"]["start_page"] == 1
    assert result["selection"]["end_page"] == 1


@pytest.mark.asyncio
async def test_navigate_validates_mode_and_bounds(fake_word: tuple[Any, FakeDocument]) -> None:
    with pytest.raises(ValueError, match="exactly one target"):
        await navigation.word_live_navigate()
    with pytest.raises(ValueError, match="exactly one target"):
        await navigation.word_live_navigate(page=1, char_start=0)
    with pytest.raises(ValueError, match="char_start is required"):
        await navigation.word_live_navigate(char_end=1)
    with pytest.raises(ValueError, match="greater than or equal"):
        await navigation.word_live_navigate(char_start=3, char_end=2)
    with pytest.raises(ValueError, match="out of bounds"):
        await navigation.word_live_navigate(char_start=0, char_end=41)
    with pytest.raises(ValueError, match="out of range"):
        await navigation.word_live_navigate(page=3)


@pytest.mark.skipif(sys.platform != "win32", reason="Microsoft Word COM requires Windows")
@pytest.mark.asyncio
async def test_navigate_real_hidden_word_is_non_mutating(tmp_path: Path) -> None:
    import win32com.client

    try:
        application = win32com.client.DispatchEx("Word.Application")
    except Exception as exc:
        pytest.skip(f"Microsoft Word is unavailable: {exc}")

    application.Visible = False
    application.DisplayAlerts = 0
    document = None
    try:
        document = application.Documents.Add()
        document.Content.Text = "First page marker\rSecond page marker\r"
        break_at = int(document.Paragraphs(1).Range.End) - 1
        document.Range(Start=break_at, End=break_at).InsertBreak(7)  # wdPageBreak
        document.SaveAs2(
            str(tmp_path / "navigation-verification.docx"),
            FileFormat=16,
            AddToRecentFiles=False,
        )
        original_text = str(document.Content.Text)
        original_saved = bool(document.Saved)
        assert original_saved is True
        word_com.remember_word_app(application)

        page_result = await navigation.word_live_navigate(page=2)

        assert page_result["selection"]["active_end_page"] == 2
        assert page_result["selection"]["collapsed"] is True
        assert page_result["application_visible"] is False
        assert str(document.Content.Text) == original_text
        assert bool(document.Saved) is original_saved

        range_result = await navigation.word_live_navigate(char_start=1, char_end=6)

        assert range_result["selection"]["char_start"] == 1
        assert range_result["selection"]["char_end"] == 6
        assert str(document.Content.Text) == original_text
        assert bool(document.Saved) is original_saved
    finally:
        if document is not None:
            document.Close(SaveChanges=False)
            document = None
        word_com._WORD_APP = None
        gc.collect()
        application.Quit(SaveChanges=False)
        application = None
