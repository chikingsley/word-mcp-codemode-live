from typing import Any

import pytest

from word_mcp_codemode_live.core import word_com
from word_mcp_codemode_live.tools import layout_diagnostics


class FakeFind:
    def __init__(self, word_range: "FakeRange") -> None:
        self.word_range = word_range

    def ClearFormatting(self) -> None:
        pass

    def Execute(self, **_kwargs: Any) -> bool:
        match = next(
            (
                offset
                for offset in self.word_range.break_offsets
                if self.word_range.Start <= offset < self.word_range.End
            ),
            None,
        )
        if match is None:
            return False
        self.word_range.Start = match
        self.word_range.End = match + 1
        return True


class FakeRange:
    def __init__(
        self,
        start: int,
        end: int,
        text: str,
        page: int = 1,
        break_offsets: tuple[int, ...] = (),
    ) -> None:
        self.Start = start
        self.End = end
        self.Text = text
        self.page = page
        self.break_offsets = break_offsets
        self.Find = FakeFind(self)

    @property
    def Duplicate(self) -> "FakeRange":
        return FakeRange(self.Start, self.End, self.Text, self.page, self.break_offsets)

    def Information(self, _kind: int) -> int:
        return self.page

    def SetRange(self, start: int, end: int) -> None:
        self.Start = start
        self.End = end


class FakeFormat:
    def __init__(
        self,
        *,
        page_break_before: int = 0,
        keep_with_next: int = 0,
        keep_together: int = 0,
        widow_control: int = 0,
    ) -> None:
        self.PageBreakBefore = page_break_before
        self.KeepWithNext = keep_with_next
        self.KeepTogether = keep_together
        self.WidowControl = widow_control


class FakeParagraph:
    def __init__(self, word_range: FakeRange, paragraph_format: FakeFormat) -> None:
        self.Range = word_range
        self.Format = paragraph_format


class FakeCollection:
    def __init__(self, *items: Any) -> None:
        self.items = list(items)

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> Any:
        return self.items[index - 1]


class FakePageSetup:
    PageWidth = 612.0
    PageHeight = 792.0
    LeftMargin = 72.0
    RightMargin = 72.0
    TopMargin = 72.0
    BottomMargin = 72.0
    Gutter = 0.0
    GutterPos = 0
    Orientation = 0
    SectionStart = 2


class FakeSection:
    def __init__(self) -> None:
        self.Range = FakeRange(0, 24, "Alpha\rBeta\x0cGamma\r", 1)
        self.PageSetup = FakePageSetup()


class FakeDocument:
    Name = "layout.docx"

    def __init__(self) -> None:
        self.Content = FakeRange(0, 24, "Alpha\rBeta\x0cGamma\r", 1, (10,))
        self.Sections = FakeCollection(FakeSection())
        self.Paragraphs = FakeCollection(
            FakeParagraph(FakeRange(0, 6, "Alpha\r", 1), FakeFormat(keep_with_next=-1)),
            FakeParagraph(FakeRange(6, 18, "Beta\x0cGamma\r", 2), FakeFormat()),
        )

    def ComputeStatistics(self, statistic: int) -> int:
        assert statistic == 2
        return 2


@pytest.mark.asyncio
async def test_inspect_layout_reports_objective_pagination_state(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    document = FakeDocument()
    monkeypatch.setattr(layout_diagnostics.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)

    result = await layout_diagnostics.word_live_inspect_layout()

    assert result["page_count"] == 2
    assert result["sections"][0]["usable_width_points"] == 468.0
    assert result["manual_page_break_offsets"] == [10]
    assert result["controlled_paragraph_count"] == 2
    assert result["controlled_paragraphs"][0]["keep_with_next"] is True
    assert result["controlled_paragraphs"][1]["contains_manual_page_break"] is True
    assert result["findings"] == []


@pytest.mark.asyncio
async def test_inspect_layout_enforces_bounded_detail(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(layout_diagnostics.sys, "platform", "win32")
    with pytest.raises(ValueError, match="between 1 and 2000"):
        await layout_diagnostics.word_live_inspect_layout(max_controlled_paragraphs=0)


@pytest.mark.asyncio
async def test_inspect_layout_subtracts_top_gutter_from_height(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    document = FakeDocument()
    setup = document.Sections(1).PageSetup
    setup.Gutter = 36.0
    setup.GutterPos = 1
    monkeypatch.setattr(layout_diagnostics.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)

    result = await layout_diagnostics.word_live_inspect_layout()

    section = result["sections"][0]
    assert section["usable_width_points"] == 468.0
    assert section["usable_height_points"] == 612.0
    assert section["margins_points"]["gutter_position"] == "top"
