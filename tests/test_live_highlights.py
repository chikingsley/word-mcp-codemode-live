from types import SimpleNamespace
from typing import Any

import pytest

from word_mcp_codemode_live.core import word_com
from word_mcp_codemode_live.tools import highlights as live_highlights


class FakeSearchRange:
    def __init__(self, owner: "FakeStoryRange") -> None:
        self.owner = owner
        self.Start = owner.Start
        self.End = owner.End
        self.Text = owner.Text
        self.HighlightColorIndex = 0
        self.Find = FakeFind(self)

    def SetRange(self, start: int, end: int) -> None:
        self.Start = start
        self.End = end
        self.Text = self.owner.Text[start - self.owner.Start : end - self.owner.Start]


class FakeFind:
    def __init__(self, search_range: FakeSearchRange) -> None:
        self.search_range = search_range

    def ClearFormatting(self) -> None:
        pass

    def Execute(self) -> bool:
        for start, end, text, color in self.search_range.owner.highlights:
            if start >= self.search_range.Start:
                self.search_range.Start = start
                self.search_range.End = end
                self.search_range.Text = text
                self.search_range.HighlightColorIndex = color
                return True
        return False


class FakeStoryRange:
    def __init__(
        self,
        start: int,
        text: str,
        highlights: list[tuple[int, int, str, int]],
        next_story: "FakeStoryRange | None" = None,
        uniform_color: int = 9999999,
    ) -> None:
        self.Start = start
        self.End = start + len(text)
        self.Text = text
        self.highlights = highlights
        self.NextStoryRange = next_story
        self.HighlightColorIndex = uniform_color

    @property
    def Duplicate(self) -> FakeSearchRange:
        return FakeSearchRange(self)


class FakeStoryRanges:
    def __init__(self, items: dict[int, FakeStoryRange]) -> None:
        self.items = items

    def __call__(self, story_type: int) -> Any:
        if story_type not in self.items:
            raise RuntimeError("story absent")
        return self.items[story_type]


@pytest.fixture
def fake_document(monkeypatch: pytest.MonkeyPatch) -> SimpleNamespace:
    second_header = FakeStoryRange(0, "More", [(0, 4, "More", 6)], uniform_color=6)
    document = SimpleNamespace(
        Name="highlights.docx",
        StoryRanges=FakeStoryRanges(
            {
                1: FakeStoryRange(
                    0, "Alpha beta gamma", [(0, 5, "Alpha", 7), (11, 16, "gamma", 3)]
                ),
                7: FakeStoryRange(0, "Head", [], second_header),
            }
        ),
    )
    monkeypatch.setattr(live_highlights.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    return document


@pytest.mark.asyncio
async def test_highlights_traverse_linked_stories_and_report_absent_types(
    fake_document: SimpleNamespace,
) -> None:
    result = await live_highlights.word_live_inspect_highlighted_text()

    assert result["highlight_count"] == 3
    assert [entry["story"] for entry in result["highlights"]] == [
        "main_text",
        "main_text",
        "primary_header",
    ]
    assert result["highlights"][0]["color"] == "yellow"
    assert result["highlights"][2]["story_instance_index"] == 2
    assert "footnotes" in result["absent_story_types"]
    assert len(result["searched_stories"]) == 3


@pytest.mark.asyncio
async def test_highlights_limit_is_explicit(fake_document: SimpleNamespace) -> None:
    result = await live_highlights.word_live_inspect_highlighted_text(max_results=2)
    assert result["highlight_count"] == 2
    assert result["truncated"] is True
    with pytest.raises(ValueError, match="at least 1"):
        await live_highlights.word_live_inspect_highlighted_text(max_results=0)
