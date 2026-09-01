from types import SimpleNamespace

import pytest

from word_mcp_codemode_live.tools import content
from word_mcp_codemode_live.word import session as word_com


class _Range:
    def __init__(self, start: int) -> None:
        self.Start = start
        self.breaks: list[int] = []

    @property
    def Duplicate(self):
        return self

    def Collapse(self, _direction: int) -> None:
        pass

    def InsertBreak(self, break_type: int) -> None:
        self.breaks.append(break_type)


class _Document:
    Name = "content.docx"

    def __init__(self) -> None:
        self.Content = SimpleNamespace(End=21)
        self._paragraph_ranges = [_Range(0), _Range(10)]
        self.Paragraphs = _Paragraphs(self._paragraph_ranges)
        self.created_ranges: list[_Range] = []

    def Range(self, start: int, _end: int) -> _Range:
        result = _Range(start)
        self.created_ranges.append(result)
        return result

    def ComputeStatistics(self, _kind: int) -> int:
        return 1

    def Repaginate(self) -> None:
        pass


class _Paragraphs:
    Count = 2

    def __init__(self, ranges: list[_Range]) -> None:
        self._ranges = ranges

    def __call__(self, index: int):
        return SimpleNamespace(Range=self._ranges[index - 1])


@pytest.mark.asyncio
async def test_insert_page_break_targets_one_based_paragraph(monkeypatch) -> None:
    document = _Document()
    app = SimpleNamespace(UndoRecord=SimpleNamespace(IsRecordingCustomRecord=True))
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: app)
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)

    result = await content.word_live_insert_page_break(paragraph_index=2)

    assert result["target"] == {"kind": "paragraph", "paragraph_index": 2}
    assert document._paragraph_ranges[1].breaks == [7]


@pytest.mark.asyncio
async def test_insert_page_break_rejects_ambiguous_target(monkeypatch) -> None:
    monkeypatch.setattr(word_com.sys, "platform", "win32")

    with pytest.raises(ValueError, match="not both"):
        await content.word_live_insert_page_break(paragraph_index=1, character_offset=0)
