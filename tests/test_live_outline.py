from types import SimpleNamespace

import pytest

from word_mcp_codemode_live.core import word_com
from word_mcp_codemode_live.tools import outline as live_outline


class FakeRange:
    def __init__(self, start: int, text: str, style: str, page: int) -> None:
        self.Start = start
        self.End = start + len(text) + 1
        self.Text = text + "\r"
        self.Style = SimpleNamespace(NameLocal=style)
        self.page = page

    def Information(self, kind: int) -> int:
        assert kind == 3
        return self.page


class FakeParagraph:
    def __init__(self, level: int, word_range: FakeRange) -> None:
        self.OutlineLevel = level
        self.Range = word_range


class FakeParagraphs:
    def __init__(self, items: list[FakeParagraph]) -> None:
        self.items = items

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> FakeParagraph:
        return self.items[index - 1]


@pytest.fixture
def fake_document(monkeypatch: pytest.MonkeyPatch) -> SimpleNamespace:
    document = SimpleNamespace(
        Name="outline.docx",
        Paragraphs=FakeParagraphs(
            [
                FakeParagraph(1, FakeRange(0, "Scope", "Heading 1", 1)),
                FakeParagraph(10, FakeRange(6, "Body", "Normal", 1)),
                FakeParagraph(3, FakeRange(11, "Details", "Custom Outline", 2)),
                FakeParagraph(7, FakeRange(19, "Appendix", "Heading 7", 3)),
            ]
        ),
    )
    monkeypatch.setattr(live_outline.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    return document


@pytest.mark.asyncio
async def test_outline_includes_native_and_custom_outline_levels(
    fake_document: SimpleNamespace,
) -> None:
    result = await live_outline.word_live_inspect_document_outline(maximum_level=3)

    assert result["outline_entry_count"] == 2
    assert result["body_paragraph_count"] == 1
    assert [entry["paragraph_index"] for entry in result["entries"]] == [1, 3]
    assert result["entries"][1] == {
        "paragraph_index": 3,
        "outline_level": 3,
        "is_body_text": False,
        "style": "Custom Outline",
        "start_offset": 11,
        "end_offset": 19,
        "page_number": 2,
        "text": "Details",
    }


@pytest.mark.asyncio
async def test_outline_can_include_body_and_validates_level(fake_document: SimpleNamespace) -> None:
    result = await live_outline.word_live_inspect_document_outline(
        maximum_level=1, include_body_text=True
    )
    assert [entry["paragraph_index"] for entry in result["entries"]] == [1, 2]
    with pytest.raises(ValueError, match="between 1 and 9"):
        await live_outline.word_live_inspect_document_outline(maximum_level=10)
