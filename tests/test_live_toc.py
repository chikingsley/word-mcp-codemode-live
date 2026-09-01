from contextlib import contextmanager
from typing import Any, cast

import pytest

from word_mcp_codemode_live.tools import toc as live_toc
from word_mcp_codemode_live.word import session as word_com


class FakeCode:
    Text = ' TOC \\o "1-3" \\h '


class FakeField:
    Code = FakeCode()


class FakeFields:
    def __call__(self, index: int) -> FakeField:
        assert index == 1
        return FakeField()


class FakeRange:
    def __init__(self, start: int, end: int | None = None, text: str = "") -> None:
        self.Start = start
        self.End = start if end is None else end
        self.Text = text
        self.Fields = FakeFields()


class FakeToc:
    def __init__(self, collection: "FakeTocs", word_range: FakeRange, **options: Any) -> None:
        self._collection = collection
        self.Range = FakeRange(word_range.Start, word_range.Start + 12, "Heading\t1\r")
        self.UseHeadingStyles = options["UseHeadingStyles"]
        self.UpperHeadingLevel = options["UpperHeadingLevel"]
        self.LowerHeadingLevel = options["LowerHeadingLevel"]
        self.IncludePageNumbers = options["IncludePageNumbers"]
        self.RightAlignPageNumbers = options["RightAlignPageNumbers"]
        self.UseHyperlinks = options["UseHyperlinks"]
        self.HidePageNumbersInWeb = options["HidePageNumbersInWeb"]
        self.update_count = 0
        self.page_update_count = 0

    def Update(self) -> None:
        self.update_count += 1

    def UpdatePageNumbers(self) -> None:
        self.page_update_count += 1

    def Delete(self) -> None:
        self._collection.items.remove(self)


class FakeTocs:
    def __init__(self) -> None:
        self.items: list[FakeToc] = []
        self.last_add: dict[str, Any] | None = None

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> FakeToc:
        return self.items[index - 1]

    def Add(self, **kwargs: Any) -> FakeToc:
        self.last_add = kwargs
        word_range = kwargs.pop("Range")
        item = FakeToc(self, word_range, **kwargs)
        kwargs["Range"] = word_range
        self.items.append(item)
        return item


class FakeParagraph:
    def __init__(self, start: int, end: int) -> None:
        self.Range = FakeRange(start, end)


class FakeParagraphs:
    def __init__(self) -> None:
        self.items = [FakeParagraph(0, 8), FakeParagraph(8, 20)]

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> FakeParagraph:
        return self.items[index - 1]


class FakeDocument:
    def __init__(self) -> None:
        self.Name = "toc.docx"
        self.Content = FakeRange(0, 20)
        self.Paragraphs = FakeParagraphs()
        self.TablesOfContents = FakeTocs()

    def Range(self, start: int, end: int) -> FakeRange:
        return FakeRange(start, end)


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
async def test_native_toc_create_list_update_and_delete(fake_document: FakeDocument) -> None:
    created = await live_toc.word_live_create_table_of_contents(
        target="paragraph_end",
        paragraph_index=2,
        upper_heading_level=2,
        lower_heading_level=5,
        include_page_numbers=False,
    )

    assert created["created_at_offset"] == 19
    assert created["table_of_contents"]["upper_heading_level"] == 2
    assert created["table_of_contents"]["lower_heading_level"] == 5
    assert fake_document.TablesOfContents.last_add is not None
    assert fake_document.TablesOfContents.last_add["Range"].Start == 19
    assert fake_document.TablesOfContents.last_add["UseFields"] is False

    listed = await live_toc.word_live_list_tables_of_contents()
    assert listed["toc_count"] == 1
    assert listed["tables_of_contents"][0]["field_code"].startswith(" TOC")

    updated = await live_toc.word_live_update_table_of_contents(toc_index=1, mode="page_numbers")
    assert updated["mode"] == "page_numbers"
    assert fake_document.TablesOfContents(1).page_update_count == 1

    deleted = await live_toc.word_live_delete_table_of_contents(toc_index=1)
    assert deleted["remaining_toc_count"] == 0


@pytest.mark.asyncio
async def test_toc_target_and_heading_validation(fake_document: FakeDocument) -> None:
    with pytest.raises(ValueError, match="required"):
        await live_toc.word_live_create_table_of_contents(target="paragraph_start")
    with pytest.raises(ValueError, match="only valid"):
        await live_toc.word_live_create_table_of_contents(
            target="document_start", paragraph_index=1
        )
    with pytest.raises(ValueError, match="cannot exceed"):
        await live_toc.word_live_create_table_of_contents(
            upper_heading_level=4, lower_heading_level=3
        )
    with pytest.raises(ValueError, match="target must"):
        await live_toc.word_live_create_table_of_contents(target=cast(Any, "middle"))
    with pytest.raises(ValueError, match="mode must"):
        await live_toc.word_live_update_table_of_contents(mode=cast(Any, "entries"))
