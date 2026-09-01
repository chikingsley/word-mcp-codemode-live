from contextlib import nullcontext
from typing import Any

import pytest

from word_mcp_codemode_live.tools import hyperlinks as hyperlink_tools
from word_mcp_codemode_live.word import session as word_com


class _Range:
    def __init__(self, document: "_Document", start: int, end: int) -> None:
        self._document = document
        self.Start = start
        self.End = end

    @property
    def Text(self) -> str:
        return self._document.text[self.Start : self.End]

    @Text.setter
    def Text(self, value: str) -> None:
        self._document.text = (
            self._document.text[: self.Start] + value + self._document.text[self.End :]
        )
        self.End = self.Start + len(value)

    @property
    def Duplicate(self) -> "_Range":
        return _Range(self._document, self.Start, self.End)


class _Hyperlink:
    def __init__(
        self,
        collection: "_Hyperlinks",
        start: int,
        end: int,
        address: str,
        subaddress: str,
    ) -> None:
        self._collection = collection
        self._start = start
        self._end = end
        self.Address = address
        self.SubAddress = subaddress

    @property
    def Index(self) -> int:
        return self._collection.items.index(self) + 1

    @property
    def Range(self) -> _Range:
        return _Range(self._collection.document, self._start, self._end)

    @property
    def TextToDisplay(self) -> str:
        return self.Range.Text

    @TextToDisplay.setter
    def TextToDisplay(self, value: str) -> None:
        document = self._collection.document
        old_length = self._end - self._start
        document.text = document.text[: self._start] + value + document.text[self._end :]
        self._end = self._start + len(value)
        difference = len(value) - old_length
        for item in self._collection.items:
            if item is not self and item._start >= self._end - difference:
                item._start += difference
                item._end += difference

    def Delete(self) -> None:
        self._collection.items.remove(self)


class _Hyperlinks:
    def __init__(self, document: "_Document") -> None:
        self.document = document
        self.items: list[_Hyperlink] = []
        self.last_add_arguments: dict[str, object] | None = None

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> _Hyperlink:
        return self.items[index - 1]

    def Add(self, **arguments: object) -> _Hyperlink:
        self.last_add_arguments = arguments
        anchor = arguments["Anchor"]
        assert isinstance(anchor, _Range)
        start = anchor.Start
        end = anchor.End
        display_text = arguments.get("TextToDisplay")
        if display_text is not None:
            assert isinstance(display_text, str)
            self.document.text = (
                self.document.text[:start] + display_text + self.document.text[end:]
            )
            end = start + len(display_text)
        hyperlink = _Hyperlink(
            self,
            start,
            end,
            str(arguments["Address"]),
            str(arguments["SubAddress"]),
        )
        self.items.append(hyperlink)
        return hyperlink


class _Document:
    Name = "links.docx"

    def __init__(self, text: str) -> None:
        self.text = text
        self.Hyperlinks = _Hyperlinks(self)

    @property
    def Content(self) -> _Range:
        # Word includes the final paragraph mark in Content.End.
        return _Range(self, 0, len(self.text) + 1)

    def Range(self, start: int, end: int) -> _Range:
        return _Range(self, start, end)


@pytest.fixture
def document(monkeypatch) -> _Document:
    document = _Document("Alpha beta omega")
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    monkeypatch.setattr(word_com, "undo_record", lambda _app, _name: nullcontext())
    return document


@pytest.mark.asyncio
async def test_add_and_list_external_hyperlink_over_existing_range(document: _Document) -> None:
    added = await hyperlink_tools.word_live_add_hyperlink(
        "links.docx",
        address="https://example.com",
        start=6,
        end=10,
    )

    assert document.text == "Alpha beta omega"
    assert added["hyperlink"] == {
        "index": 1,
        "address": "https://example.com",
        "subaddress": "",
        "target_kind": "external",
        "display_text": "beta",
        "range": {"start": 6, "end": 10},
    }

    listed = await hyperlink_tools.word_live_list_hyperlinks()
    assert listed["hyperlink_count"] == 1
    assert listed["hyperlinks"] == [added["hyperlink"]]


@pytest.mark.asyncio
async def test_add_internal_hyperlink_with_display_text_at_document_end(
    document: _Document,
) -> None:
    result = await hyperlink_tools.word_live_add_hyperlink(
        subaddress="DestinationBookmark",
        display_text=" jump",
    )

    assert document.text == "Alpha beta omega jump"
    assert result["hyperlink"]["target_kind"] == "internal"
    assert result["hyperlink"]["range"] == {"start": 16, "end": 21}


@pytest.mark.asyncio
async def test_add_replaces_only_the_requested_display_range(document: _Document) -> None:
    result = await hyperlink_tools.word_live_add_hyperlink(
        address="https://example.com/new",
        display_text="linked phrase",
        start=6,
        end=10,
    )

    assert document.text == "Alpha linked phrase omega"
    assert result["hyperlink"]["display_text"] == "linked phrase"
    assert result["hyperlink"]["range"] == {"start": 6, "end": 19}


@pytest.mark.asyncio
async def test_update_target_and_display_text_by_one_based_index(document: _Document) -> None:
    await hyperlink_tools.word_live_add_hyperlink(
        address="https://old.example",
        start=6,
        end=10,
    )

    result = await hyperlink_tools.word_live_update_hyperlink(
        hyperlink_index=1,
        address="",
        subaddress="LocalTarget",
        display_text="replacement",
    )

    assert document.text == "Alpha replacement omega"
    assert result["hyperlink"] == {
        "index": 1,
        "address": "",
        "subaddress": "LocalTarget",
        "target_kind": "internal",
        "display_text": "replacement",
        "range": {"start": 6, "end": 17},
    }


@pytest.mark.asyncio
async def test_remove_hyperlink_preserves_display_text(document: _Document) -> None:
    await hyperlink_tools.word_live_add_hyperlink(
        address="https://example.com",
        start=6,
        end=10,
    )

    result = await hyperlink_tools.word_live_remove_hyperlink(hyperlink_index=1)

    assert document.text == "Alpha beta omega"
    assert result["removed_hyperlink"]["display_text"] == "beta"
    assert result["remaining_hyperlinks"] == 0


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("arguments", "message"),
    [
        ({}, "Provide address, subaddress, or both"),
        ({"address": "https://example.com", "start": 3}, "start and end"),
        (
            {"address": "https://example.com", "start": 20, "end": 21},
            "Range must satisfy",
        ),
        (
            {"address": "https://example.com", "start": 4, "end": 4},
            "display_text is required",
        ),
    ],
)
async def test_add_rejects_invalid_targets_and_ranges(
    document: _Document,
    arguments: dict[str, Any],
    message: str,
) -> None:
    with pytest.raises(ValueError, match=message):
        await hyperlink_tools.word_live_add_hyperlink(**arguments)


@pytest.mark.asyncio
async def test_update_rejects_invalid_index_and_target_removal(document: _Document) -> None:
    with pytest.raises(ValueError, match="out of range"):
        await hyperlink_tools.word_live_update_hyperlink(
            hyperlink_index=1,
            address="https://example.com",
        )

    await hyperlink_tools.word_live_add_hyperlink(
        address="https://example.com",
        start=6,
        end=10,
    )
    with pytest.raises(ValueError, match="must retain"):
        await hyperlink_tools.word_live_update_hyperlink(
            hyperlink_index=1,
            address="",
        )
