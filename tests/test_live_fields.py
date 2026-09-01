from __future__ import annotations

from contextlib import contextmanager
from typing import Any

import pytest

from word_mcp_codemode_live.tools import fields as live_fields
from word_mcp_codemode_live.word import session as word_com


class FakeTextRange:
    def __init__(self, start: int, end: int, text: str) -> None:
        self.Start = start
        self.End = end
        self.Text = text


class FakeField:
    def __init__(self, field_type: int, start: int, code: str, result: str) -> None:
        self.Type = field_type
        self.Code = FakeTextRange(start, start + len(code), code)
        self.Result = FakeTextRange(start, start + len(result), result)
        self.Locked = False
        self.update_count = 0
        self.unlink_count = 0
        self._collection: FakeFields | None = None

    def Update(self) -> bool:
        self.update_count += 1
        self.Result.Text = f"updated-{self.update_count}"
        return True

    def Unlink(self) -> None:
        self.unlink_count += 1
        if self._collection is not None:
            self._collection.items.remove(self)


class FakeNoResultField:
    Type = 4
    Locked = False

    def __init__(self) -> None:
        self.Code = FakeTextRange(12, 18, ' XE "entry" ')

    @property
    def Result(self) -> FakeTextRange:
        raise RuntimeError("result is unavailable")

    def Update(self) -> bool:
        return True


class FakeFields:
    def __init__(self, *fields: FakeField) -> None:
        self.items: list[Any] = list(fields)
        for field in self.items:
            field._collection = self

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int) -> FakeField:
        return self.items[index - 1]


class FakeStoryRange:
    def __init__(self, fields: FakeFields, next_story: FakeStoryRange | None = None) -> None:
        self.Fields = fields
        self.NextStoryRange = next_story


class FakeStoryRanges:
    def __init__(self, stories: dict[int, FakeStoryRange]) -> None:
        self.stories = stories

    def __call__(self, story_type: int) -> FakeStoryRange:
        if story_type not in self.stories:
            raise RuntimeError("story is absent")
        return self.stories[story_type]


class FakeDocument:
    def __init__(self) -> None:
        self.Name = "fields.docx"
        self.page = FakeField(33, 4, " PAGE ", "1")
        self.primary_header = FakeField(26, 20, " NUMPAGES ", "4")
        self.second_header = FakeField(29, 30, " FILENAME ", "fields.docx")
        linked_header = FakeStoryRange(FakeFields(self.second_header))
        self.StoryRanges = FakeStoryRanges(
            {
                1: FakeStoryRange(FakeFields(self.page)),
                7: FakeStoryRange(FakeFields(self.primary_header), linked_header),
            }
        )


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
async def test_list_fields_traverses_linked_word_stories(fake_document: FakeDocument) -> None:
    result = await live_fields.word_live_list_fields()

    assert result["field_count"] == 3
    assert result["fields"][0] == {
        "index": 1,
        "story": "main_text",
        "story_type_id": 1,
        "story_instance_index": 1,
        "story_field_index": 1,
        "type": "page",
        "type_id": 33,
        "code_start_offset": 4,
        "code_end_offset": 10,
        "result_start_offset": 4,
        "result_end_offset": 5,
        "result_available": True,
        "code": " PAGE ",
        "result": "1",
        "locked": False,
    }
    assert result["fields"][1]["story_instance_index"] == 1
    assert result["fields"][2]["story_instance_index"] == 2


@pytest.mark.asyncio
async def test_update_fields_uses_unique_one_based_indexes(fake_document: FakeDocument) -> None:
    result = await live_fields.word_live_update_fields(field_indices=[3, 1])

    assert result["updated_count"] == 2
    assert [entry["index"] for entry in result["updated_fields"]] == [3, 1]
    assert fake_document.page.update_count == 1
    assert fake_document.primary_header.update_count == 0
    assert fake_document.second_header.update_count == 1

    with pytest.raises(ValueError, match="duplicates"):
        await live_fields.word_live_update_fields(field_indices=[1, 1])
    with pytest.raises(ValueError, match="outside"):
        await live_fields.word_live_update_fields(field_indices=[4])


@pytest.mark.asyncio
async def test_unlink_fields_uses_one_based_indexes_and_preserves_request_order(
    fake_document: FakeDocument,
) -> None:
    result = await live_fields.word_live_unlink_fields(field_indices=[3, 1])

    assert result["unlinked_count"] == 2
    assert [entry["index"] for entry in result["unlinked_fields"]] == [3, 1]
    assert fake_document.page.unlink_count == 1
    assert fake_document.primary_header.unlink_count == 0
    assert fake_document.second_header.unlink_count == 1


@pytest.mark.asyncio
async def test_unlink_fields_rejects_known_unsupported_fields_before_editing(
    fake_document: FakeDocument,
) -> None:
    fake_document.page.Type = 4

    with pytest.raises(ValueError, match="XE"):
        await live_fields.word_live_unlink_fields(field_indices=[1, 3])

    assert fake_document.page.unlink_count == 0
    assert fake_document.second_header.unlink_count == 0


@pytest.mark.asyncio
async def test_list_fields_handles_field_types_without_result_ranges(
    fake_document: FakeDocument,
) -> None:
    no_result = FakeNoResultField()
    main_fields = fake_document.StoryRanges.stories[1].Fields
    main_fields.items.append(no_result)

    result = await live_fields.word_live_list_fields()

    xe = result["fields"][1]
    assert xe["type"] == "index_entry"
    assert xe["result_available"] is False
    assert xe["result"] is None
    with pytest.raises(ValueError, match="XE"):
        await live_fields.word_live_unlink_fields(field_indices=[2])

    updated = await live_fields.word_live_update_fields(field_indices=[2])
    assert updated["updated_fields"][0]["result_available"] is False
    assert updated["updated_fields"][0]["result"] is None
