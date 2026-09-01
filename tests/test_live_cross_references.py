from contextlib import contextmanager
from types import SimpleNamespace
from typing import Any

import pytest

from word_mcp_codemode_live.tools.cross_references import (
    word_live_insert_cross_reference,
    word_live_list_cross_reference_targets,
)
from word_mcp_codemode_live.word import session as word_com


class _Range:
    def __init__(self, start: int, end: int) -> None:
        self.Start = start
        self.End = end
        self.insert_calls: list[dict[str, Any]] = []
        self.collapses: list[int] = []

    @property
    def Duplicate(self) -> "_Range":
        return self

    def Collapse(self, direction: int) -> None:
        self.collapses.append(direction)

    def InsertCrossReference(self, **kwargs: Any) -> None:
        self.insert_calls.append(kwargs)


class _Paragraphs:
    Count = 2

    def __init__(self) -> None:
        self.ranges = {1: _Range(0, 10), 2: _Range(10, 20)}

    def __call__(self, index: int) -> SimpleNamespace:
        return SimpleNamespace(Range=self.ranges[index])


class _Document:
    Name = "references.docx"

    def __init__(self) -> None:
        self.Content = SimpleNamespace(End=101)
        self.Paragraphs = _Paragraphs()
        self.native_items: dict[int, tuple[str, ...]] = {
            1: ("1 Introduction", "2 Scope"),
            2: ("scope_anchor", "appendix_anchor"),
            -1: ("Figure 1 Diagram", "Figure 2 Detail"),
        }
        self.requested_types: list[int] = []
        self.ranges: list[_Range] = []

    def GetCrossReferenceItems(self, reference_type: int) -> tuple[str, ...]:
        self.requested_types.append(reference_type)
        return self.native_items.get(reference_type, ())

    def Range(self, start: int, end: int) -> _Range:
        word_range = _Range(start, end)
        self.ranges.append(word_range)
        return word_range


@pytest.fixture
def fake_word(monkeypatch: pytest.MonkeyPatch) -> tuple[_Document, list[str]]:
    document = _Document()
    undo_names: list[str] = []
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)

    @contextmanager
    def fake_undo(_app: Any, name: str):
        undo_names.append(name)
        yield

    monkeypatch.setattr(word_com, "undo_record", fake_undo)
    return document, undo_names


@pytest.mark.asyncio
async def test_list_targets_uses_native_caption_items(fake_word) -> None:
    document, _undo_names = fake_word

    result = await word_live_list_cross_reference_targets("references.docx", "figure")

    assert result["success"] is True
    assert document.requested_types == [-1]
    assert result["targets"] == [
        {"reference_type": "figure", "target_index": 1, "label": "Figure 1 Diagram"},
        {"reference_type": "figure", "target_index": 2, "label": "Figure 2 Detail"},
    ]
    assert "rediscover" in result["index_stability"]


@pytest.mark.asyncio
async def test_list_wraps_word_com_failure_as_runtime_error(fake_word) -> None:
    document, _undo_names = fake_word

    def fail(_reference_type: int) -> tuple[str, ...]:
        raise OSError("COM disconnected")

    document.GetCrossReferenceItems = fail  # type: ignore[method-assign]

    with pytest.raises(RuntimeError, match="Word could not list.*COM disconnected"):
        await word_live_list_cross_reference_targets(reference_type="heading")


@pytest.mark.asyncio
async def test_insert_bookmark_passes_native_label_and_paragraph_range(fake_word) -> None:
    document, undo_names = fake_word

    result = await word_live_insert_cross_reference(
        "references.docx",
        reference_type="bookmark",
        target_index=2,
        reference_kind="page_number",
        paragraph_index=2,
        insert_as_hyperlink=True,
    )

    paragraph_range = document.Paragraphs.ranges[2]
    assert result["success"] is True
    assert paragraph_range.collapses == [1]
    assert paragraph_range.insert_calls == [
        {
            "ReferenceType": 2,
            "ReferenceKind": 7,
            "ReferenceItem": "appendix_anchor",
            "InsertAsHyperlink": True,
            "IncludePosition": False,
        }
    ]
    assert undo_names == ["MCP: Insert Cross-Reference"]


@pytest.mark.asyncio
async def test_insert_heading_uses_correct_negative_kind_and_document_end(fake_word) -> None:
    document, _undo_names = fake_word

    result = await word_live_insert_cross_reference(
        reference_type="heading",
        target_index=1,
        reference_kind="number_full_context",
        position="end",
    )

    assert result["success"] is True
    assert (document.ranges[0].Start, document.ranges[0].End) == (100, 100)
    assert document.ranges[0].insert_calls[0]["ReferenceType"] == 1
    assert document.ranges[0].insert_calls[0]["ReferenceKind"] == -4
    assert document.ranges[0].insert_calls[0]["ReferenceItem"] == 1


@pytest.mark.asyncio
async def test_insert_rejects_out_of_range_target_before_mutation(fake_word) -> None:
    document, undo_names = fake_word

    with pytest.raises(ValueError, match="between 1 and 2"):
        await word_live_insert_cross_reference(target_index=3)

    assert document.ranges == []
    assert undo_names == []


@pytest.mark.asyncio
async def test_insert_rejects_incompatible_reference_kind(fake_word) -> None:
    document, _undo_names = fake_word

    with pytest.raises(ValueError, match="not valid"):
        await word_live_insert_cross_reference(
            reference_type="figure",
            reference_kind="number_full_context",
        )

    assert document.requested_types == []


@pytest.mark.asyncio
async def test_insert_rejects_unknown_reference_type_cleanly(fake_word) -> None:
    document, _undo_names = fake_word
    unknown_type: Any = "unknown"

    with pytest.raises(ValueError, match="reference_type must be one of"):
        await word_live_insert_cross_reference(reference_type=unknown_type)

    assert document.requested_types == []
