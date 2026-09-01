from contextlib import contextmanager
from types import SimpleNamespace
from typing import Any

import pytest

from word_mcp_codemode_live.tools.comments import word_live_set_comment_status
from word_mcp_codemode_live.word import session as word_com


class _Comment:
    def __init__(self, resolved: bool = False) -> None:
        self.Done = resolved


class _Comments:
    def __init__(self, *comments: Any) -> None:
        self._comments = comments
        self.Count = len(comments)

    def __call__(self, index: int) -> _Comment:
        return self._comments[index - 1]


class _StickyComment:
    @property
    def Done(self) -> bool:
        return False

    @Done.setter
    def Done(self, _value: bool) -> None:
        pass


def _install_document(monkeypatch, comment: object):
    document = SimpleNamespace(Name="comments.docx", Comments=_Comments(comment))
    app = object()
    undo_labels: list[str] = []

    @contextmanager
    def undo_record(_app, label: str):
        undo_labels.append(label)
        yield

    monkeypatch.setattr(word_com, "get_word_app", lambda: app)
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    monkeypatch.setattr(word_com, "undo_record", undo_record)
    return undo_labels


@pytest.mark.asyncio
async def test_resolve_comment_thread(monkeypatch) -> None:
    comment = _Comment(resolved=False)
    undo_labels = _install_document(monkeypatch, comment)

    result = await word_live_set_comment_status("comments.docx", 1, True)

    assert result.model_dump() == {
        "success": True,
        "document": "comments.docx",
        "comment_index": 1,
        "previous_resolved": False,
        "resolved": True,
        "changed": True,
    }
    assert comment.Done is True
    assert len(undo_labels) == 1
    assert undo_labels[0].startswith("MCP: Resolve Comment [")


@pytest.mark.asyncio
async def test_reopen_comment_thread(monkeypatch) -> None:
    comment = _Comment(resolved=True)
    undo_labels = _install_document(monkeypatch, comment)

    result = await word_live_set_comment_status("comments.docx", 1, False)

    assert result.previous_resolved is True
    assert result.resolved is False
    assert result.changed is True
    assert comment.Done is False
    assert len(undo_labels) == 1
    assert undo_labels[0].startswith("MCP: Reopen Comment [")


@pytest.mark.asyncio
async def test_matching_status_is_a_noop(monkeypatch) -> None:
    comment = _Comment(resolved=True)
    undo_labels = _install_document(monkeypatch, comment)

    result = await word_live_set_comment_status("comments.docx", 1, True)

    assert result.resolved is True
    assert result.changed is False
    assert undo_labels == []


@pytest.mark.asyncio
async def test_comment_index_is_one_based(monkeypatch) -> None:
    _install_document(monkeypatch, _Comment())

    with pytest.raises(ValueError, match=r"comment_index 0 out of range \(1-1\)"):
        await word_live_set_comment_status("comments.docx", 0, True)


@pytest.mark.asyncio
async def test_missing_done_property_is_reported(monkeypatch) -> None:
    _install_document(monkeypatch, object())

    with pytest.raises(RuntimeError, match="does not expose the Comment.Done property"):
        await word_live_set_comment_status("comments.docx", 1, True)


@pytest.mark.asyncio
async def test_word_refusing_status_change_is_reported(monkeypatch) -> None:
    _install_document(monkeypatch, _StickyComment())

    with pytest.raises(RuntimeError, match="Word did not apply the requested comment status"):
        await word_live_set_comment_status("comments.docx", 1, True)


def test_status_tool_is_batchable() -> None:
    tool: Any = word_live_set_comment_status
    metadata = tool.__fastmcp__

    assert metadata.tags is not None
    assert {"comments", "edit", "batchable"} <= metadata.tags
