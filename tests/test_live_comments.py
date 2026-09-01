from contextlib import nullcontext
from types import SimpleNamespace

import pytest

from word_mcp_codemode_live.tools.comments import word_live_delete_comment
from word_mcp_codemode_live.word import session as word_com


class _Comment:
    Index = 1

    def __init__(self) -> None:
        self.Range = SimpleNamespace(Text="Threaded comment")
        self.Replies = SimpleNamespace(Count=2)
        self.recursively_deleted = False

    def DeleteRecursively(self) -> None:
        self.recursively_deleted = True


class _Comments:
    Count = 1

    def __init__(self, comment: _Comment) -> None:
        self.comment = comment

    def __call__(self, index: int) -> _Comment:
        assert index == 1
        return self.comment


@pytest.mark.asyncio
async def test_delete_comment_removes_threaded_replies(monkeypatch) -> None:
    comment = _Comment()
    document = SimpleNamespace(Name="comments.docx", Comments=_Comments(comment))
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    monkeypatch.setattr(word_com, "undo_record", lambda _app, _name: nullcontext())

    result = await word_live_delete_comment("comments.docx", 1)

    assert result.success is True
    assert result.deleted_replies == 2
    assert comment.recursively_deleted is True
