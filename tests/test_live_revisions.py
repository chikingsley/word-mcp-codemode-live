from types import SimpleNamespace

import pytest

from word_mcp_codemode_live.tools import revisions
from word_mcp_codemode_live.word import session as word_session


class _RevisionCollection:
    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int):
        return self._items[index - 1]


@pytest.mark.asyncio
async def test_accept_revisions_returns_typed_result(monkeypatch) -> None:
    accepted: list[str] = []
    items = [
        SimpleNamespace(Author="A", Accept=lambda: accepted.append("first")),
        SimpleNamespace(Author="B", Accept=lambda: accepted.append("second")),
    ]
    document = SimpleNamespace(Name="review.docx", Revisions=_RevisionCollection(items))
    app = SimpleNamespace(ActiveDocument=document, Documents=SimpleNamespace(Count=1))
    monkeypatch.setattr(word_session.sys, "platform", "win32")
    monkeypatch.setattr(word_session, "get_word_app", lambda: app)

    result = await revisions.word_live_accept_revisions(revision_ids=[2])

    assert result.action == "accepted"
    assert result.changed == 1
    assert result.mode == "specific_ids"
    assert accepted == ["second"]


@pytest.mark.asyncio
async def test_revision_ids_must_be_unique_and_in_range(monkeypatch) -> None:
    item = SimpleNamespace(Author="A", Reject=lambda: None)
    document = SimpleNamespace(Name="review.docx", Revisions=_RevisionCollection([item]))
    app = SimpleNamespace(ActiveDocument=document, Documents=SimpleNamespace(Count=1))
    monkeypatch.setattr(word_session.sys, "platform", "win32")
    monkeypatch.setattr(word_session, "get_word_app", lambda: app)

    with pytest.raises(ValueError, match="duplicates"):
        await revisions.word_live_reject_revisions(revision_ids=[1, 1])
    with pytest.raises(ValueError, match="outside"):
        await revisions.word_live_reject_revisions(revision_ids=[2])
