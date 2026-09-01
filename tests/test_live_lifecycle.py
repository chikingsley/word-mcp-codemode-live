import json
import os
from types import SimpleNamespace

import pytest

from word_mcp_codemode_live.core import word_com
from word_mcp_codemode_live.tools import lifecycle


@pytest.mark.asyncio
async def test_open_uses_the_shared_word_instance_resolver(tmp_path, monkeypatch) -> None:
    path = tmp_path / "target.docx"
    path.touch()
    opened: list[str] = []
    document = SimpleNamespace(
        Name=path.name,
        FullName=str(path),
        Activate=lambda: None,
    )

    class Documents:
        Count = 0

        def Open(self, filename: str):
            opened.append(filename)
            return document

    app = SimpleNamespace(Documents=Documents(), Visible=False)
    monkeypatch.setattr(lifecycle.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: app)
    monkeypatch.setattr(word_com, "remember_word_app", lambda value: value)

    result = json.loads(await lifecycle.word_live_open(str(path)))

    assert result["success"] is True
    assert opened == [str(path)]
    assert app.Visible is True


@pytest.mark.asyncio
async def test_close_requires_explicit_unsaved_policy(monkeypatch) -> None:
    close_calls: list[int] = []
    document = SimpleNamespace(
        Name="draft.docx",
        FullName=r"C:\Docs\draft.docx",
        Saved=False,
        Close=lambda **kwargs: close_calls.append(kwargs["SaveChanges"]),
    )
    documents = SimpleNamespace(Count=1)
    app = SimpleNamespace(Documents=documents, ActiveDocument=document)
    monkeypatch.setattr(lifecycle.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: app)

    result = json.loads(await lifecycle.word_live_close())

    assert "unsaved changes" in result["error"]
    assert close_calls == []


@pytest.mark.asyncio
async def test_close_can_save_before_closing(monkeypatch) -> None:
    calls: list[str] = []
    documents = SimpleNamespace(Count=1)
    document = SimpleNamespace(
        Name="draft.docx",
        FullName=r"C:\Docs\draft.docx",
        Saved=False,
    )

    def save() -> None:
        calls.append("save")
        document.Saved = True

    def close(**kwargs) -> None:
        calls.append(f"close:{kwargs['SaveChanges']}")
        documents.Count = 0

    document.Save = save
    document.Close = close
    app = SimpleNamespace(Documents=documents, ActiveDocument=document)
    monkeypatch.setattr(lifecycle.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: app)

    result = json.loads(await lifecycle.word_live_close(save_mode="save"))

    assert result["success"] is True
    assert result["remaining_open_documents"] == 0
    assert calls == ["save", "close:0"]


@pytest.mark.asyncio
async def test_rename_moves_open_document_without_leaving_source(tmp_path, monkeypatch) -> None:
    original = tmp_path / "original.docx"
    destination = tmp_path / "renamed.docx"
    original.write_bytes(b"document payload")

    class Document:
        Name = original.name
        FullName = str(original)
        SaveFormat = 16
        Saved = True

        def SaveAs2(self, path: str, **_kwargs) -> None:
            current = type(self).FullName
            with open(current, "rb") as source:
                payload = source.read()
            with open(path, "wb") as target:
                target.write(payload)
            type(self).FullName = path
            type(self).Name = os.path.basename(path)
            type(self).Saved = True

    document = Document()
    documents = SimpleNamespace(Count=1)
    app = SimpleNamespace(Documents=documents, ActiveDocument=document, DisplayAlerts=-1)
    monkeypatch.setattr(lifecycle.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: app)

    result = json.loads(await lifecycle.word_live_rename(str(destination)))

    assert result["success"] is True
    assert result["original_removed"] is True
    assert not original.exists()
    assert destination.read_bytes() == b"document payload"
    assert document.FullName == str(destination)
