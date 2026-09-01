from typing import Any

import pytest

from word_mcp_codemode_live.core.word_com import find_document, undo_record, undo_transaction


class _UndoRecord:
    IsRecordingCustomRecord = False

    def __init__(self, app: "_App | None" = None) -> None:
        self.app = app
        self.started: list[str] = []
        self.ended = 0

    def StartCustomRecord(self, name: str) -> None:
        self.started.append(name)
        self.IsRecordingCustomRecord = True
        if self.app is not None and self.app.CommandBars is not None:
            self.app.CommandBars.menu.label = f"Undo {name}"

    def EndCustomRecord(self) -> None:
        self.ended += 1
        self.IsRecordingCustomRecord = False


class _App:
    def __init__(self) -> None:
        self.CommandBars: Any = None
        self.UndoRecord = _UndoRecord(self)
        self.ActiveDocument: Any = None


class _UndoMenu:
    ListCount = 1

    def __init__(self, label: str) -> None:
        self.label = label

    def List(self, _index: int) -> str:
        return self.label


class _CommandBars:
    def __init__(self, label: str) -> None:
        self.menu = _UndoMenu(label)

    def FindControl(self, **_kwargs):  # type: ignore[no-untyped-def]
        return self.menu


def test_nested_undo_records_join_the_outer_transaction() -> None:
    app = _App()

    with undo_record(app, "outer"):
        with undo_record(app, "inner one"):
            pass
        with undo_record(app, "inner two"):
            pass

    assert app.UndoRecord.started == ["outer"]
    assert app.UndoRecord.ended == 1


def test_failed_direct_transaction_undoes_only_its_named_record() -> None:
    app = _App()
    app.CommandBars = _CommandBars("")
    document = type(
        "Document", (), {"Undo": lambda self, _count: True, "Activate": lambda self: None}
    )()
    app.ActiveDocument = document

    with pytest.raises(ValueError, match="mutation failed"):
        with undo_transaction(app, document, "MCP: Transaction"):
            raise ValueError("mutation failed")


def test_failed_direct_transaction_runs_cleanup_only_after_confirmed_undo() -> None:
    app = _App()
    app.CommandBars = _CommandBars("")
    document = type(
        "Document", (), {"Undo": lambda self, _count: True, "Activate": lambda self: None}
    )()
    app.ActiveDocument = document
    cleanup_calls: list[str] = []

    with pytest.raises(ValueError, match="mutation failed"):
        with undo_transaction(
            app,
            document,
            "MCP: Transaction",
            rollback_cleanup=lambda: cleanup_calls.append("cleaned"),
        ):
            raise ValueError("mutation failed")

    assert cleanup_calls == ["cleaned"]


def test_failed_direct_transaction_reports_unconfirmed_rollback() -> None:
    app = _App()
    app.CommandBars = _CommandBars("Undo Someone else's edit")
    app.UndoRecord.app = None
    document = type(
        "Document", (), {"Undo": lambda self, _count: True, "Activate": lambda self: None}
    )()
    app.ActiveDocument = document

    cleanup_calls: list[str] = []
    with pytest.raises(RuntimeError, match="could not confirm automatic rollback"):
        with undo_transaction(
            app,
            document,
            "MCP: Transaction",
            rollback_cleanup=lambda: cleanup_calls.append("cleaned"),
        ):
            raise ValueError("mutation failed")
    assert cleanup_calls == []


class _Document:
    def __init__(self, full_name: str) -> None:
        self.FullName = full_name
        self.Name = full_name.rsplit("\\", 1)[-1]


class _Documents:
    def __init__(self, *documents: _Document) -> None:
        self._documents = documents
        self.Count = len(documents)

    def __call__(self, index: int) -> _Document:
        return self._documents[index - 1]


def test_find_document_prioritizes_exact_absolute_path() -> None:
    first = _Document(r"C:\First\proposal.docx")
    second = _Document(r"C:\Second\proposal.docx")
    app = type("App", (), {"Documents": _Documents(first, second)})()

    assert find_document(app, second.FullName) is second


def test_find_document_rejects_ambiguous_basename() -> None:
    first = _Document(r"C:\First\proposal.docx")
    second = _Document(r"C:\Second\proposal.docx")
    app = type("App", (), {"Documents": _Documents(first, second)})()

    with pytest.raises(ValueError, match="ambiguous"):
        find_document(app, "proposal.docx")
