import gc
import sys
from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from word_mcp_codemode_live.core import word_com
from word_mcp_codemode_live.tools import merge


class _Range:
    def __init__(self, document: "_Document", start: int, end: int) -> None:
        self.document = document
        self.Start = start
        self.End = end
        self.calls: list[tuple[Any, ...]] = []

    def InsertFile(self, *arguments: Any) -> None:
        self.calls.append(arguments)
        self.document.Content.End += 17
        self.document.Styles.add("ImportedStyle")
        self.document.ListTemplates.add("ImportedListTemplate")


class _Style:
    def __init__(self, owner: "_Styles", name: str, *, built_in: bool = False) -> None:
        self.owner = owner
        self.NameLocal = name
        self.BuiltIn = built_in

    def Delete(self) -> None:
        self.owner.items.remove(self)


class _Styles:
    def __init__(self) -> None:
        self.items = [_Style(self, "Normal", built_in=True)]

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, key: int | str) -> _Style:
        if isinstance(key, int):
            return self.items[key - 1]
        return next(item for item in self.items if item.NameLocal == key)

    def add(self, name: str) -> None:
        self.items.append(_Style(self, name))


class _ListTemplate:
    def __init__(self, name: str) -> None:
        self.Name = name


class _ListTemplates:
    def __init__(self) -> None:
        self.items: list[_ListTemplate] = []

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, key: int) -> _ListTemplate:
        return self.items[key - 1]

    def add(self, name: str) -> None:
        self.items.append(_ListTemplate(name))


class _Document:
    Name = "destination.docx"

    def __init__(self, full_name: Path) -> None:
        self.FullName = str(full_name)
        self.Content = SimpleNamespace(End=101, Text="before")
        self.Saved = True
        self.Paragraphs = SimpleNamespace(Count=3)
        self.Sections = SimpleNamespace(Count=1)
        self.Tables = SimpleNamespace(Count=0)
        self.Fields = SimpleNamespace(Count=0)
        self.Comments = SimpleNamespace(Count=0)
        self.Footnotes = SimpleNamespace(Count=0)
        self.Endnotes = SimpleNamespace(Count=0)
        self.Revisions = SimpleNamespace(Count=0)
        self.Styles = _Styles()
        self.ListTemplates = _ListTemplates()
        self.ranges: list[_Range] = []

    def Range(self, start: int, end: int) -> _Range:
        word_range = _Range(self, start, end)
        self.ranges.append(word_range)
        return word_range


@contextmanager
def _transaction(
    _app: Any,
    _document: Any,
    _name: str,
    rollback_cleanup: Any = None,
):
    yield


@pytest.fixture
def live_document(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    destination = tmp_path / "destination.docx"
    destination.touch()
    source = tmp_path / "source.docx"
    source.touch()
    document = _Document(destination)
    monkeypatch.setattr(merge.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    monkeypatch.setattr(word_com, "undo_transaction", _transaction)
    return document, source


@pytest.mark.asyncio
async def test_insert_file_uses_native_word_operation_at_explicit_position(live_document) -> None:
    document, source = live_document

    result = await merge.word_live_insert_file(str(source), position="end")

    assert result["success"] is True
    assert result["native_operation"] == "Range.InsertFile"
    assert result["target"] == {"kind": "position", "position": "end", "start": 100, "end": 100}
    assert result["inserted_range"] == {"start": 100, "end": 117}
    assert result["introduced_styles"] == ["ImportedStyle"]
    assert result["introduced_list_templates"] == ["ImportedListTemplate"]
    assert document.ranges[0].calls == [(str(source.resolve()), "", False, False, False)]


@pytest.mark.asyncio
async def test_insert_file_accepts_explicit_replacement_range_and_source_bookmark(
    live_document,
) -> None:
    document, source = live_document

    result = await merge.word_live_insert_file(
        str(source), target_start=10, target_end=25, source_bookmark="SelectedSource"
    )

    assert result["target"] == {
        "kind": "range",
        "start": 10,
        "end": 25,
        "replaced_characters": 15,
    }
    assert document.ranges[0].calls[0][1:] == ("SelectedSource", False, False, False)


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("arguments", "message"),
    [
        ({}, "Provide exactly one target"),
        ({"position": "middle"}, "position must be"),
        ({"position": "start", "target_start": 1, "target_end": 1}, "exactly one target"),
        ({"target_start": 2}, "provided together"),
        ({"target_start": 20, "target_end": 10}, "less than or equal"),
        ({"target_start": 0, "target_end": 101}, "outside document range"),
    ],
)
async def test_insert_file_rejects_invalid_or_ambiguous_targets(
    live_document, arguments: dict[str, Any], message: str
) -> None:
    _, source = live_document

    with pytest.raises(ValueError, match=message):
        await merge.word_live_insert_file(str(source), **arguments)


@pytest.mark.asyncio
async def test_insert_file_requires_absolute_existing_source(live_document) -> None:
    with pytest.raises(ValueError, match="absolute path"):
        await merge.word_live_insert_file("relative.docx", position="end")

    missing = Path.cwd() / "does-not-exist.docx"
    with pytest.raises(ValueError, match="does not exist"):
        await merge.word_live_insert_file(str(missing), position="end")


@pytest.mark.asyncio
async def test_insert_file_rejects_destination_as_source(live_document) -> None:
    document, _ = live_document

    with pytest.raises(ValueError, match="must not identify the destination"):
        await merge.word_live_insert_file(document.FullName, position="end")


@pytest.mark.asyncio
async def test_insert_file_uses_named_transaction_with_definition_cleanup(
    live_document, monkeypatch: pytest.MonkeyPatch
) -> None:
    document, source = live_document
    calls: list[tuple[Any, Any, str]] = []

    @contextmanager
    def recording_transaction(
        app: Any,
        destination: Any,
        name: str,
        rollback_cleanup: Any = None,
    ):
        calls.append((app, destination, name))
        assert callable(rollback_cleanup)
        yield

    monkeypatch.setattr(word_com, "undo_transaction", recording_transaction)

    await merge.word_live_insert_file(str(source), position="start")

    assert len(calls) == 1
    assert calls[0][1] is document
    assert calls[0][2] == "MCP: Insert File"


def test_failed_insert_cleanup_removes_styles_and_reports_list_template_residue(
    live_document,
) -> None:
    document, _source = live_document
    before_styles = merge._style_inventory(document)
    before_lists = merge._list_template_inventory(document)
    before_structure = merge._structure_counts(document)
    before_text = "before"
    document.Content.Text = before_text
    document.Saved = True
    document.Styles.add("ImportedStyle")
    document.ListTemplates.add("ImportedListTemplate")

    with pytest.raises(RuntimeError, match="cannot delete"):
        merge._rollback_definition_cleanup(
            document,
            before_styles,
            before_lists,
            before_structure,
            before_text,
            True,
        )

    assert "importedstyle" not in merge._style_inventory(document)
    assert merge._introduced_list_templates(document, before_lists) == ["ImportedListTemplate"]


@pytest.mark.skipif(sys.platform != "win32", reason="Microsoft Word COM requires Windows")
@pytest.mark.asyncio
async def test_real_failed_insert_restores_text_styles_structure_and_saved_state(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import win32com.client

    try:
        application = win32com.client.DispatchEx("Word.Application")
    except Exception as exc:
        pytest.skip(f"Microsoft Word is unavailable: {exc}")

    application.Visible = False
    application.DisplayAlerts = 0
    source = None
    destination = None
    try:
        source_path = tmp_path / "rollback-source.docx"
        source = application.Documents.Add()
        source.Content.Text = "Rollback imported content\r"
        imported_style = source.Styles.Add(Name="P2RollbackStyle", Type=1)
        source.Paragraphs(1).Style = imported_style
        source.SaveAs2(str(source_path), AddToRecentFiles=False)
        source.Close(SaveChanges=False)
        source = None

        destination_path = tmp_path / "rollback-destination.docx"
        destination = application.Documents.Add()
        destination.Content.Text = "Destination sentinel\r"
        destination.SaveAs2(str(destination_path), AddToRecentFiles=False)
        destination.Activate()
        word_com.remember_word_app(application)

        before_text = str(destination.Content.Text)
        before_styles = merge._style_inventory(destination)
        before_structure = merge._structure_counts(destination)

        def fail_post_insert(_before: dict[str, int], _after: dict[str, int]) -> dict[str, int]:
            raise RuntimeError("forced post-insert verification failure")

        monkeypatch.setattr(merge, "_count_deltas", fail_post_insert)
        with pytest.raises(RuntimeError, match="forced post-insert verification failure"):
            await merge.word_live_insert_file(
                str(source_path),
                filename=str(destination_path),
                position="end",
            )

        assert str(destination.Content.Text) == before_text
        assert merge._style_inventory(destination) == before_styles
        assert merge._structure_counts(destination) == before_structure
        assert "Rollback imported content" not in str(destination.Content.Text)
        assert bool(destination.Saved) is True
    finally:
        if source is not None:
            source.Close(SaveChanges=False)
            source = None
        if destination is not None:
            destination.Close(SaveChanges=False)
            destination = None
        word_com._WORD_APP = None
        gc.collect()
        application.Quit(SaveChanges=False)
        application = None
