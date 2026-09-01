import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from word_mcp_codemode_live import snapshot_format
from word_mcp_codemode_live.tools import snapshots
from word_mcp_codemode_live.word import session as word_com


class FakeCollection:
    def __init__(self, items=None) -> None:
        self.items = list(items or [])

    @property
    def Count(self) -> int:
        return len(self.items)

    def __call__(self, index: int):
        return self.items[index - 1]


class FakeRange:
    def __init__(self, start: int, text: str, style: str = "Normal") -> None:
        self.Start = start
        self.End = start + len(text) + 1
        self.Text = text + "\r"
        self.Style = SimpleNamespace(NameLocal=style)
        self.ListFormat = SimpleNamespace(ListType=0)

    def Information(self, kind: int) -> bool:
        assert kind == 12
        return False


def fake_document(path: Path, paragraph_texts: list[str]) -> SimpleNamespace:
    offset = 0
    paragraphs = []
    for text in paragraph_texts:
        word_range = FakeRange(offset, text)
        paragraphs.append(SimpleNamespace(Range=word_range, OutlineLevel=10))
        offset = word_range.End
    empty = FakeCollection()
    return SimpleNamespace(
        Name=path.name,
        FullName=str(path),
        Saved=True,
        Paragraphs=FakeCollection(paragraphs),
        Tables=empty,
        Sections=empty,
        Footnotes=empty,
        Endnotes=empty,
        Fields=empty,
        Bookmarks=empty,
        Comments=empty,
        Revisions=empty,
    )


@pytest.mark.asyncio
async def test_create_snapshot_is_versioned_deterministic_and_refuses_overwrite(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    document = fake_document(tmp_path / "source.docx", ["Alpha", "Beta"])
    monkeypatch.setattr(word_com.sys, "platform", "win32")
    monkeypatch.setattr(word_com, "get_word_app", lambda: object())
    monkeypatch.setattr(word_com, "find_document", lambda _app, _filename: document)
    destination = tmp_path / "baseline.json"

    result = await snapshots.word_live_create_document_snapshot(str(destination))
    first_bytes = destination.read_bytes()
    payload = json.loads(first_bytes)

    assert result["success"] is True
    assert payload["schema"] == "word-mcp-live.document-snapshot"
    assert payload["version"] == 1
    assert payload["source"]["full_path"] == str(document.FullName)
    assert "captured_at" not in payload
    assert payload["content"]["paragraphs"][1]["text"] == "Beta"
    assert result["visual_equivalence_captured"] is False
    assert "same care" in result["sensitivity_warning"]

    with pytest.raises(FileExistsError, match="overwrite=true"):
        await snapshots.word_live_create_document_snapshot(str(destination))
    await snapshots.word_live_create_document_snapshot(str(destination), overwrite=True)
    assert destination.read_bytes() == first_bytes


@pytest.mark.asyncio
async def test_diff_uses_sequence_semantics_and_ignores_word_offsets(tmp_path: Path) -> None:
    source = tmp_path / "source.docx"
    before = snapshots._capture(fake_document(source, ["Alpha", "Beta"]))
    after = snapshots._capture(fake_document(source, ["Inserted", "Alpha", "Beta"]))
    # Simulate repagination/range movement without semantic content change.
    after["content"]["paragraphs"][1]["start_offset"] = 500
    after["content"]["paragraphs"][1]["end_offset"] = 506
    after["content_sha256"] = snapshot_format.content_hash(after["content"])
    after["envelope_sha256"] = snapshot_format.envelope_hash(after)
    before_path = tmp_path / "before.json"
    after_path = tmp_path / "after.json"
    snapshot_format.write_snapshot(before_path, before, overwrite=False)
    snapshot_format.write_snapshot(after_path, after, overwrite=False)

    result = await snapshots.word_live_diff_document_snapshots(str(before_path), str(after_path))

    assert result["identical_semantic_content"] is False
    assert result["paragraph_change_group_count"] == 1
    assert result["component_leaf_change_count"] == 0
    assert result["paragraph_operations"] == [
        {
            "operation": "insert",
            "before_start_index": 1,
            "before_count": 0,
            "after_start_index": 1,
            "after_count": 1,
            "before": [],
            "after": [
                {
                    "text": "Inserted",
                    "style": "Normal",
                    "outline_level": 10,
                    "in_table": False,
                    "list": {"type_id": 0, "level": None, "label": None},
                }
            ],
        }
    ]
    assert result["component_changes"] == {}


@pytest.mark.asyncio
async def test_diff_refuses_cross_document_and_detects_tampering(tmp_path: Path) -> None:
    before = snapshots._capture(fake_document(tmp_path / "one.docx", ["Same"]))
    after = snapshots._capture(fake_document(tmp_path / "two.docx", ["Same"]))
    before_path = tmp_path / "before.json"
    after_path = tmp_path / "after.json"
    snapshot_format.write_snapshot(before_path, before, overwrite=False)
    snapshot_format.write_snapshot(after_path, after, overwrite=False)

    with pytest.raises(ValueError, match="different source paths"):
        await snapshots.word_live_diff_document_snapshots(str(before_path), str(after_path))
    allowed = await snapshots.word_live_diff_document_snapshots(
        str(before_path), str(after_path), allow_cross_document=True
    )
    assert allowed["same_source_path"] is False
    assert allowed["identical_semantic_content"] is True
    assert allowed["paragraph_change_group_count"] == 0
    assert allowed["component_leaf_change_count"] == 0

    payload = json.loads(before_path.read_text(encoding="utf-8"))
    payload["capabilities"]["captures"].append("tampered_metadata")
    before_path.write_text(json.dumps(payload), encoding="utf-8")
    with pytest.raises(ValueError, match="envelope hash mismatch"):
        await snapshots.word_live_diff_document_snapshots(
            str(before_path), str(after_path), allow_cross_document=True
        )


def test_snapshot_path_validation_requires_json_and_existing_parent(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match=".json extension"):
        snapshot_format.snapshot_path(str(tmp_path / "snapshot.txt"))
    with pytest.raises(ValueError, match="parent directory"):
        snapshot_format.snapshot_path(str(tmp_path / "missing" / "snapshot.json"))
