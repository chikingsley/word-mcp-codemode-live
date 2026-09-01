"""Canonical persistence and semantic comparison for document snapshots."""

import hashlib
import json
import ntpath
import os
import tempfile
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

SCHEMA = "word-mcp-live.document-snapshot"
VERSION = 1
MAX_BYTES = 50 * 1024 * 1024
OFFSET_KEYS = {
    "start_offset",
    "end_offset",
    "reference_offset",
    "code_start_offset",
    "code_end_offset",
    "result_start_offset",
    "result_end_offset",
}


def canonical_bytes(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode(
        "utf-8"
    )


def content_hash(content: dict[str, Any]) -> str:
    return hashlib.sha256(canonical_bytes(content)).hexdigest()


def envelope_hash(snapshot: dict[str, Any]) -> str:
    envelope = {key: value for key, value in snapshot.items() if key != "envelope_sha256"}
    return hashlib.sha256(canonical_bytes(envelope)).hexdigest()


def snapshot_path(path: str) -> Path:
    if not path or not path.strip():
        raise ValueError("snapshot_path is required")
    candidate = Path(path).expanduser().resolve(strict=False)
    if candidate.suffix.casefold() != ".json":
        raise ValueError("snapshot_path must use a .json extension")
    if not candidate.parent.is_dir():
        raise ValueError(f"Snapshot parent directory does not exist: {candidate.parent}")
    if candidate.exists() and not candidate.is_file():
        raise ValueError(f"Snapshot path is not a file: {candidate}")
    return candidate


def write_snapshot(path: Path, snapshot: dict[str, Any], overwrite: bool) -> None:
    serialized = json.dumps(snapshot, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    if len(serialized.encode("utf-8")) > MAX_BYTES:
        raise ValueError(f"Serialized snapshot exceeds the {MAX_BYTES // (1024 * 1024)} MiB limit")
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(serialized)
            handle.flush()
            os.fsync(handle.fileno())
        if overwrite:
            os.replace(temporary_name, path)
        else:
            try:
                # This tool is Windows-only; os.rename publishes atomically and
                # refuses an existing destination instead of replacing it.
                os.rename(temporary_name, path)
            except FileExistsError as exc:
                raise FileExistsError(
                    f"Snapshot already exists: {path}. Set overwrite=true to replace it."
                ) from exc
    except Exception:
        try:
            os.unlink(temporary_name)
        except FileNotFoundError:
            pass
        raise


def _validate_content(content: Any, path: Path) -> dict[str, Any]:
    if not isinstance(content, dict):
        raise ValueError(f"Snapshot content must be a JSON object: {path}")
    list_components = {
        "paragraphs",
        "tables",
        "sections",
        "headers_footers",
        "fields",
        "bookmarks",
        "comments",
        "revisions",
    }
    for component in sorted(list_components):
        records = content.get(component)
        if not isinstance(records, list) or any(not isinstance(item, dict) for item in records):
            raise ValueError(f"Snapshot content.{component} must be a list of objects: {path}")
    notes = content.get("notes")
    if not isinstance(notes, dict) or set(notes) != {"footnotes", "endnotes"}:
        raise ValueError(f"Snapshot content.notes must contain footnotes and endnotes: {path}")
    if any(
        not isinstance(notes[kind], list) or any(not isinstance(item, dict) for item in notes[kind])
        for kind in ("footnotes", "endnotes")
    ):
        raise ValueError(f"Snapshot note collections must be lists of objects: {path}")
    return content


def _validate_source(snapshot: dict[str, Any], path: Path) -> None:
    source = snapshot.get("source")
    if not isinstance(source, dict):
        raise ValueError(f"Snapshot source must be a JSON object: {path}")
    if not isinstance(source.get("full_path"), str) or not source["full_path"]:
        raise ValueError(f"Snapshot source full_path must be a nonempty string: {path}")
    identity = source.get("normalized_path_sha256")
    expected_identity = hashlib.sha256(
        ntpath.normcase(source["full_path"]).encode("utf-8")
    ).hexdigest()
    if identity != expected_identity:
        raise ValueError(f"Snapshot source normalized_path_sha256 is invalid: {path}")


def load_snapshot(path_value: str) -> tuple[Path, dict[str, Any]]:
    path = snapshot_path(path_value)
    if not path.is_file():
        raise FileNotFoundError(f"Snapshot does not exist: {path}")
    if path.stat().st_size > MAX_BYTES:
        raise ValueError(f"Snapshot exceeds the {MAX_BYTES // (1024 * 1024)} MiB limit")
    try:
        snapshot = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read snapshot JSON {path}: {exc}") from exc
    if not isinstance(snapshot, dict):
        raise ValueError(f"Snapshot root must be a JSON object: {path}")
    if snapshot.get("schema") != SCHEMA or snapshot.get("version") != VERSION:
        raise ValueError(
            f"Unsupported snapshot schema/version in {path}; expected {SCHEMA} version {VERSION}"
        )
    if snapshot.get("envelope_sha256") != envelope_hash(snapshot):
        raise ValueError(f"Snapshot envelope hash mismatch: {path}")
    content = _validate_content(snapshot.get("content"), path)
    _validate_source(snapshot, path)
    expected_hash = snapshot.get("content_sha256")
    actual_hash = content_hash(content)
    if expected_hash != actual_hash:
        raise ValueError(f"Snapshot content hash mismatch: {path}")
    return path, snapshot


def without_offsets(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: without_offsets(item) for key, item in value.items() if key not in OFFSET_KEYS}
    if isinstance(value, list):
        return [without_offsets(item) for item in value]
    return value


def _paragraph_semantic(paragraph: dict[str, Any]) -> dict[str, Any]:
    return {
        key: without_offsets(value)
        for key, value in paragraph.items()
        if key not in {"index", *OFFSET_KEYS}
    }


def paragraph_diff(before: list[Any], after: list[Any]) -> list[dict[str, Any]]:
    before_semantic = [_paragraph_semantic(item) for item in before]
    after_semantic = [_paragraph_semantic(item) for item in after]
    matcher = SequenceMatcher(
        None,
        [canonical_bytes(item) for item in before_semantic],
        [canonical_bytes(item) for item in after_semantic],
        autojunk=False,
    )
    operations: list[dict[str, Any]] = []
    for operation, before_start, before_end, after_start, after_end in matcher.get_opcodes():
        if operation == "equal":
            continue
        operations.append(
            {
                "operation": operation,
                "before_start_index": before_start + 1,
                "before_count": before_end - before_start,
                "after_start_index": after_start + 1,
                "after_count": after_end - after_start,
                "before": before_semantic[before_start:before_end],
                "after": after_semantic[after_start:after_end],
            }
        )
    return operations


def _escape_pointer(value: str) -> str:
    return value.replace("~", "~0").replace("/", "~1")


def recursive_diff(before: Any, after: Any, path: str = "") -> list[dict[str, Any]]:
    if type(before) is not type(after):
        return [{"path": path or "/", "operation": "replace", "before": before, "after": after}]
    if isinstance(before, dict):
        changes: list[dict[str, Any]] = []
        for key in sorted(before.keys() | after.keys()):
            child_path = f"{path}/{_escape_pointer(str(key))}"
            if key not in before:
                changes.append({"path": child_path, "operation": "add", "after": after[key]})
            elif key not in after:
                changes.append({"path": child_path, "operation": "remove", "before": before[key]})
            else:
                changes.extend(recursive_diff(before[key], after[key], child_path))
        return changes
    if isinstance(before, list):
        changes = []
        shared = min(len(before), len(after))
        for index in range(shared):
            changes.extend(recursive_diff(before[index], after[index], f"{path}/{index}"))
        for index in range(shared, len(before)):
            changes.append(
                {"path": f"{path}/{index}", "operation": "remove", "before": before[index]}
            )
        for index in range(shared, len(after)):
            changes.append({"path": f"{path}/{index}", "operation": "add", "after": after[index]})
        return changes
    if before != after:
        return [{"path": path or "/", "operation": "replace", "before": before, "after": after}]
    return []
