# ty: ignore[unresolved-import, unresolved-attribute]

"""Monkey-patch python-docx Document.save() to preserve custom XML parts.

Problem: python-docx strips parts it doesn't manage (comments.xml, etc.)
Solution: Before save, extract custom parts from the original file. After save,
re-inject them into the new file.
"""

import logging
import zipfile
from io import BytesIO
from pathlib import Path

from lxml import etree

logger = logging.getLogger(__name__)

# Parts that python-docx strips on save
CUSTOM_PARTS_TO_PRESERVE = [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
]

# Relationship types that accompany comments
COMMENT_REL_TYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
}

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def _next_relationship_id(existing_ids: set[str]) -> str:
    numbers = []
    for relationship_id in existing_ids:
        if relationship_id.startswith("rId"):
            try:
                numbers.append(int(relationship_id[3:]))
            except ValueError:
                continue
    return f"rId{max(numbers, default=0) + 1}"


def _patch_relationships(zf_in, preserved: dict, existing_names: set[str]) -> bytes | None:
    rels_path = "word/_rels/document.xml.rels"
    if not preserved["rels"] or rels_path not in existing_names:
        return None
    root = etree.fromstring(zf_in.read(rels_path))
    relationships = list(root.iter(f"{{{REL_NS}}}Relationship"))
    existing_ids = {rel.get("Id", "") for rel in relationships}
    existing_types = {rel.get("Type", "") for rel in relationships}
    for info in preserved["rels"]:
        if info["Type"] in existing_types:
            continue
        relationship_id = info["Id"]
        if relationship_id in existing_ids:
            relationship_id = _next_relationship_id(existing_ids)
        new_relationship = etree.SubElement(root, f"{{{REL_NS}}}Relationship")
        new_relationship.set("Id", relationship_id)
        new_relationship.set("Type", info["Type"])
        new_relationship.set("Target", info["Target"])
        existing_ids.add(relationship_id)
        existing_types.add(info["Type"])
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _patch_content_types(zf_in, preserved: dict, existing_names: set[str]) -> bytes | None:
    content_types_path = "[Content_Types].xml"
    if not preserved["overrides"] or content_types_path not in existing_names:
        return None
    root = etree.fromstring(zf_in.read(content_types_path))
    existing_parts = {
        override.get("PartName", "") for override in root.iter(f"{{{CT_NS}}}Override")
    }
    for info in preserved["overrides"]:
        if info["PartName"] in existing_parts:
            continue
        override = etree.SubElement(root, f"{{{CT_NS}}}Override")
        override.set("PartName", info["PartName"])
        override.set("ContentType", info["ContentType"])
        existing_parts.add(info["PartName"])
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rebuilt_archive(zf_in, preserved: dict, replacements: dict[str, bytes]) -> bytes:
    existing_names = set(zf_in.namelist())
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for item in zf_in.infolist():
            zf_out.writestr(item, replacements.get(item.filename, zf_in.read(item.filename)))
        for part_name, part_bytes in preserved["parts"].items():
            if part_name not in existing_names:
                zf_out.writestr(part_name, part_bytes)
    return buffer.getvalue()


def _extract_custom_parts(zip_bytes: bytes) -> dict | None:
    """Extract custom parts, relationships, and content-type overrides from a docx zip.

    Returns a dict with keys 'parts', 'rels', 'overrides', or None if nothing to preserve.
    """
    parts = {}
    rels = []
    overrides = []

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf:
        namelist = zf.namelist()

        # 1. Extract custom part files
        for part_name in CUSTOM_PARTS_TO_PRESERVE:
            if part_name in namelist:
                parts[part_name] = zf.read(part_name)

        if not parts:
            return None  # Nothing to preserve

        # 2. Extract comment relationships from document.xml.rels
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in namelist:
            rels_root = etree.fromstring(zf.read(rels_path))
            for rel in rels_root.iter(f"{{{REL_NS}}}Relationship"):
                if rel.get("Type", "") in COMMENT_REL_TYPES:
                    rels.append(
                        {
                            "Id": rel.get("Id"),
                            "Type": rel.get("Type"),
                            "Target": rel.get("Target"),
                        }
                    )

        # 3. Extract comment-related content-type overrides
        if "[Content_Types].xml" in namelist:
            ct_root = etree.fromstring(zf.read("[Content_Types].xml"))
            for override in ct_root.iter(f"{{{CT_NS}}}Override"):
                part_name = override.get("PartName", "")
                if "comment" in part_name.lower():
                    overrides.append(
                        {
                            "PartName": part_name,
                            "ContentType": override.get("ContentType"),
                        }
                    )

    return {"parts": parts, "rels": rels, "overrides": overrides}


def _reinject_custom_parts(filepath: Path, preserved: dict) -> None:
    """Re-inject preserved custom parts into a saved docx file."""
    zip_bytes = filepath.read_bytes()

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf_in:
        existing_names = set(zf_in.namelist())
        rels_path = "word/_rels/document.xml.rels"
        replacements = {
            name: content
            for name, content in (
                (rels_path, _patch_relationships(zf_in, preserved, existing_names)),
                (
                    "[Content_Types].xml",
                    _patch_content_types(zf_in, preserved, existing_names),
                ),
            )
            if content is not None
        }
        rebuilt = _rebuilt_archive(zf_in, preserved, replacements)

    filepath.write_bytes(rebuilt)


def install_save_hook() -> None:
    """Monkey-patch docx.document.Document.save to preserve custom XML parts.

    Only intercepts file-path saves (not stream saves).
    Safe to call multiple times — will not double-patch.
    """
    import docx.document

    # Guard against double-patching
    if hasattr(docx.document.Document.save, "_custom_parts_hooked"):
        return

    _original_save = docx.document.Document.save

    def _hooked_save(self, path_or_stream):
        # Only intercept when saving to a file path (str or Path)
        if isinstance(path_or_stream, (str, Path)):
            filepath = Path(path_or_stream)
            preserved = None

            # Extract custom parts from the existing file before python-docx overwrites it
            if filepath.exists():
                try:
                    preserved = _extract_custom_parts(filepath.read_bytes())
                except Exception:
                    preserved = None

            # Let python-docx do its normal save
            _original_save(self, path_or_stream)

            # Re-inject custom parts if we had any
            if preserved is not None:
                try:
                    _reinject_custom_parts(filepath, preserved)
                except Exception as exc:
                    logger.error(
                        "Saved %s but failed to restore preserved comment XML: %s", filepath, exc
                    )
        else:
            # Stream-based save — don't interfere
            _original_save(self, path_or_stream)

    _hooked_save._custom_parts_hooked = True
    docx.document.Document.save = _hooked_save
