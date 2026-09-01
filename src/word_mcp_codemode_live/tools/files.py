"""Small, safe lifecycle surface for closed DOCX files."""

import json
import shutil
from pathlib import Path

from docx import Document

from word_mcp_codemode_live.core.styles import ensure_heading_style, ensure_table_style
from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.utils.document_utils import get_document_properties
from word_mcp_codemode_live.utils.file_utils import ensure_docx_extension


def _result(**values: object) -> str:
    return json.dumps(values, ensure_ascii=False)


@word_tool(title="Create Word Document", domain="files", change="edit")
async def create_document(
    filename: str,
    title: str | None = None,
    author: str | None = None,
    overwrite: bool = False,
) -> str:
    """Create a DOCX file without silently overwriting an existing document."""
    path = Path(ensure_docx_extension(filename)).resolve()
    existed_before = path.exists()
    if existed_before and not overwrite:
        return _result(error=f"Document already exists: {path}")
    if not path.parent.is_dir():
        return _result(error=f"Parent directory does not exist: {path.parent}")

    try:
        document = Document()
        if title is not None:
            document.core_properties.title = title
        if author is not None:
            document.core_properties.author = author
        ensure_heading_style(document)
        ensure_table_style(document)
        document.save(str(path))
        return _result(success=True, document=str(path), overwritten=overwrite and existed_before)
    except Exception as exc:
        return _result(error=f"Failed to create document: {exc}")


@word_tool(title="Copy Word Document", domain="files", change="edit")
async def copy_document(
    source_filename: str,
    destination_filename: str | None = None,
    overwrite: bool = False,
) -> str:
    """Copy a DOCX file without silently overwriting the destination."""
    source = Path(ensure_docx_extension(source_filename)).resolve()
    if not source.is_file():
        return _result(error=f"Source document does not exist: {source}")

    if destination_filename is None:
        destination = source.with_name(f"{source.stem}_copy{source.suffix}")
    else:
        destination = Path(ensure_docx_extension(destination_filename)).resolve()
    if destination.exists() and not overwrite:
        return _result(error=f"Destination already exists: {destination}")
    if not destination.parent.is_dir():
        return _result(error=f"Destination directory does not exist: {destination.parent}")

    try:
        shutil.copy2(source, destination)
        return _result(success=True, source=str(source), destination=str(destination))
    except Exception as exc:
        return _result(error=f"Failed to copy document: {exc}")


@word_tool(title="Get DOCX Body Metadata", domain="files", change="read")
async def get_document_info(filename: str) -> str:
    """Return closed-file metadata; text counts cover body paragraphs only."""
    path = Path(ensure_docx_extension(filename)).resolve()
    if not path.is_file():
        return _result(error=f"Document does not exist: {path}")
    properties = get_document_properties(str(path))
    properties["document"] = str(path)
    properties["count_scope"] = "top-level body paragraphs and tables only"
    return json.dumps(properties, ensure_ascii=False, indent=2)


@word_tool(title="List Local Word Documents", domain="files", change="read")
async def list_available_documents(directory: str = ".", recursive: bool = False) -> str:
    """List DOCX files deterministically, including uppercase .DOCX names."""
    root = Path(directory).resolve()
    if not root.is_dir():
        return _result(error=f"Directory does not exist: {root}")

    iterator = root.rglob("*") if recursive else root.iterdir()
    documents = sorted(
        (path for path in iterator if path.is_file() and path.suffix.casefold() == ".docx"),
        key=lambda path: str(path).casefold(),
    )
    return _result(
        success=True,
        directory=str(root),
        recursive=recursive,
        count=len(documents),
        documents=[{"path": str(path), "size_bytes": path.stat().st_size} for path in documents],
    )
