"""Small, safe lifecycle surface for closed DOCX files."""

import shutil
from pathlib import Path
from typing import Literal

from docx import Document
from pydantic import BaseModel

from word_mcp_codemode_live.filesystem import ensure_docx_extension
from word_mcp_codemode_live.ooxml.documents import get_document_properties
from word_mcp_codemode_live.ooxml.styles import ensure_heading_style, ensure_table_style
from word_mcp_codemode_live.tools.metadata import word_tool


class CreateDocumentResult(BaseModel):
    success: Literal[True] = True
    document: str
    overwritten: bool


class CopyDocumentResult(BaseModel):
    success: Literal[True] = True
    source: str
    destination: str


class DocumentInfoResult(BaseModel):
    success: Literal[True] = True
    document: str
    count_scope: str
    title: str
    author: str
    subject: str
    keywords: str
    created: str
    modified: str
    last_modified_by: str
    revision: int
    section_count: int
    body_paragraph_word_count: int
    body_paragraph_count: int
    body_table_count: int


class DocumentFile(BaseModel):
    path: str
    size_bytes: int


class DocumentListResult(BaseModel):
    success: Literal[True] = True
    directory: str
    recursive: bool
    count: int
    documents: list[DocumentFile]


@word_tool(title="Create Word Document", domain="files", change="edit")
async def create_document(
    filename: str,
    title: str | None = None,
    author: str | None = None,
    overwrite: bool = False,
) -> CreateDocumentResult:
    """Create a DOCX file without silently overwriting an existing document."""
    path = Path(ensure_docx_extension(filename)).resolve()
    existed_before = path.exists()
    if existed_before and not overwrite:
        raise FileExistsError(f"Document already exists: {path}")
    if not path.parent.is_dir():
        raise FileNotFoundError(f"Parent directory does not exist: {path.parent}")

    document = Document()
    if title is not None:
        document.core_properties.title = title
    if author is not None:
        document.core_properties.author = author
    ensure_heading_style(document)
    ensure_table_style(document)
    document.save(str(path))
    return CreateDocumentResult(document=str(path), overwritten=overwrite and existed_before)


@word_tool(title="Copy Word Document", domain="files", change="edit")
async def copy_document(
    source_filename: str,
    destination_filename: str | None = None,
    overwrite: bool = False,
) -> CopyDocumentResult:
    """Copy a DOCX file without silently overwriting the destination."""
    source = Path(ensure_docx_extension(source_filename)).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"Source document does not exist: {source}")

    if destination_filename is None:
        destination = source.with_name(f"{source.stem}_copy{source.suffix}")
    else:
        destination = Path(ensure_docx_extension(destination_filename)).resolve()
    if destination.exists() and not overwrite:
        raise FileExistsError(f"Destination already exists: {destination}")
    if not destination.parent.is_dir():
        raise FileNotFoundError(f"Destination directory does not exist: {destination.parent}")

    shutil.copy2(source, destination)
    return CopyDocumentResult(source=str(source), destination=str(destination))


@word_tool(title="Get DOCX Body Metadata", domain="files", change="read")
async def get_document_info(filename: str) -> DocumentInfoResult:
    """Return closed-file metadata; text counts cover body paragraphs only."""
    path = Path(ensure_docx_extension(filename)).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Document does not exist: {path}")
    properties = get_document_properties(str(path))
    return DocumentInfoResult(
        **properties,
        document=str(path),
        count_scope="top-level body paragraphs and tables only",
    )


@word_tool(title="List Local Word Documents", domain="files", change="read")
async def list_available_documents(
    directory: str = ".", recursive: bool = False
) -> DocumentListResult:
    """List DOCX files deterministically, including uppercase .DOCX names."""
    root = Path(directory).resolve()
    if not root.is_dir():
        raise NotADirectoryError(f"Directory does not exist: {root}")

    iterator = root.rglob("*") if recursive else root.iterdir()
    documents = sorted(
        (path for path in iterator if path.is_file() and path.suffix.casefold() == ".docx"),
        key=lambda path: str(path).casefold(),
    )
    entries = [DocumentFile(path=str(path), size_bytes=path.stat().st_size) for path in documents]
    return DocumentListResult(
        directory=str(root),
        recursive=recursive,
        count=len(entries),
        documents=entries,
    )
