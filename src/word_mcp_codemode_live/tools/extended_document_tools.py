"""
Extended document tools for Word Document Server.

These tools provide enhanced document content extraction and search capabilities.
"""

import json
import os
import sys

from word_mcp_codemode_live.utils.extended_document_utils import (
    find_text,
    get_highlighted_text,
    get_paragraph_text,
)
from word_mcp_codemode_live.utils.file_utils import (
    check_file_writeable,
    ensure_docx_extension,
    get_file_lock,
)


async def get_paragraph_text_from_document(filename: str, paragraph_index: int) -> str:
    """Get text from a specific paragraph in a Word document.

    Args:
        filename: Path to the Word document
        paragraph_index: Index of the paragraph to retrieve (0-based)
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    if paragraph_index < 0:
        return "Invalid parameter: paragraph_index must be a non-negative integer"

    try:
        result = get_paragraph_text(filename, paragraph_index)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Failed to get paragraph text: {str(e)}"


async def find_text_in_document(
    filename: str, text_to_find: str, match_case: bool = True, whole_word: bool = False
) -> str:
    """Find occurrences of specific text in a Word document.

    Args:
        filename: Path to the Word document
        text_to_find: Text to search for in the document
        match_case: Whether to match case (True) or ignore case (False)
        whole_word: Whether to match whole words only (True) or substrings (False)
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    if not text_to_find:
        return "Search text cannot be empty"

    try:
        result = find_text(filename, text_to_find, match_case, whole_word)
        return json.dumps(result, indent=2)
    except Exception as e:
        return f"Failed to search for text: {str(e)}"


async def get_highlighted_text_from_document(filename: str, color: str | None = None) -> str:
    """Extract all highlighted text from a Word document, including table cells.

    Args:
        filename: Path to the Word document
        color: Optional color filter (e.g. "yellow", "green"). If omitted, returns all.
    """
    filename = ensure_docx_extension(filename)

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    try:
        result = get_highlighted_text(filename, color)
        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return f"Failed to extract highlighted text: {str(e)}"


async def convert_to_pdf(filename: str, output_filename: str | None = None) -> str:
    """Convert a Word document to PDF with Microsoft Word on Windows.

    Args:
        filename: Path to the Word document
        output_filename: Optional path for the output PDF. If not provided,
                         will use the same name with .pdf extension
    """
    if sys.platform != "win32":
        return "PDF conversion requires Microsoft Word on Windows"

    filename = os.path.abspath(ensure_docx_extension(filename))

    if not os.path.exists(filename):
        return f"Document {filename} does not exist"

    if not output_filename:
        base_name, _ = os.path.splitext(filename)
        output_filename = f"{base_name}.pdf"
    elif not output_filename.lower().endswith(".pdf"):
        output_filename = f"{output_filename}.pdf"

    output_filename = os.path.abspath(output_filename)

    output_dir = os.path.dirname(output_filename)
    os.makedirs(output_dir, exist_ok=True)

    is_writeable, error_message = check_file_writeable(output_filename)
    if not is_writeable:
        return f"Cannot create PDF: {error_message} (Path: {output_filename}, Dir: {output_dir})"

    try:
        async with get_file_lock(filename):
            from word_mcp_codemode_live.core.word_com import find_document, get_word_app

            document = None
            dedicated_app = None
            try:
                try:
                    document = find_document(get_word_app(), filename)
                except (RuntimeError, ValueError):
                    import win32com.client

                    dedicated_app = win32com.client.DispatchEx("Word.Application")
                    dedicated_app.Visible = False
                    dedicated_app.DisplayAlerts = 0
                    document = dedicated_app.Documents.Open(filename, ReadOnly=True)

                document.ExportAsFixedFormat(
                    OutputFileName=output_filename,
                    ExportFormat=17,  # wdExportFormatPDF
                    OpenAfterExport=False,
                    OptimizeFor=0,  # wdExportOptimizeForPrint
                    Item=0,  # wdExportDocumentContent
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )
            finally:
                if dedicated_app is not None:
                    if document is not None:
                        document.Close(SaveChanges=False)
                    dedicated_app.Quit(SaveChanges=False)

            if os.path.exists(output_filename) and os.path.getsize(output_filename) > 0:
                return f"Document successfully converted to PDF: {output_filename}"
            return "Microsoft Word completed PDF export but did not create a valid output file"

    except Exception as e:
        return f"Failed to convert document to PDF: {str(e)}"
