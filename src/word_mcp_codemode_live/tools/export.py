"""Export Word documents to other file formats."""

import os
import sys

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.utils.file_utils import (
    check_file_writeable,
    ensure_docx_extension,
    get_file_lock,
)


@word_tool(title="Convert Word to PDF", domain="export", change="edit")
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
