"""Render live Microsoft Word pages through Word's layout engine."""

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pymupdf


@dataclass(frozen=True, slots=True)
class RenderedPage:
    """One Word page rendered to PNG bytes."""

    page: int
    data: bytes
    width: int
    height: int
    output_path: str | None = None


def _normalize_pages(pages: list[int] | None, total_pages: int, current_page: int) -> list[int]:
    requested = pages if pages is not None else [current_page]
    if not requested:
        raise ValueError("pages must contain at least one page number")

    normalized = sorted(set(requested))
    invalid = [page for page in normalized if page < 1 or page > total_pages]
    if invalid:
        raise ValueError(
            f"Page numbers out of range: {invalid}; document contains {total_pages} pages"
        )
    if len(normalized) > 8:
        raise ValueError("At most 8 pages can be returned in one call")
    return normalized


def render_word_pages(
    document: Any,
    pages: list[int] | None = None,
    *,
    dpi: int = 144,
    output_dir: str | None = None,
) -> list[RenderedPage]:
    """Export selected pages through Word's PDF renderer and rasterize them."""
    if dpi < 72 or dpi > 300:
        raise ValueError("dpi must be between 72 and 300")

    total_pages = int(document.ComputeStatistics(2))  # wdStatisticPages
    try:
        current_page = int(document.Application.Selection.Information(3))
    except Exception:
        current_page = 1
    selected_pages = _normalize_pages(pages, total_pages, current_page)
    first_page, last_page = selected_pages[0], selected_pages[-1]

    destination: Path | None = None
    if output_dir:
        destination = Path(output_dir).expanduser().resolve()
        destination.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="word-mcp-render-") as temp_dir:
        pdf_path = os.path.join(temp_dir, "pages.pdf")
        document.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            ExportFormat=17,  # wdExportFormatPDF
            OpenAfterExport=False,
            OptimizeFor=0,  # wdExportOptimizeForPrint
            Range=3,  # wdExportFromTo
            From=first_page,
            To=last_page,
            Item=0,  # wdExportDocumentContent
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )

        rendered: list[RenderedPage] = []
        with pymupdf.open(pdf_path) as pdf:
            for page_number in selected_pages:
                pdf_page = pdf[page_number - first_page]
                pixmap = pdf_page.get_pixmap(dpi=dpi, alpha=False)
                png_bytes = pixmap.tobytes("png")
                saved_path = None
                if destination is not None:
                    page_path = destination / f"word-page-{page_number}.png"
                    page_path.write_bytes(png_bytes)
                    saved_path = str(page_path)
                rendered.append(
                    RenderedPage(
                        page=page_number,
                        data=png_bytes,
                        width=pixmap.width,
                        height=pixmap.height,
                        output_path=saved_path,
                    )
                )
    return rendered
