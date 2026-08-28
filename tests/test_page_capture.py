from pathlib import Path
from types import SimpleNamespace

import pymupdf

from word_mcp_codemode_live.tools.page_capture_tools import render_word_pages


class _Selection:
    def Information(self, _kind: int) -> int:
        return 2


class _Document:
    Application = SimpleNamespace(Selection=_Selection())

    def ComputeStatistics(self, _kind: int) -> int:
        return 3

    def ExportAsFixedFormat(self, **kwargs) -> None:
        pdf = pymupdf.open()
        for page_number in range(kwargs["From"], kwargs["To"] + 1):
            page = pdf.new_page(width=612, height=792)
            page.insert_text((72, 72), f"Word page {page_number}")
        pdf.save(kwargs["OutputFileName"])
        pdf.close()


def test_render_word_pages_supports_sparse_pages(tmp_path: Path) -> None:
    rendered = render_word_pages(_Document(), [1, 3], dpi=96, output_dir=str(tmp_path))

    assert [page.page for page in rendered] == [1, 3]
    assert all(page.data.startswith(b"\x89PNG") for page in rendered)
    assert all(page.width == 816 for page in rendered)
    assert (tmp_path / "word-page-1.png").exists()
    assert (tmp_path / "word-page-3.png").exists()


def test_render_word_pages_defaults_to_selection_page() -> None:
    rendered = render_word_pages(_Document(), dpi=72)

    assert [page.page for page in rendered] == [2]
