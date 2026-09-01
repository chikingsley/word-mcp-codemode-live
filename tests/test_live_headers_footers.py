import pytest

from word_mcp_codemode_live.tools.headers_footers import (
    word_live_edit_headers_footers,
)


@pytest.mark.asyncio
async def test_header_footer_rejects_invalid_color_before_word_mutation() -> None:
    with pytest.raises(ValueError, match="six-digit RGB"):
        await word_live_edit_headers_footers(font_color="yellow")


@pytest.mark.asyncio
async def test_header_footer_start_number_requires_restart() -> None:
    with pytest.raises(ValueError, match="requires restart_page_numbering=true"):
        await word_live_edit_headers_footers(start_at=3, restart_page_numbering=False)
