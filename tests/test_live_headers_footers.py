import json

import pytest

from word_mcp_codemode_live.tools.headers_footers import (
    word_live_edit_headers_footers,
)


@pytest.mark.asyncio
async def test_header_footer_rejects_invalid_color_before_word_mutation() -> None:
    result = json.loads(await word_live_edit_headers_footers(font_color="yellow"))

    assert "six-digit RGB" in result["error"]


@pytest.mark.asyncio
async def test_header_footer_start_number_requires_restart() -> None:
    result = json.loads(
        await word_live_edit_headers_footers(start_at=3, restart_page_numbering=False)
    )

    assert "requires restart_page_numbering=true" in result["error"]
