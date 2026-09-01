import pytest
from fastmcp import Client

from word_mcp_codemode_live.main import create_server


@pytest.mark.asyncio
async def test_code_mode_discovers_batch_and_page_capture_tools() -> None:
    async with Client(create_server(tool_mode="code")) as client:
        result = await client.call_tool(
            "search", {"query": "atomic batch edit verify screenshot Word pages"}
        )

    text = result.content[0].text
    assert "word_live_edit_batch" in text
    assert "word_live_capture_pages" in text


@pytest.mark.asyncio
@pytest.mark.parametrize(
    ("query", "expected"),
    [
        ("center and bold a heading in the open Word document", "word_live_format_text"),
        (
            "put Roman page numbers in the footer and restart the body at one",
            "word_live_edit_headers_footers",
        ),
        ("create a Word table and format its cells", "word_live_add_table"),
        ("review the comments in this Word document", "word_live_get_comments"),
        ("list the tracked revisions in the open document", "word_live_list_revisions"),
    ],
)
async def test_code_mode_routes_natural_word_requests(query: str, expected: str) -> None:
    async with Client(create_server(tool_mode="code")) as client:
        result = await client.call_tool("search", {"query": query})

    assert expected in result.content[0].text
