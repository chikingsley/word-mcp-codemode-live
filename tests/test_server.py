import asyncio

import pytest
from fastmcp import Client
from fastmcp.exceptions import ToolError

from word_mcp_codemode_live.main import SERVER_INSTRUCTIONS, ServerConfig, create_server
from word_mcp_codemode_live.tools.batch import _batch_tools

DOMAINS = {
    "batch",
    "capture",
    "comments",
    "content",
    "export",
    "files",
    "formatting",
    "headers_footers",
    "inspection",
    "layout",
    "lifecycle",
    "notes",
    "numbering",
    "references",
    "revisions",
    "styles",
    "tables",
}
CHANGE_KINDS = {"read", "safe_write", "edit"}


def test_server_exposes_code_mode_by_default() -> None:
    server = create_server(tool_mode="code")
    tools = asyncio.run(server.list_tools())
    names = {tool.name for tool in tools}

    assert server.instructions == SERVER_INSTRUCTIONS
    assert "Word's active document" in server.instructions
    assert "first discover and" in server.instructions
    assert "word_live_list_open" in server.instructions
    assert "Do not create a new document" in server.instructions
    assert "word_live_edit_batch" in server.instructions
    assert names == {
        "search",
        "get_schema",
        "execute",
        "word_live_edit_batch",
        "word_live_capture_pages",
    }


@pytest.mark.asyncio
async def test_server_advertises_word_workflow_instructions() -> None:
    # FastMCP 4 defaults to the sessionless protocol, which has no initialize
    # result. Exercise the legacy handshake explicitly because that is where MCP
    # server instructions are transported.
    async with Client(create_server(tool_mode="code"), mode="legacy") as client:
        initialize_result = client.initialize_result

    assert initialize_result is not None
    assert initialize_result.instructions == SERVER_INSTRUCTIONS


def test_full_mode_exposes_stable_tool_surface() -> None:
    tools = asyncio.run(create_server(tool_mode="full").list_tools())
    names = {tool.name for tool in tools}

    assert len(tools) == 78
    assert len(names) == 78
    assert {
        "create_document",
        "word_live_get_text",
        "word_live_edit_batch",
        "word_live_capture_pages",
        "word_live_list_footnotes_endnotes",
        "word_live_edit_footnotes_endnotes",
        "word_live_get_headers_footers",
        "word_live_edit_headers_footers",
        "word_live_insert_page_break",
        "word_live_list_fields",
        "word_live_update_fields",
        "word_live_list_tables_of_contents",
        "word_live_create_table_of_contents",
        "word_live_list_hyperlinks",
        "word_live_add_hyperlink",
        "word_live_list_cross_reference_targets",
        "word_live_insert_cross_reference",
        "word_live_inspect_heading_numbering",
        "word_live_setup_heading_numbering",
        "word_live_set_comment_status",
        "word_live_list_custom_styles",
        "word_live_create_custom_style",
        "word_live_update_custom_style",
        "word_live_delete_custom_style",
        "word_live_inspect_document_outline",
        "word_live_inspect_highlighted_text",
        "word_live_get_note_configuration",
        "word_live_set_note_configuration",
        "word_live_inspect_layout",
        "word_live_unlink_fields",
        "word_live_insert_file",
        "word_live_create_document_snapshot",
        "word_live_diff_document_snapshots",
        "word_live_navigate",
        "word_live_close",
        "word_live_rename",
        "convert_to_pdf",
    } <= names
    assert {
        "add_table",
        "format_text",
        "format_table",
        "add_comment",
        "set_page_layout",
        "merge_documents",
        "add_table_of_contents",
        "manage_hyperlinks",
        "word_live_take_snapshot",
        "word_live_diagnose_layout",
    }.isdisjoint(names)


def test_batch_covers_every_live_mutation_except_transaction_controls() -> None:
    tools = asyncio.run(create_server(tool_mode="full").list_tools())
    destructive_live = {
        tool.name
        for tool in tools
        if tool.name.startswith("word_live_")
        and tool.annotations is not None
        and tool.annotations.destructive_hint
    }
    deliberately_non_batchable = {
        "word_live_edit_batch",
        "word_live_save",
        "word_live_undo",
        "word_live_close",
        "word_live_rename",
        # These application/document settings are not reliably restored by
        # Word's document Undo stack.
        "word_live_toggle_track_changes",
        "word_live_set_core_properties",
        # Word Undo removes inserted content but not imported style definitions;
        # the direct tool performs verified rollback cleanup that cannot nest.
        "word_live_insert_file",
        # Snapshot creation writes or explicitly replaces a caller-selected JSON
        # file; it is not a reversible Word document edit.
        "word_live_create_document_snapshot",
    }

    assert destructive_live <= set(_batch_tools()) | deliberately_non_batchable


def test_batchable_registration_tags_match_batch_dispatch() -> None:
    tools = asyncio.run(create_server(tool_mode="full").list_tools())
    tagged = {tool.name for tool in tools if "batchable" in tool.tags}

    assert tagged == set(_batch_tools())


def test_every_tool_has_one_domain_and_change_kind() -> None:
    tools = asyncio.run(create_server(tool_mode="full").list_tools())

    for tool in tools:
        assert tool.annotations is not None, tool.name
        assert len(tool.tags & DOMAINS) == 1, tool.name
        assert len(tool.tags & CHANGE_KINDS) == 1, tool.name


def test_formatting_tools_publish_typed_contracts() -> None:
    tools = asyncio.run(create_server(tool_mode="full").list_tools())
    by_name = {tool.name: tool for tool in tools}

    inspect_schema = by_name["word_live_get_paragraph_format"].output_schema
    edit_schema = by_name["word_live_format_text"].output_schema
    assert inspect_schema is not None
    assert inspect_schema["properties"]["paragraphs"]["type"] == "array"
    assert edit_schema is not None
    assert edit_schema["properties"]["success"]["const"] is True
    assert edit_schema["required"] == ["document", "range", "text_preview", "tracked"]


@pytest.mark.asyncio
async def test_formatting_validation_is_an_mcp_error() -> None:
    async with Client(create_server(tool_mode="full")) as client:
        with pytest.raises(ToolError) as caught:
            await client.call_tool(
                "word_live_format_text",
                {"start": -1, "end": 3, "paragraph_alignment": "diagonal"},
            )

    assert "greater than or equal to 0" in str(caught.value)
    assert "left" in str(caught.value)


def test_http_transport(monkeypatch) -> None:
    monkeypatch.setenv("MCP_TRANSPORT", "http")
    monkeypatch.delenv("MCP_HOST", raising=False)

    config = ServerConfig.from_env()

    assert config.transport == "http"
    assert config.host == "127.0.0.1"


def test_legacy_transport_is_rejected(monkeypatch) -> None:
    monkeypatch.setenv("MCP_TRANSPORT", "sse")

    with pytest.raises(ValueError, match="expected stdio or http"):
        ServerConfig.from_env()
