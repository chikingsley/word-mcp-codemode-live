import asyncio

import pytest

from word_mcp_codemode_live.main import ServerConfig, create_server
from word_mcp_codemode_live.registry import TOOL_SPECS
from word_mcp_codemode_live.tools.live_batch_tools import _tool_catalog


def test_server_exposes_code_mode_by_default() -> None:
    tools = asyncio.run(create_server(tool_mode="code").list_tools())
    names = {tool.name for tool in tools}

    assert names == {
        "search",
        "get_schema",
        "execute",
        "word_live_edit_batch",
        "word_live_capture_pages",
    }


def test_full_mode_exposes_stable_tool_surface() -> None:
    tools = asyncio.run(create_server(tool_mode="full").list_tools())
    names = {tool.name for tool in tools}

    assert len(tools) == 118
    assert len(names) == 118
    assert {
        "create_document",
        "word_live_get_text",
        "word_live_edit_batch",
        "word_live_capture_pages",
        "merge_documents",
    } <= names


def test_batch_covers_every_live_mutation_except_transaction_controls() -> None:
    destructive_live = {
        name
        for _function, name, annotations in TOOL_SPECS
        if name.startswith("word_live_") and annotations.destructiveHint
    }
    transaction_controls = {
        "word_live_edit_batch",
        "word_live_save",
        "word_live_undo",
    }

    assert destructive_live <= set(_tool_catalog()) | transaction_controls


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
