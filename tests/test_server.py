import asyncio

import pytest

from word_mcp_codemode_live.main import ServerConfig, create_server


def test_server_exposes_code_mode_by_default() -> None:
    tools = asyncio.run(create_server(tool_mode="code").list_tools())
    names = {tool.name for tool in tools}

    assert names == {"search", "get_schema", "execute"}


def test_full_mode_exposes_stable_tool_surface() -> None:
    tools = asyncio.run(create_server(tool_mode="full").list_tools())
    names = {tool.name for tool in tools}

    assert len(tools) == 117
    assert len(names) == 117
    assert {"create_document", "word_live_get_text", "merge_documents"} <= names


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
