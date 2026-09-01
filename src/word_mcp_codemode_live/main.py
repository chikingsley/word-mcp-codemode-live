"""Application entry point for the Word MCP server."""

import logging
import os
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from fastmcp import FastMCP
from fastmcp.experimental.transforms.code_mode import CodeMode, GetSchemas, Search
from fastmcp.server.providers import FileSystemProvider
from fastmcp.tools import Tool

from word_mcp_codemode_live.ooxml.package_access import install_path_hook
from word_mcp_codemode_live.ooxml.preservation import install_save_hook

Transport = Literal["stdio", "http"]
ToolMode = Literal["code", "full"]
LOGGER = logging.getLogger(__name__)
SERVER_INSTRUCTIONS = """\
Use this server whenever the user asks to inspect, edit, format, review, or render a
Microsoft Word document on this Windows host. Treat "this document", "the open
document", and "the current document" as Word's active document unless the user
identifies another file. When a document task does not name a path, first discover and
call word_live_list_open. If Word has an open document, identify the active document and
inspect that document before editing it. If exactly one document is open, work on it.
Do not create a new document merely because the user omitted a filename. Use Word-live
tools for inspection and editing. For a named closed document, open it with
word_live_open first. Closed-file tools are limited to create, copy, list, body-only
metadata, and PDF export.

For edits, inspect the relevant text, page, structure, or formatting first. Search for
the required operations and fetch their schemas, then prefer one
word_live_edit_batch so edits share one undo record, verification, save, and affected-
page capture. Use word_live_capture_pages when visual evidence is needed without an
edit. Do not use UI or computer automation to read or edit Word content when this
server can perform the operation directly.

If more than one document is open, use the active document unless the user names a
different one; list open documents when the target remains unclear. Preserve content
outside the requested scope. Verify the requested text and layout outcome, report any
rollback or unsupported operation, and never claim success without tool evidence. Use
word_live_close with its safe default when the requested workflow is complete; choose
save or discard only when the user's intent supports it.
"""


class WordCodeMode(CodeMode):
    """Code Mode with direct access to the two image-producing workflows.

    Nested tool calls made through ``execute`` intentionally unwrap results to
    structured data, which drops MCP image content. Exposing these two workflows
    directly keeps the general tool surface compact while allowing clients to receive
    the rendered Word pages.
    """

    DIRECT_TOOLS = frozenset({"word_live_edit_batch", "word_live_capture_pages"})

    async def transform_tools(self, tools: Sequence[Tool]) -> Sequence[Tool]:
        code_tools = list(await super().transform_tools(tools))
        return [*code_tools, *(tool for tool in tools if tool.name in self.DIRECT_TOOLS)]


@dataclass(frozen=True, slots=True)
class ServerConfig:
    """Runtime transport configuration."""

    transport: Transport = "stdio"
    host: str = "127.0.0.1"
    port: int = 8000
    path: str = "/mcp"

    @classmethod
    def from_env(cls) -> "ServerConfig":
        raw_transport = os.getenv("MCP_TRANSPORT", "stdio").lower()
        if raw_transport not in {"stdio", "http"}:
            raise ValueError(f"Unsupported MCP_TRANSPORT={raw_transport!r}; expected stdio or http")

        return cls(
            transport=raw_transport,
            host=os.getenv("MCP_HOST", "127.0.0.1"),
            port=int(os.getenv("MCP_PORT", "8000")),
            path=os.getenv("MCP_PATH", "/mcp"),
        )


def create_server(*, tool_mode: ToolMode | None = None) -> FastMCP:
    """Create the server with Code Mode discovery or the full tool surface."""
    selected_mode = tool_mode or os.getenv("MCP_TOOL_MODE", "code").lower()
    if selected_mode not in {"code", "full"}:
        raise ValueError(f"Unsupported MCP_TOOL_MODE={selected_mode!r}; expected code or full")

    install_save_hook()
    install_path_hook()
    transforms = (
        [WordCodeMode(discovery_tools=[Search(default_limit=10), GetSchemas()])]
        if selected_mode == "code"
        else []
    )
    provider = FileSystemProvider(Path(__file__).parent / "tools")
    return FastMCP(
        "Word MCP CodeMode Live",
        instructions=SERVER_INSTRUCTIONS,
        providers=[provider],
        transforms=transforms,
    )


mcp = create_server()


def run_server(config: ServerConfig | None = None) -> None:
    """Run the server using environment configuration by default."""
    runtime = config or ServerConfig.from_env()
    LOGGER.info("Starting Word MCP server with %s transport", runtime.transport)

    if runtime.transport == "stdio":
        mcp.run()
        return

    mcp.run(
        transport=runtime.transport,
        host=runtime.host,
        port=runtime.port,
        path=runtime.path,
    )


def main() -> None:
    """Console-script entry point."""
    run_server()


if __name__ == "__main__":
    main()
