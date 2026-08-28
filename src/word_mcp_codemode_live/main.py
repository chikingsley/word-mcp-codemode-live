"""Application entry point for the Word MCP server."""

import logging
import os
from dataclasses import dataclass
from typing import Literal

from fastmcp import FastMCP
from fastmcp.experimental.transforms.code_mode import CodeMode, GetSchemas, Search

from word_mcp_codemode_live.registry import register_tools
from word_mcp_codemode_live.utils.path_utils import install_path_hook
from word_mcp_codemode_live.utils.save_utils import install_save_hook

Transport = Literal["stdio", "http"]
ToolMode = Literal["code", "full"]
LOGGER = logging.getLogger(__name__)


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
    """Create the server with Code Mode discovery or the full tool catalog."""
    selected_mode = tool_mode or os.getenv("MCP_TOOL_MODE", "code").lower()
    if selected_mode not in {"code", "full"}:
        raise ValueError(f"Unsupported MCP_TOOL_MODE={selected_mode!r}; expected code or full")

    install_save_hook()
    install_path_hook()
    transforms = (
        [CodeMode(discovery_tools=[Search(default_limit=5), GetSchemas()])]
        if selected_mode == "code"
        else []
    )
    return register_tools(FastMCP("Word MCP CodeMode Live", transforms=transforms))


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
