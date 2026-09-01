"""Verify the repository's console entry point over an actual stdio transport."""

import asyncio
from pathlib import Path

from fastmcp import Client
from fastmcp.client.transports import StdioTransport

ROOT = Path(__file__).resolve().parents[2]
UV = Path.home() / ".local" / "bin" / "uv.exe"
CONSOLE_SCRIPT = ROOT / ".venv" / "Scripts" / "word-mcp-codemode-live.exe"


async def main() -> None:
    if CONSOLE_SCRIPT.exists():
        transport = StdioTransport(command=str(CONSOLE_SCRIPT), args=[])
    else:
        transport = StdioTransport(
            command=str(UV),
            args=["run", "--directory", str(ROOT), "word-mcp-codemode-live"],
        )
    async with Client(transport, init_timeout=30) as client:
        tools = await client.list_tools()
        names = {tool.name for tool in tools}
        expected = {
            "search",
            "get_schema",
            "execute",
            "word_live_edit_batch",
            "word_live_capture_pages",
        }
        if names != expected:
            raise RuntimeError(f"Unexpected Code Mode tools: {sorted(names)}")
        print(f"stdio handshake passed; tools={sorted(names)}")


if __name__ == "__main__":
    asyncio.run(main())
