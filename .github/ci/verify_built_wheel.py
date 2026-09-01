"""Install the built wheel in isolation and verify its real console entry point."""

from __future__ import annotations

import argparse
import asyncio
import os
import shutil
import subprocess
import tempfile
from importlib.metadata import version
from pathlib import Path

EXPECTED_CODE_TOOLS = {
    "execute",
    "get_schema",
    "search",
    "word_live_capture_pages",
    "word_live_edit_batch",
}


async def _verify_installed() -> None:
    from fastmcp import Client
    from fastmcp.client.transports import StdioTransport

    import word_mcp_codemode_live
    from word_mcp_codemode_live.main import create_server

    package_path = Path(word_mcp_codemode_live.__file__).resolve()
    repository = Path(__file__).resolve().parents[2]
    if repository in package_path.parents:
        raise RuntimeError(
            f"Smoke test imported the source tree instead of the wheel: {package_path}"
        )

    full_tools = await create_server(tool_mode="full").list_tools()
    if len(full_tools) != 78:
        raise RuntimeError(f"Installed wheel exposed {len(full_tools)} tools instead of 78")

    executable = shutil.which("word-mcp-codemode-live")
    if executable is None:
        raise RuntimeError("Installed wheel did not provide word-mcp-codemode-live")

    transport = StdioTransport(command=executable, args=[])
    async with Client(transport, init_timeout=30) as client:
        names = {tool.name for tool in await client.list_tools()}
    if names != EXPECTED_CODE_TOOLS:
        raise RuntimeError(f"Unexpected installed Code Mode tools: {sorted(names)}")

    print(
        "built-wheel smoke passed; "
        f"package={version('word-mcp-codemode-live')}; "
        f"fastmcp={version('fastmcp')}; full_tools={len(full_tools)}; "
        f"code_tools={sorted(names)}; module={package_path}"
    )


def _run_outer(wheel: Path) -> None:
    uv = shutil.which("uv")
    if uv is None:
        raise RuntimeError("uv is required for the isolated wheel smoke test")
    subprocess.run(
        [
            uv,
            "run",
            "--isolated",
            "--no-project",
            "--directory",
            tempfile.gettempdir(),
            "--with",
            str(wheel.resolve()),
            "python",
            str(Path(__file__).resolve()),
            "--installed",
        ],
        check=True,
        env={**os.environ, "PYTHONPATH": ""},
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--installed", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--wheel", type=Path)
    args = parser.parse_args()
    if args.installed:
        asyncio.run(_verify_installed())
        return

    wheel = args.wheel
    if wheel is None:
        wheels = sorted(Path("dist").glob("word_mcp_codemode_live-*.whl"))
        if len(wheels) != 1:
            raise RuntimeError(f"Expected exactly one built wheel in dist, found {len(wheels)}")
        wheel = wheels[0]
    if not wheel.is_file():
        raise FileNotFoundError(wheel)
    _run_outer(wheel)


if __name__ == "__main__":
    main()
