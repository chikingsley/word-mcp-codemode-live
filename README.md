# word-mcp-codemode-live

Code Mode MCP server for editing Microsoft Word documents. It exposes a compact
discovery and execution interface over file-based DOCX tools and live Microsoft
Word automation instead of loading the entire tool catalog into an agent context.

This project is currently beta software.

## Install

Python 3.10 or newer is required. Run the published package in an isolated
environment with `uvx`:

```json
{
  "mcpServers": {
    "word": {
      "command": "uvx",
      "args": ["word-mcp-codemode-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

To install the command permanently:

```bash
uv tool install word-mcp-codemode-live
word-mcp-codemode-live
```

Windows live editing installs `pywin32`; Word-window screenshots also install
Pillow. Both dependencies are limited to Windows by environment markers.

## Modes

- Code Mode is the default. Clients see only `search`, `get_schema`, and
  `execute`, while the underlying Word tool catalog is discovered on demand.
- `MCP_TOOL_MODE=full` exposes the complete catalog for development and
  diagnostics.
- Cross-platform tools edit saved DOCX files with `python-docx` and OOXML.
- Windows live tools automate an open Word instance through COM.
- macOS live tools automate Microsoft Word through JavaScript for Automation.
- Linux supports file-based DOCX tools; live Word automation is unavailable.

## Transports

The server supports two transports:

- `stdio` is the default and the normal choice for a local MCP client.
- `http` is opt-in for clients that need a network endpoint.

HTTP binds to loopback by default. This server can read and modify local files
and open Word documents, so do not expose it to an untrusted network without an
authentication boundary.

## Configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| `MCP_AUTHOR` | `Author` | Author for comments and tracked changes |
| `MCP_AUTHOR_INITIALS` | empty | Comment author initials |
| `MCP_TOOL_MODE` | `code` | `code` for Code Mode or `full` for the raw catalog |
| `MCP_TRANSPORT` | `stdio` | `stdio` or `http` |
| `MCP_HOST` | `127.0.0.1` | HTTP bind address |
| `MCP_PORT` | `8000` | HTTP bind port |
| `MCP_PATH` | `/mcp` | HTTP endpoint path |

## Development

```bash
git clone https://github.com/chikingsley/word-mcp-codemode-live.git
cd word-mcp-codemode-live
uv sync

uv run ruff check .
uv run ruff format --check .
uv run ty check
uv run vulture
uv run pytest
uv build --no-sources
```

Pushes to `main` run the same checks in GitHub Actions. If the version in
`pyproject.toml` does not yet exist on PyPI, the workflow publishes it with
PyPI Trusted Publishing. Bump the version before a release:

```bash
uv version --bump patch
git add pyproject.toml uv.lock
git commit -m "release: bump version"
git push
```

Source lives under `src/word_mcp_codemode_live/`. Tool implementations are
grouped by domain in `tools/`, lower-level Word and OOXML operations live in
`core/`, and `registry.py` defines the internal tool catalog used by Code Mode.

## Acknowledgments

Forked from [ykarapazar/word-mcp-live](https://github.com/ykarapazar/word-mcp-live),
which was built from
[GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server).

## License

MIT. See [LICENSE](LICENSE).
