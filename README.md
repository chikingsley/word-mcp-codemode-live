# word-mcp-codemode-live

Windows-first Code Mode MCP server for editing Microsoft Word documents. It
exposes a compact discovery and execution interface over file-based DOCX tools
and live Word automation instead of loading the entire tool catalog into an
agent context.

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

Live editing requires Microsoft Word on Windows. `pywin32` drives Word, and
PyMuPDF renders Word's PDF output into page images for visual verification.

## Modes

- Code Mode is the default. Clients see `search`, `get_schema`, `execute`, and
  the two image-producing workflow tools (`word_live_edit_batch` and
  `word_live_capture_pages`). The rest of the catalog is discovered on demand.
- `MCP_TOOL_MODE=full` exposes the complete catalog for development and
  diagnostics.
- File-based tools edit saved DOCX files with `python-docx` and OOXML.
- Windows live tools automate an open Word instance through COM.
- `word_live_edit_batch` groups multiple edits into one Undo action, verifies
  text/layout assertions, saves once, and can return rendered affected pages.
- `word_live_capture_pages` returns page images rendered by Word itself.

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

Pushes to `main` run the same checks on Windows. If the version in
`pyproject.toml` does not yet exist on PyPI, the workflow publishes it with
`uv publish` and the repository's `PYPI_TOKEN` Actions secret. Bump the version
before a release:

```bash
uv version --bump patch
git add pyproject.toml uv.lock
git commit -m "release: bump version"
git push origin main
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
