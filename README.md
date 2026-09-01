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
- Closed-file tools are limited to create, copy, list, body-only metadata, and
  PDF export. Editing uses Word's native object model.
- Windows live tools automate an open Word instance through COM.
- `word_live_list_footnotes_endnotes` and `word_live_edit_footnotes_endnotes`
  inspect and mutate genuine Word footnotes and endnotes through Word itself.
- `word_live_get_headers_footers` inspects every section and primary/first/even
  story. `word_live_edit_headers_footers` edits one story at a time with page
  fields, linkage, Arabic/Roman/letter numbering, restart values, alignment,
  and font styling.
- `word_live_edit_batch` groups multiple edits into one Undo action, verifies
  factual text/page assertions, saves once, and can return rendered affected
  pages. Rollback occurs only when Word confirms the batch is its latest Undo
  entry.
- `word_live_capture_pages` returns page images rendered by Word itself.

## Reviewed scope

The full development catalog currently contains 78 reviewed tools. The main
live families cover text and paragraph formatting, lists, tables, comments,
tracked revisions, images, equations, native notes, page layout, section
breaks, headers/footers/page numbering, bookmarks, watermarks, Undo, opening,
saving, safe close, verified rename/move operations, and rendered-page
verification.

The server now includes native Word fields, tables of contents, hyperlinks,
cross-reference discovery/insertion, explicit page-break insertion, native heading
numbering, Modern Comment resolve/reopen, custom style management, outline and
highlight inspection, note configuration, objective layout inspection, field
unlinking, native file insertion, persisted semantic snapshots and diffing, and
validated Word UI navigation. Misleading or lossy inherited tools remain absent.

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
uv run python .github/ci/verify_built_wheel.py
```

Run the complete Microsoft Word integration suite on a Windows machine with Word installed:

```powershell
uv run python tests/integration/word_live_suite.py
```

This exhaustive live-Word test uses disposable documents in a private Word instance. By default,
its Markdown and JSON results are written to the system temporary directory, not the repository.

Pull requests and pushes to `main` run the same checks on Windows, including an
isolated install of the built wheel. PyPI publication runs only for a `v*` tag
that exactly matches the version in `pyproject.toml`, using PyPI Trusted
Publishing rather than a stored API token. Configure the `pypi` GitHub
environment and the repository's pending Trusted Publisher on PyPI before the
first release. Then bump and tag the release:

```bash
uv version --bump patch
git add pyproject.toml uv.lock
git commit -m "release: bump version"
git push origin main
git tag "v$(uv version --short)"
git push origin "v$(uv version --short)"
```

Source lives under `src/word_mcp_codemode_live/`. Tool implementations are
grouped by domain in flat modules under `tools/`, lower-level Word and OOXML operations live in
`core/`, and each tool carries its FastMCP metadata beside its implementation.
FastMCP's filesystem provider discovers the tool modules used by Code Mode.

## Acknowledgments

Forked from [ykarapazar/word-mcp-live](https://github.com/ykarapazar/word-mcp-live),
which was built from
[GongRzhe/Office-Word-MCP-Server](https://github.com/GongRzhe/Office-Word-MCP-Server).

## License

MIT. See [LICENSE](LICENSE).
