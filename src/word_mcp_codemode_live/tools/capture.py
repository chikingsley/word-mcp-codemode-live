"""Render selected pages from an open Word document as MCP image content."""

import json
from typing import Any

from fastmcp.tools import ToolResult
from fastmcp.utilities.types import Image
from mcp.types import TextContent

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word.rendering import RenderedPage, render_word_pages


def rendered_pages_result(
    *,
    document_name: str,
    total_pages: int,
    rendered: list[RenderedPage],
    extra: dict[str, Any] | None = None,
) -> ToolResult:
    """Build a mixed text/image MCP result for rendered Word pages."""
    metadata: dict[str, Any] = {
        "success": True,
        "document": document_name,
        "total_pages": total_pages,
        "pages": [
            {
                "page": page.page,
                "width": page.width,
                "height": page.height,
                "path": page.output_path,
            }
            for page in rendered
        ],
    }
    if extra:
        metadata.update(extra)

    content = [
        TextContent(type="text", text=json.dumps(metadata, ensure_ascii=False)),
        *(Image(data=page.data, format="png").to_image_content() for page in rendered),
    ]
    return ToolResult(content=content, structured_content=metadata)


@word_tool(title="Render Live Word Pages", domain="capture", change="read")
async def word_live_capture_pages(
    filename: str | None = None,
    pages: list[int] | None = None,
    dpi: int = 144,
    output_dir: str | None = None,
) -> ToolResult:
    """Render selected pages from an open Word document and return MCP images.

    Args:
        filename: Open document name or full path. Uses the active document when omitted.
        pages: One-based page numbers. Defaults to the page containing Word's selection.
        dpi: PNG resolution from 72 to 300 DPI.
        output_dir: Optional directory in which to retain the PNG files.
    """

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    document.Repaginate()
    total_pages = int(document.ComputeStatistics(2))
    rendered = render_word_pages(document, pages, dpi=dpi, output_dir=output_dir)
    return rendered_pages_result(
        document_name=document.Name,
        total_pages=total_pages,
        rendered=rendered,
    )
