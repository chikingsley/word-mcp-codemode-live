"""Shared metadata policy for Word MCP tool definitions."""

from collections.abc import Callable
from typing import Any, Literal, TypeVar

from fastmcp.tools import tool
from mcp.types import ToolAnnotations

ToolChange = Literal["read", "safe_write", "edit"]
F = TypeVar("F", bound=Callable[..., Any])


def word_tool(
    *,
    title: str,
    domain: str,
    change: ToolChange,
    batchable: bool = False,
) -> Callable[[F], F]:
    """Attach the complete public MCP contract to a tool function."""
    if batchable and change != "edit":
        raise ValueError("Only reversible document edits can be batchable")

    tags = {domain, change}
    if batchable:
        tags.add("batchable")

    annotations = ToolAnnotations(
        title=title,
        read_only_hint=change == "read",
        destructive_hint=change == "edit",
    )
    decorator = tool(title=title, tags=tags, annotations=annotations)

    def attach(function: F) -> F:
        return decorator(function)

    return attach
