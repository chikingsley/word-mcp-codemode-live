"""Inspect and edit comments in live Word documents."""

import json
import logging
import sys
from typing import Any

from word_mcp_codemode_live.defaults import DEFAULT_AUTHOR
from word_mcp_codemode_live.tools.metadata import word_tool

logger = logging.getLogger(__name__)


def _require_windows() -> None:
    if sys.platform != "win32":
        raise RuntimeError("Live comment tools are only available on Windows")


def _top_level_comment_rows(document: Any) -> list[tuple[int, Any]]:
    rows: list[tuple[int, Any]] = []
    for document_index in range(1, int(document.Comments.Count) + 1):
        comment = document.Comments(document_index)
        try:
            ancestor = comment.Ancestor
        except Exception:
            ancestor = None
        if ancestor is None:
            rows.append((document_index, comment))
    return rows


def _get_comment(document: Any, comment_index: int) -> Any:
    """Return a top-level thread selected by its one-based public index."""
    rows = _top_level_comment_rows(document)
    if comment_index < 1 or comment_index > len(rows):
        raise ValueError(f"comment_index {comment_index} out of range (1-{len(rows)})")
    return rows[comment_index - 1][1]


def _resolved_state(comment: Any) -> bool:
    """Read Word's native closed/resolved flag with an actionable error."""
    try:
        return bool(comment.Done)
    except Exception as exc:
        raise RuntimeError(
            "This Word installation does not expose the Comment.Done property "
            "required to inspect or change a comment thread's resolved state"
        ) from exc


@word_tool(title="Word Live Get Comments", domain="comments", change="read")
async def word_live_get_comments(filename: str | None = None) -> str:
    """Get all comments from an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with list of comments (author, date, text, scope).
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        comments = []
        for thread_index, (document_index, c) in enumerate(_top_level_comment_rows(doc), 1):
            scope_text = ""
            try:
                scope_text = c.Scope.Text[:100] if c.Scope and c.Scope.Text else ""
            except Exception as exc:
                logger.debug("Comment scope is unavailable: %s", exc)

            # Collect replies (Word 2016+)
            replies = []
            try:
                for r_idx in range(1, c.Replies.Count + 1):
                    r = c.Replies(r_idx)
                    replies.append(
                        {
                            "index": r_idx,
                            "document_index": int(r.Index),
                            "author": str(r.Author) if r.Author else "",
                            "date": str(r.Date) if r.Date else "",
                            "text": str(r.Range.Text) if r.Range and r.Range.Text else "",
                        }
                    )
            except Exception as exc:
                logger.debug("Comment replies are unavailable in this Word version: %s", exc)

            comment_data = {
                "index": thread_index,
                "document_index": document_index,
                "author": str(c.Author) if c.Author else "",
                "date": str(c.Date) if c.Date else "",
                "text": str(c.Range.Text) if c.Range and c.Range.Text else "",
                "scope": scope_text,
            }
            try:
                comment_data["resolved"] = _resolved_state(c)
            except RuntimeError:
                # Older Word object models can still expose comments without Done.
                # Keep comment inspection useful while making the missing state explicit.
                comment_data["resolved"] = None
            if replies:
                comment_data["replies"] = replies
                comment_data["reply_count"] = len(replies)
            comments.append(comment_data)

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "comment_count": len(comments),
                "raw_comment_count": int(doc.Comments.Count),
                "comments": comments,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(
    title="Word Live Set Comment Status",
    domain="comments",
    change="edit",
    batchable=True,
)
async def word_live_set_comment_status(
    filename: str | None = None,
    comment_index: int = 1,
    resolved: bool = True,
) -> dict[str, Any]:
    """Resolve or reopen one top-level comment thread using Word's native status.

    The comment index is one-based and comes from ``word_live_get_comments``.
    ``resolved=True`` marks the thread resolved; ``resolved=False`` reopens it.
    This intentionally targets the top-level thread. Microsoft documents that
    setting ``Comment.Done`` on an individual reply has no visible effect in the
    redesigned comments experience.

    Args:
        filename: Open document name or full path (None selects the active document).
        comment_index: One-based top-level comment index.
        resolved: True to resolve the thread, or False to reopen it.
    """
    _require_windows()

    from word_mcp_codemode_live.core.word_com import (
        find_document,
        get_word_app,
        undo_transaction,
    )

    app = get_word_app()
    document = find_document(app, filename)
    comment = _get_comment(document, comment_index)
    previous_resolved = _resolved_state(comment)
    requested_resolved = bool(resolved)

    if previous_resolved != requested_resolved:
        with undo_transaction(
            app,
            document,
            "MCP: Resolve Comment" if requested_resolved else "MCP: Reopen Comment",
        ):
            try:
                comment.Done = requested_resolved
            except Exception as exc:
                raise RuntimeError(
                    "This Word installation does not support changing a comment "
                    "thread's resolved state through Comment.Done"
                ) from exc
            current_resolved = _resolved_state(comment)
            if current_resolved != requested_resolved:
                raise RuntimeError(
                    "Word did not apply the requested comment status; the thread state "
                    "was left unchanged"
                )
    else:
        current_resolved = previous_resolved

    return {
        "success": True,
        "document": str(document.Name),
        "comment_index": comment_index,
        "previous_resolved": previous_resolved,
        "resolved": current_resolved,
        "changed": previous_resolved != current_resolved,
    }


@word_tool(title="Word Live Add Comment", domain="comments", change="edit", batchable=True)
async def word_live_add_comment(
    filename: str | None = None,
    start: int | None = None,
    end: int | None = None,
    paragraph_index: int | None = None,
    text: str = "",
    author: str = DEFAULT_AUTHOR,
) -> str:
    """Add a comment to an open Word document.

    Specify either start/end character positions or paragraph_index.
    If paragraph_index is given, the comment is attached to the entire paragraph.

    Args:
        filename: Document name or path (None = active document).
        start: Start character position.
        end: End character position.
        paragraph_index: 1-indexed paragraph to attach comment to.
        text: Comment text.
        author: Comment author name.

    Returns:
        JSON with result info.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if not text:
        return json.dumps({"error": "Comment text is required"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        # Determine the range to attach the comment to
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                return json.dumps(
                    {
                        "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                    }
                )
            rng = doc.Paragraphs(paragraph_index).Range
        elif start is not None and end is not None:
            rng = doc.Range(start, end)
        else:
            return json.dumps({"error": "Provide either start/end positions or paragraph_index"})

        with undo_record(app, "MCP: Add Comment"):
            # Save and restore author
            prev_author = app.UserName
            app.UserName = author
            try:
                comment = doc.Comments.Add(rng, text)
            finally:
                app.UserName = prev_author

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "comment_index": len(_top_level_comment_rows(doc)),
                "document_index": int(comment.Index),
                "author": author,
                "text": text[:100],
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Reply to Comment", domain="comments", change="edit", batchable=True)
async def word_live_reply_to_comment(
    filename: str | None = None,
    comment_index: int | None = None,
    text: str = "",
    author: str = DEFAULT_AUTHOR,
) -> str:
    """Reply to an existing comment in an open Word document.

    Adds a threaded reply to a top-level comment. Requires Word 2016 or later.
    Use word_live_get_comments to find the comment_index.

    Args:
        filename: Document name or path (None = active document).
        comment_index: 1-indexed comment to reply to.
        text: Reply text.
        author: Reply author name.

    Returns:
        JSON with reply info.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if comment_index is None:
        return json.dumps({"error": "comment_index is required"})
    if not text:
        return json.dumps({"error": "Reply text is required"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        try:
            comment = _get_comment(doc, comment_index)
        except ValueError as exc:
            return json.dumps({"error": str(exc)})

        with undo_record(app, "MCP: Reply to Comment"):
            prev_author = app.UserName
            app.UserName = author
            try:
                reply = comment.Replies.Add(comment.Scope, text)
            except AttributeError:
                return json.dumps({"error": "Comment replies require Word 2016 or later."})
            finally:
                app.UserName = prev_author

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "comment_index": comment_index,
                "reply_text": text[:100],
                "reply_index": int(comment.Replies.Count),
                "reply_document_index": int(reply.Index),
                "author": author,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Delete Comment", domain="comments", change="edit", batchable=True)
async def word_live_delete_comment(
    filename: str | None = None,
    comment_index: int | None = None,
    delete_replies: bool = True,
) -> str:
    """Delete a comment from an open Word document.

    Args:
        filename: Document name or path (None = active document).
        comment_index: 1-indexed comment to delete.
        delete_replies: Delete the entire reply thread when replies exist.

    Returns:
        JSON with result info.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    if comment_index is None:
        return json.dumps({"error": "comment_index is required"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        try:
            comment = _get_comment(doc, comment_index)
        except ValueError as exc:
            return json.dumps({"error": str(exc)})
        comment_text = str(comment.Range.Text)[:100] if comment.Range else ""
        try:
            reply_count = int(comment.Replies.Count)
        except Exception:
            reply_count = 0
        if reply_count and not delete_replies:
            return json.dumps(
                {
                    "error": f"Comment has {reply_count} replies; set delete_replies=true "
                    "to delete the whole thread"
                }
            )

        with undo_record(app, "MCP: Delete Comment"):
            if reply_count:
                comment.DeleteRecursively()
            else:
                comment.Delete()

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "deleted_comment_index": comment_index,
                "deleted_comment_text": comment_text,
                "deleted_replies": reply_count,
                "remaining_comments": len(_top_level_comment_rows(doc)),
                "remaining_raw_comments": int(doc.Comments.Count),
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})
