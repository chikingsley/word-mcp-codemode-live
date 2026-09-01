"""Inspect and edit comments in live Word documents."""

import logging
from typing import Annotated, Any, Literal

from pydantic import BaseModel, Field

from word_mcp_codemode_live.defaults import DEFAULT_AUTHOR
from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word.ranges import character_or_paragraph_range

logger = logging.getLogger(__name__)

PositiveIndex = Annotated[int, Field(ge=1)]
CharacterOffset = Annotated[int, Field(ge=0)]
NonEmptyText = Annotated[str, Field(min_length=1)]


class CommentReply(BaseModel):
    index: int
    document_index: int
    author: str
    date: str
    text: str


class CommentInfo(BaseModel):
    index: int
    document_index: int
    author: str
    date: str
    text: str
    scope: str
    resolved: bool | None
    replies: list[CommentReply] = Field(default_factory=list)
    reply_count: int = 0


class CommentListResult(BaseModel):
    success: Literal[True] = True
    document: str
    comment_count: int
    raw_comment_count: int
    comments: list[CommentInfo]


class CommentStatusResult(BaseModel):
    success: Literal[True] = True
    document: str
    comment_index: int
    previous_resolved: bool
    resolved: bool
    changed: bool


class CommentAddedResult(BaseModel):
    success: Literal[True] = True
    document: str
    comment_index: int
    document_index: int
    author: str
    text: str


class CommentReplyResult(BaseModel):
    success: Literal[True] = True
    document: str
    comment_index: int
    reply_text: str
    reply_index: int
    reply_document_index: int
    author: str


class CommentDeletedResult(BaseModel):
    success: Literal[True] = True
    document: str
    deleted_comment_index: int
    deleted_comment_text: str
    deleted_replies: int
    remaining_comments: int
    remaining_raw_comments: int


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


def _comment_reply_rows(comment: Any) -> list[CommentReply]:
    replies: list[CommentReply] = []
    try:
        count = int(comment.Replies.Count)
    except Exception as exc:
        logger.debug("Comment replies are unavailable in this Word version: %s", exc)
        return replies
    for reply_index in range(1, count + 1):
        reply = comment.Replies(reply_index)
        replies.append(
            CommentReply(
                index=reply_index,
                document_index=int(reply.Index),
                author=str(reply.Author) if reply.Author else "",
                date=str(reply.Date) if reply.Date else "",
                text=str(reply.Range.Text) if reply.Range and reply.Range.Text else "",
            )
        )
    return replies


@word_tool(title="Word Live Get Comments", domain="comments", change="read")
async def word_live_get_comments(filename: str | None = None) -> CommentListResult:
    """Get all top-level comment threads and their replies from an open document."""
    word_session.require_windows("Live comment tools")
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)

    comments: list[CommentInfo] = []
    for thread_index, (document_index, comment) in enumerate(_top_level_comment_rows(document), 1):
        try:
            scope = comment.Scope.Text[:100] if comment.Scope and comment.Scope.Text else ""
        except Exception as exc:
            logger.debug("Comment scope is unavailable: %s", exc)
            scope = ""
        try:
            resolved = _resolved_state(comment)
        except RuntimeError:
            resolved = None
        replies = _comment_reply_rows(comment)
        comments.append(
            CommentInfo(
                index=thread_index,
                document_index=document_index,
                author=str(comment.Author) if comment.Author else "",
                date=str(comment.Date) if comment.Date else "",
                text=(str(comment.Range.Text) if comment.Range and comment.Range.Text else ""),
                scope=scope,
                resolved=resolved,
                replies=replies,
                reply_count=len(replies),
            )
        )

    return CommentListResult(
        document=str(document.Name),
        comment_count=len(comments),
        raw_comment_count=int(document.Comments.Count),
        comments=comments,
    )


@word_tool(
    title="Word Live Set Comment Status",
    domain="comments",
    change="edit",
    batchable=True,
)
async def word_live_set_comment_status(
    filename: str | None = None,
    comment_index: PositiveIndex = 1,
    resolved: bool = True,
) -> CommentStatusResult:
    """Resolve or reopen one top-level comment thread using Word's native status."""
    word_session.require_windows("Live comment tools")
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    comment = _get_comment(document, comment_index)
    previous_resolved = _resolved_state(comment)
    requested_resolved = bool(resolved)

    if previous_resolved != requested_resolved:
        with word_session.undo_transaction(
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
                raise RuntimeError("Word did not apply the requested comment status")
    else:
        current_resolved = previous_resolved

    return CommentStatusResult(
        document=str(document.Name),
        comment_index=comment_index,
        previous_resolved=previous_resolved,
        resolved=current_resolved,
        changed=previous_resolved != current_resolved,
    )


@word_tool(title="Word Live Add Comment", domain="comments", change="edit", batchable=True)
async def word_live_add_comment(
    filename: str | None = None,
    start: CharacterOffset | None = None,
    end: CharacterOffset | None = None,
    paragraph_index: PositiveIndex | None = None,
    text: NonEmptyText = "",
    author: str = DEFAULT_AUTHOR,
) -> CommentAddedResult:
    """Attach a comment to a character range or one complete paragraph."""
    word_session.require_windows("Live comment tools")
    if not text:
        raise ValueError("Comment text is required")
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    target = character_or_paragraph_range(
        document,
        start=start,
        end=end,
        start_paragraph=paragraph_index,
        end_paragraph=None,
    )

    with word_session.undo_record(app, "MCP: Add Comment"):
        previous_author = app.UserName
        app.UserName = author
        try:
            comment = document.Comments.Add(target.com_range, text)
        finally:
            app.UserName = previous_author

    return CommentAddedResult(
        document=str(document.Name),
        comment_index=len(_top_level_comment_rows(document)),
        document_index=int(comment.Index),
        author=author,
        text=text[:100],
    )


@word_tool(title="Word Live Reply to Comment", domain="comments", change="edit", batchable=True)
async def word_live_reply_to_comment(
    filename: str | None = None,
    comment_index: PositiveIndex | None = None,
    text: NonEmptyText = "",
    author: str = DEFAULT_AUTHOR,
) -> CommentReplyResult:
    """Add a threaded reply to a top-level comment."""
    word_session.require_windows("Live comment tools")
    if comment_index is None:
        raise ValueError("comment_index is required")
    if not text:
        raise ValueError("Reply text is required")
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    comment = _get_comment(document, comment_index)

    with word_session.undo_record(app, "MCP: Reply to Comment"):
        previous_author = app.UserName
        app.UserName = author
        try:
            try:
                reply = comment.Replies.Add(comment.Scope, text)
            except AttributeError as exc:
                raise RuntimeError("Comment replies require Word 2016 or later") from exc
        finally:
            app.UserName = previous_author

    return CommentReplyResult(
        document=str(document.Name),
        comment_index=comment_index,
        reply_text=text[:100],
        reply_index=int(comment.Replies.Count),
        reply_document_index=int(reply.Index),
        author=author,
    )


@word_tool(title="Word Live Delete Comment", domain="comments", change="edit", batchable=True)
async def word_live_delete_comment(
    filename: str | None = None,
    comment_index: PositiveIndex | None = None,
    delete_replies: bool = True,
) -> CommentDeletedResult:
    """Delete one top-level comment, optionally including its reply thread."""
    word_session.require_windows("Live comment tools")
    if comment_index is None:
        raise ValueError("comment_index is required")
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    comment = _get_comment(document, comment_index)
    comment_text = str(comment.Range.Text)[:100] if comment.Range else ""
    try:
        reply_count = int(comment.Replies.Count)
    except Exception:
        reply_count = 0
    if reply_count and not delete_replies:
        raise RuntimeError(
            f"Comment has {reply_count} replies; set delete_replies=true to delete the whole thread"
        )

    with word_session.undo_record(app, "MCP: Delete Comment"):
        if reply_count:
            comment.DeleteRecursively()
        else:
            comment.Delete()

    return CommentDeletedResult(
        document=str(document.Name),
        deleted_comment_index=comment_index,
        deleted_comment_text=comment_text,
        deleted_replies=reply_count,
        remaining_comments=len(_top_level_comment_rows(document)),
        remaining_raw_comments=int(document.Comments.Count),
    )
