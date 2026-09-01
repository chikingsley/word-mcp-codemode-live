"""Inspect and control tracked revisions in live Word documents."""

import logging
from typing import Annotated, Any, Literal

from pydantic import BaseModel, Field

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session

logger = logging.getLogger(__name__)

_REVISION_TYPES = {
    1: "insert",
    2: "delete",
    3: "property",
    4: "paragraph_number",
    5: "display_field",
    6: "reconcile",
    7: "conflict",
    8: "style",
    9: "replace",
    10: "section_property",
    11: "table_property",
    12: "cell_insert",
    13: "cell_delete",
    14: "cell_merge",
}
RevisionId = Annotated[int, Field(ge=1)]


class RevisionInfo(BaseModel):
    index: int
    type: str
    type_id: int
    author: str
    date: str
    text: str


class RevisionListResult(BaseModel):
    success: Literal[True] = True
    document: str
    revision_count: int
    revisions: list[RevisionInfo]


class RevisionChangeResult(BaseModel):
    success: Literal[True] = True
    document: str
    action: Literal["accepted", "rejected"]
    changed: int
    mode: str


class TrackChangesResult(BaseModel):
    success: Literal[True] = True
    document: str
    previous_state: bool
    track_changes: bool


def _change_revisions(
    document: Any,
    action: Literal["accept", "reject"],
    author: str | None,
    revision_ids: list[int] | None,
) -> tuple[int, str]:
    """Apply one revision action using a stable, validated selection."""
    count = int(document.Revisions.Count)
    method_name = "Accept" if action == "accept" else "Reject"

    if revision_ids is not None:
        if len(revision_ids) != len(set(revision_ids)):
            raise ValueError("revision_ids must not contain duplicates")
        invalid = sorted(
            revision_id for revision_id in revision_ids if not 1 <= revision_id <= count
        )
        if invalid:
            raise ValueError(f"Revision IDs are outside the current 1..{count} range: {invalid}")
        for revision_id in sorted(revision_ids, reverse=True):
            getattr(document.Revisions(revision_id), method_name)()
        return len(revision_ids), "specific_ids"

    if author is not None:
        changed = 0
        for index in range(count, 0, -1):
            revision = document.Revisions(index)
            if str(revision.Author) == author:
                getattr(revision, method_name)()
                changed += 1
        return changed, f"by_author:{author}"

    getattr(document, f"{method_name}AllRevisions")()
    return count, "all"


@word_tool(title="Word Live List Revisions", domain="revisions", change="read")
async def word_live_list_revisions(filename: str | None = None) -> RevisionListResult:
    """List all tracked changes in an open Word document."""
    word_session.require_windows("Live Word tools")
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)

    revisions: list[RevisionInfo] = []
    for index in range(1, document.Revisions.Count + 1):
        revision = document.Revisions(index)
        text = ""
        try:
            text = revision.Range.Text[:200] if revision.Range and revision.Range.Text else ""
        except Exception as exc:
            logger.debug("Revision text is unavailable for revision %s: %s", index, exc)
        type_id = int(revision.Type)
        revisions.append(
            RevisionInfo(
                index=index,
                type=_REVISION_TYPES.get(type_id, f"unknown({type_id})"),
                type_id=type_id,
                author=str(revision.Author) if revision.Author else "",
                date=str(revision.Date) if revision.Date else "",
                text=text,
            )
        )

    return RevisionListResult(
        document=str(document.Name),
        revision_count=len(revisions),
        revisions=revisions,
    )


async def _apply_revision_change(
    *,
    action: Literal["accept", "reject"],
    filename: str | None,
    author: str | None,
    revision_ids: list[int] | None,
) -> RevisionChangeResult:
    word_session.require_windows("Live Word tools")
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    label = "Accept" if action == "accept" else "Reject"
    with word_session.undo_record(app, f"MCP: {label} Revisions"):
        changed, mode = _change_revisions(document, action, author, revision_ids)
    return RevisionChangeResult(
        document=str(document.Name),
        action="accepted" if action == "accept" else "rejected",
        changed=changed,
        mode=mode,
    )


@word_tool(
    title="Word Live Accept Revisions",
    domain="revisions",
    change="edit",
    batchable=True,
)
async def word_live_accept_revisions(
    filename: str | None = None,
    author: str | None = None,
    revision_ids: list[RevisionId] | None = None,
) -> RevisionChangeResult:
    """Accept all revisions, or only a validated set of IDs or one author's revisions."""
    return await _apply_revision_change(
        action="accept", filename=filename, author=author, revision_ids=revision_ids
    )


@word_tool(
    title="Word Live Reject Revisions",
    domain="revisions",
    change="edit",
    batchable=True,
)
async def word_live_reject_revisions(
    filename: str | None = None,
    author: str | None = None,
    revision_ids: list[RevisionId] | None = None,
) -> RevisionChangeResult:
    """Reject all revisions, or only a validated set of IDs or one author's revisions."""
    return await _apply_revision_change(
        action="reject", filename=filename, author=author, revision_ids=revision_ids
    )


@word_tool(title="Word Live Toggle Track Changes", domain="revisions", change="edit")
async def word_live_toggle_track_changes(
    filename: str | None = None,
    enable: bool | None = None,
) -> TrackChangesResult:
    """Toggle or explicitly set track-changes mode on an open Word document."""
    word_session.require_windows("Live Word editing")
    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    previous = bool(document.TrackRevisions)
    document.TrackRevisions = not previous if enable is None else enable
    return TrackChangesResult(
        document=str(document.Name),
        previous_state=previous,
        track_changes=bool(document.TrackRevisions),
    )
