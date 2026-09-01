"""Inspect and control tracked revisions in live Word documents."""

import json
import logging
import sys

from word_mcp_codemode_live.tools.metadata import word_tool

logger = logging.getLogger(__name__)


@word_tool(title="Word Live List Revisions", domain="revisions", change="read")
async def word_live_list_revisions(filename: str | None = None) -> str:
    """List all tracked changes (revisions) in an open Word document.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with list of revisions (type, author, date, text).
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        # Revision type names
        REV_TYPES = {
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

        revisions = []
        for i in range(1, doc.Revisions.Count + 1):
            rev = doc.Revisions(i)
            rev_text = ""
            try:
                rev_text = rev.Range.Text[:200] if rev.Range and rev.Range.Text else ""
            except Exception as exc:
                logger.debug("Revision text is unavailable for revision %s: %s", i, exc)

            revisions.append(
                {
                    "index": i,
                    "type": REV_TYPES.get(rev.Type, f"unknown({rev.Type})"),
                    "type_id": rev.Type,
                    "author": str(rev.Author) if rev.Author else "",
                    "date": str(rev.Date) if rev.Date else "",
                    "text": rev_text,
                }
            )

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "revision_count": len(revisions),
                "revisions": revisions,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(
    title="Word Live Accept Revisions",
    domain="revisions",
    change="edit",
    batchable=True,
)
async def word_live_accept_revisions(
    filename: str | None = None,
    author: str | None = None,
    revision_ids: list | None = None,
) -> str:
    """Accept tracked changes in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        author: Only accept revisions by this author.
        revision_ids: List of 1-indexed revision IDs to accept. If None + no author, accept all.

    Returns:
        JSON with count of accepted revisions.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Accept Revisions"):
            if revision_ids is not None:
                # Accept specific revisions (process in reverse to preserve indices)
                accepted = 0
                for rid in sorted(revision_ids, reverse=True):
                    if 1 <= rid <= doc.Revisions.Count:
                        doc.Revisions(rid).Accept()
                        accepted += 1
                return json.dumps(
                    {
                        "success": True,
                        "document": doc.Name,
                        "accepted": accepted,
                        "mode": "specific_ids",
                    }
                )

            if author:
                # Accept revisions by author (iterate in reverse)
                accepted = 0
                for i in range(doc.Revisions.Count, 0, -1):
                    rev = doc.Revisions(i)
                    if str(rev.Author) == author:
                        rev.Accept()
                        accepted += 1
                return json.dumps(
                    {
                        "success": True,
                        "document": doc.Name,
                        "accepted": accepted,
                        "mode": f"by_author:{author}",
                    }
                )

            # Accept all
            total = doc.Revisions.Count
            doc.AcceptAllRevisions()
            return json.dumps(
                {
                    "success": True,
                    "document": doc.Name,
                    "accepted": total,
                    "mode": "all",
                }
            )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(
    title="Word Live Reject Revisions",
    domain="revisions",
    change="edit",
    batchable=True,
)
async def word_live_reject_revisions(
    filename: str | None = None,
    author: str | None = None,
    revision_ids: list | None = None,
) -> str:
    """Reject tracked changes in an open Word document.

    Args:
        filename: Document name or path (None = active document).
        author: Only reject revisions by this author.
        revision_ids: List of 1-indexed revision IDs to reject. If None + no author, reject all.

    Returns:
        JSON with count of rejected revisions.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Reject Revisions"):
            if revision_ids is not None:
                rejected = 0
                for rid in sorted(revision_ids, reverse=True):
                    if 1 <= rid <= doc.Revisions.Count:
                        doc.Revisions(rid).Reject()
                        rejected += 1
                return json.dumps(
                    {
                        "success": True,
                        "document": doc.Name,
                        "rejected": rejected,
                        "mode": "specific_ids",
                    }
                )

            if author:
                rejected = 0
                for i in range(doc.Revisions.Count, 0, -1):
                    rev = doc.Revisions(i)
                    if str(rev.Author) == author:
                        rev.Reject()
                        rejected += 1
                return json.dumps(
                    {
                        "success": True,
                        "document": doc.Name,
                        "rejected": rejected,
                        "mode": f"by_author:{author}",
                    }
                )

            total = doc.Revisions.Count
            doc.RejectAllRevisions()
            return json.dumps(
                {
                    "success": True,
                    "document": doc.Name,
                    "rejected": total,
                    "mode": "all",
                }
            )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Toggle Track Changes", domain="revisions", change="edit")
async def word_live_toggle_track_changes(
    filename: str | None = None,
    enable: bool | None = None,
) -> str:
    """Toggle or set track changes mode on an open Word document.

    If enable is omitted, toggles the current state.

    Args:
        filename: Document name or path (None = active document).
        enable: True to enable, False to disable, None to toggle.

    Returns:
        JSON with the new track changes state.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        previous = bool(doc.TrackRevisions)
        if enable is None:
            doc.TrackRevisions = not previous
        else:
            doc.TrackRevisions = enable

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "previous_state": previous,
                "track_changes": bool(doc.TrackRevisions),
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})
