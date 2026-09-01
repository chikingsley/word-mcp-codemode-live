"""COM connection manager for Microsoft Word on Windows.

Provides functions to connect to a running Word instance and find open documents.
Only works on Windows with pywin32 installed.
"""

import logging
import os
import secrets
import sys
import unicodedata
from collections.abc import Callable
from contextlib import contextmanager
from contextvars import ContextVar
from typing import Any

logger = logging.getLogger(__name__)

_WORD_APP: Any | None = None
_UNDO_DEPTH: ContextVar[int] = ContextVar("word_mcp_undo_depth", default=0)


def require_windows(feature: str = "Word COM automation") -> None:
    """Reject a live Word operation before importing Windows-only dependencies."""
    if sys.platform != "win32":
        raise OSError(f"{feature} is only available on Windows")


def remember_word_app(app: Any) -> Any:
    """Keep the attached Word application alive for the MCP server lifetime."""
    global _WORD_APP
    _WORD_APP = app
    return app


def get_word_app():
    """Get a reference to the running Word application via COM.

    Returns the Word.Application COM object that has open documents.
    When multiple Word instances are running, iterates through all
    Running Object Table (ROT) entries to find one with documents.
    Raises RuntimeError if Word is not running or not on Windows.
    """
    require_windows()

    global _WORD_APP

    if _WORD_APP is not None:
        try:
            _ = _WORD_APP.Documents.Count
            return _WORD_APP
        except Exception:
            _WORD_APP = None

    import win32com.client

    try:
        app = win32com.client.GetActiveObject("Word.Application")
        if app.Documents.Count > 0:
            return remember_word_app(app)
        # GetActiveObject found an empty instance — scan ROT for others
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            return remember_word_app(app_with_docs)
        # No instance has documents; return the empty one (caller may open a file)
        return remember_word_app(app)
    except Exception as exc:
        # GetActiveObject failed entirely — try ROT scan
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            return remember_word_app(app_with_docs)
        raise RuntimeError("Microsoft Word is not running. Please open Word first.") from exc


def _find_word_with_docs():
    """Scan the Running Object Table for a Word.Application with open docs.

    Handles Office 365 / OneDrive scenarios where GetActiveObject returns an
    empty Application proxy.  In these cases, documents are registered in the
    ROT as file monikers (.docx paths or https://d.docs.live.net/... URLs).
    We grab the Document COM object from such a moniker and reach the real
    Application via ``doc.Application``.

    Returns the Word.Application COM object if found, or None.
    """
    try:
        import pythoncom
        import win32com.client

        rot = pythoncom.GetRunningObjectTable(0)
        enum = rot.EnumRunning()

        # Pass 1: look for a Word.Application ROT entry with documents
        monikers_to_retry = []
        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                com_obj = win32com.client.Dispatch(dispatch)
                # Direct Application entry
                if hasattr(com_obj, "Documents") and hasattr(com_obj, "ActiveDocument"):
                    if com_obj.Documents.Count > 0:
                        return com_obj
                # Remember file monikers for pass 2
                if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                    monikers_to_retry.append((name, moniker))
            except Exception:
                # Also collect file monikers we couldn't QI yet
                try:
                    ctx = pythoncom.CreateBindCtx(0)
                    name = moniker.GetDisplayName(ctx, None)
                    if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                        monikers_to_retry.append((name, moniker))
                except Exception as display_exc:
                    logger.debug("Could not inspect fallback ROT moniker: %s", display_exc)
                continue

        # Pass 2: try file monikers → Document → Application
        for _name, moniker in monikers_to_retry:
            try:
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                doc = win32com.client.Dispatch(dispatch)
                app = doc.Application
                if app.Documents.Count > 0:
                    return app
            except Exception as exc:
                logger.debug("Could not bind Word document ROT moniker: %s", exc)
                continue
    except Exception as exc:
        logger.debug("Word Running Object Table scan failed: %s", exc)
    return None


def find_document(app, filename: str | None = None):
    """Find an open document by filename.

    Args:
        app: Word.Application COM object.
        filename: Document name (basename) or full path.
                  If None or empty, returns the active document.

    Returns:
        Document COM object.

    Raises:
        ValueError: If the document is not found or no documents are open.
    """
    if app.Documents.Count == 0:
        raise ValueError("No documents are open in Word")

    if not filename:
        return app.ActiveDocument

    def normalized_path(value: str) -> str:
        return unicodedata.normalize("NFC", os.path.normcase(os.path.normpath(value)))

    if os.path.isabs(filename):
        target_fullpath = normalized_path(filename)
        for i in range(1, app.Documents.Count + 1):
            doc = app.Documents(i)
            if normalized_path(str(doc.FullName)) == target_fullpath:
                return doc
    else:
        target_basename = unicodedata.normalize("NFC", os.path.basename(filename)).casefold()
        matches = [
            app.Documents(i)
            for i in range(1, app.Documents.Count + 1)
            if unicodedata.normalize("NFC", str(app.Documents(i).Name)).casefold()
            == target_basename
        ]
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            candidates = [str(document.FullName) for document in matches]
            raise ValueError(
                f"Document name '{filename}' is ambiguous. Use one of these full paths: {candidates}"
            )

    open_docs = [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]
    raise ValueError(f"Document '{filename}' is not open in Word. Open documents: {open_docs}")


@contextmanager
def undo_record(app, name: str):
    """Wrap a block of COM mutations in a single Word UndoRecord.

    Groups all changes into one Ctrl+Z entry in Word's undo stack.
    The undo record name appears in Edit > Undo and in the undo history.
    Degrades gracefully on Word 2007 or earlier (no UndoRecord support).

    Args:
        app: Word.Application COM object.
        name: Label for the undo entry (truncated to 64 chars by Word).

    Usage::

        with undo_record(app, "MCP: Insert Text"):
            doc.Range(0, 0).InsertBefore("Hello")
    """
    depth = _UNDO_DEPTH.get()
    token = _UNDO_DEPTH.set(depth + 1)

    # Batch tools wrap existing live tools, and those tools already create
    # their own undo records. Word does not support nested custom records, so
    # inner calls participate in the outer record instead of ending it.
    if depth:
        try:
            yield
        finally:
            _UNDO_DEPTH.reset(token)
        return

    rec = None
    try:
        rec = app.UndoRecord
        # Clean up stale undo record from a previous crash/interrupted session
        if rec.IsRecordingCustomRecord:
            try:
                rec.EndCustomRecord()
            except Exception as exc:
                logger.warning("Could not close stale Word undo record: %s", exc)
        rec.StartCustomRecord(name[:64])
    except Exception:
        rec = None  # Word 2007 or earlier — proceed without
    try:
        yield
    finally:
        if rec is not None:
            try:
                rec.EndCustomRecord()
            except Exception as exc:
                logger.warning("Could not close Word undo record %r: %s", name, exc)
        _UNDO_DEPTH.reset(token)


def undo_named_record(document: Any, app: Any, name: str) -> bool:
    """Undo only when Word confirms ``name`` is the latest Undo entry."""
    previous_document = None
    try:
        previous_document = app.ActiveDocument
        document.Activate()
        control = app.CommandBars.FindControl(Type=6, Id=128)
        if control is None or not control.ListCount:
            return False
        latest = str(control.List(1))
        if name[:64].casefold() not in latest.casefold():
            return False
        return bool(document.Undo(1))
    except Exception:
        return False
    finally:
        if previous_document is not None:
            try:
                previous_document.Activate()
            except Exception as exc:
                logger.debug("Could not reactivate previous Word document: %s", exc)


def unique_undo_name(name: str) -> str:
    """Return a human-readable, per-invocation Word Undo label."""
    suffix = f" [{secrets.token_hex(4)}]"
    return f"{name[: 64 - len(suffix)]}{suffix}"


@contextmanager
def revision_tracking(app: Any, document: Any, enabled: bool, author: str):
    """Temporarily enable tracked revisions and restore the user's Word state."""
    previous_tracking = document.TrackRevisions
    previous_author = app.UserName
    if enabled:
        document.TrackRevisions = True
        app.UserName = author
    try:
        yield
    finally:
        if enabled:
            document.TrackRevisions = previous_tracking
            app.UserName = previous_author


@contextmanager
def undo_transaction(
    app: Any,
    document: Any,
    name: str,
    rollback_cleanup: Callable[[], None] | None = None,
):
    """Group edits and roll back a failed top-level custom Undo record.

    Nested uses participate in the outer batch, whose caller owns rollback. For a
    direct tool call, rollback occurs only when Word positively identifies this
    transaction as the latest Undo entry; an unconfirmed rollback is included in
    the raised error instead of risking an older user action. ``rollback_cleanup``
    runs only after that positive Undo and is reserved for Word object definitions
    that the native Undo stack does not restore.
    """
    nested = _UNDO_DEPTH.get() > 0
    actual_name = unique_undo_name(name) if not nested else name
    try:
        with undo_record(app, actual_name):
            yield
    except Exception as exc:
        if nested:
            raise
        if undo_named_record(document, app, actual_name):
            if rollback_cleanup is not None:
                try:
                    rollback_cleanup()
                except Exception as cleanup_exc:
                    raise RuntimeError(
                        f"{exc} (Word undid {name!r}, but rollback cleanup failed: "
                        f"{cleanup_exc}; inspect the document before continuing)"
                    ) from cleanup_exc
            raise
        raise RuntimeError(
            f"{exc} (Word could not confirm automatic rollback of {name!r}; "
            "inspect the document before continuing)"
        ) from exc
