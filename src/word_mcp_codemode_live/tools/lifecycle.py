"""Open, save, inspect, and undo live Word document state."""

import json
import os
import sys
from typing import Literal

from word_mcp_codemode_live.tools.metadata import word_tool

_CORE_PROP_MAP = {
    "title": "Title",
    "subject": "Subject",
    "author": "Author",
    "keywords": "Keywords",
    "comments": "Comments",
    "category": "Category",
    "manager": "Manager",
    "company": "Company",
    "last_author": "Last Author",
}


@word_tool(title="Word Live Get Undo History", domain="lifecycle", change="read")
async def word_live_get_undo_history(
    filename: str | None = None,
) -> str:
    """[Windows only] Get the undo stack names from an open Word document.

    Uses Word's CommandBars to read the undo dropdown list. Each MCP tool call
    that was wrapped with undo_record will appear as "MCP: <tool name>".
    Degrades gracefully if the undo list is not accessible.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with undo_entries list (most recent first) and count.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        entries = []
        try:
            # CommandBar control ID 128 = Undo split dropdown (Type=6)
            # Must specify Type=6 (msoControlSplitDropdown) — without it,
            # FindControl may return a plain button (Type=1) that lacks ListCount.
            undo_control = app.CommandBars.FindControl(Type=6, Id=128)
            if undo_control is not None:
                for i in range(1, undo_control.ListCount + 1):
                    entries.append(undo_control.List(i))
        except Exception:
            # Undocumented API — may not be available in all Word versions
            return json.dumps(
                {
                    "success": True,
                    "document": doc.Name,
                    "undo_entries": [],
                    "count": 0,
                    "note": "Undo history not accessible in this Word version",
                }
            )

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "undo_entries": entries,
                "count": len(entries),
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Set Core Properties", domain="lifecycle", change="edit")
async def word_live_set_core_properties(
    filename: str | None = None,
    title: str | None = None,
    subject: str | None = None,
    author: str | None = None,
    keywords: str | None = None,
    comments: str | None = None,
    category: str | None = None,
    manager: str | None = None,
    company: str | None = None,
    last_author: str | None = None,
) -> str:
    """[Windows only] Set Word document core/built-in properties (Title, Subject, Author, etc.).

    Equivalent to File > Info > Properties in the Word UI. Pass None for any
    field to leave it unchanged. Wrapped in undo_record so a single Ctrl+Z
    reverts every property in the call.

    Args:
        filename: Document name or path (None = active document).
        title: Document Title.
        subject: Document Subject.
        author: Author (current author / "Created by").
        keywords: Keywords (semicolon-separated by Word convention).
        comments: Free-form Comments.
        category: Category.
        manager: Manager.
        company: Company.
        last_author: "Last Author" (Last saved by).

    Returns:
        JSON {ok, document, changed: {field: {old, new}}, errors: {field: msg}}.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import (
            find_document,
            get_word_app,
            undo_record,
        )

        app = get_word_app()
        doc = find_document(app, filename)

        inputs = {
            "title": title,
            "subject": subject,
            "author": author,
            "keywords": keywords,
            "comments": comments,
            "category": category,
            "manager": manager,
            "company": company,
            "last_author": last_author,
        }

        changes: dict = {}
        errors: dict = {}

        with undo_record(app, "MCP: Set Core Properties"):
            props = doc.BuiltInDocumentProperties
            for key, value in inputs.items():
                if value is None:
                    continue
                prop_name = _CORE_PROP_MAP[key]
                try:
                    raw_old = props(prop_name).Value
                    old = str(raw_old) if raw_old is not None else None
                except Exception:
                    old = None
                try:
                    props(prop_name).Value = value
                    changes[key] = {"old": old, "new": value}
                except Exception as e:
                    errors[key] = str(e)

        return json.dumps(
            {
                "ok": len(errors) == 0,
                "document": doc.Name,
                "changed": changes,
                "errors": errors or None,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live List Open", domain="lifecycle", change="read")
async def word_live_list_open() -> str:
    """[Windows only] List all documents currently open in Microsoft Word.

    Returns JSON with list of open documents including name, full_path,
    pages, saved status, and whether it is the active document.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live tools are only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import get_word_app

        app = get_word_app()

        # ActiveDocument access can throw on broken/proxy state — degrade gracefully.
        try:
            active_fullname = app.ActiveDocument.FullName if app.Documents.Count > 0 else None
        except Exception:
            active_fullname = None

        try:
            count = app.Documents.Count
        except Exception as e:
            return json.dumps(
                {
                    "error": f"could not enumerate Documents collection: {e}",
                    "documents": [],
                }
            )

        documents = []
        for i in range(1, count + 1):
            entry = {"index": i}
            try:
                doc = app.Documents(i)
            except Exception as e:
                entry.update(
                    {
                        "name": "<unavailable>",
                        "error": f"could not access Documents({i}): {e}",
                    }
                )
                documents.append(entry)
                continue

            # Defensive per-property access — one broken doc must not
            # block reporting on healthy ones. Each failure is recorded
            # under entry["errors"] but does not abort the loop.
            errors = []

            def _get(attr, transform=None, *, _doc=doc, _errors=errors):
                try:
                    val = getattr(_doc, attr)
                    return transform(val) if transform else val
                except Exception as ex:
                    _errors.append(f"{attr}: {ex}")
                    return None

            entry["name"] = _get("Name")
            entry["full_path"] = _get("FullName")
            entry["saved"] = _get("Saved", bool)
            entry["track_revisions"] = _get("TrackRevisions", bool)

            try:
                entry["pages"] = doc.ComputeStatistics(2)  # wdStatisticPages
            except Exception:
                entry["pages"] = None

            entry["active"] = (
                active_fullname is not None and entry.get("full_path") == active_fullname
            )
            if errors:
                entry["errors"] = errors
            documents.append(entry)

        return json.dumps(
            {
                "success": True,
                "count": len(documents),
                "documents": documents,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Undo", domain="lifecycle", change="edit")
async def word_live_undo(
    filename: str | None = None,
    times: int = 1,
    mcp_only: bool = True,
) -> str:
    """[Windows only] Undo the last N operations in an open Word document.

    Each MCP destructive tool call is grouped as a single undo entry (e.g.,
    "MCP: Insert Text"). Calling undo(times=1) reverts the last MCP operation;
    undo(times=3) reverts the last three.

    Args:
        filename: Document name or path (None = active document).
        times: Number of undo steps (default 1).
        mcp_only: Refuse to undo unless every requested entry is labeled "MCP:".

    Returns:
        JSON with success status and number of undone steps.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if times < 1:
        return json.dumps({"error": "times must be >= 1"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        checked_entries: list[str] = []
        if mcp_only:
            try:
                control = app.CommandBars.FindControl(Type=6, Id=128)
                if control is None or int(control.ListCount) < times:
                    return json.dumps(
                        {"error": "Word's undo history cannot confirm the requested MCP entries"}
                    )
                checked_entries = [str(control.List(index)) for index in range(1, times + 1)]
            except Exception:
                return json.dumps(
                    {"error": "Word's undo history is unavailable; no undo was performed"}
                )
            if any("mcp:" not in entry.casefold() for entry in checked_entries):
                return json.dumps(
                    {
                        "error": "The requested undo would include non-MCP user work",
                        "undo_entries": checked_entries,
                    },
                    ensure_ascii=False,
                )

        result = doc.Undo(times)

        return json.dumps(
            {
                "success": bool(result),
                "document": doc.Name,
                "times_requested": times,
                "undo_result": bool(result),
                "mcp_only": mcp_only,
                "undo_entries": checked_entries,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Open", domain="lifecycle", change="safe_write")
async def word_live_open(filename: str) -> str:
    """[Windows only] Open a document in Microsoft Word and make it active.

    Reuses a running Word instance when possible and starts a visible instance otherwise.

    Args:
        filename: Absolute path to an existing Word document.

    Returns:
        JSON with the opened document name and path.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Opening Word documents is only available on Windows"})

    path = os.path.abspath(filename)
    if not os.path.isfile(path):
        return json.dumps({"error": f"Document not found: {path}"})

    try:
        import win32com.client

        from word_mcp_codemode_live.core.word_com import get_word_app, remember_word_app

        try:
            app = get_word_app()
        except RuntimeError:
            app = win32com.client.Dispatch("Word.Application")

        remember_word_app(app)
        app.Visible = True
        normalized_path = os.path.normcase(os.path.normpath(path))
        for index in range(1, app.Documents.Count + 1):
            candidate = app.Documents(index)
            candidate_path = os.path.normcase(os.path.normpath(candidate.FullName))
            if candidate_path == normalized_path:
                candidate.Activate()
                return json.dumps(
                    {
                        "success": True,
                        "already_open": True,
                        "document": candidate.Name,
                        "path": candidate.FullName,
                    },
                    ensure_ascii=False,
                )

        document = app.Documents.Open(path)
        document.Activate()
        return json.dumps(
            {
                "success": True,
                "already_open": False,
                "document": document.Name,
                "path": document.FullName,
            },
            ensure_ascii=False,
        )
    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Close", domain="lifecycle", change="edit")
async def word_live_close(
    filename: str | None = None,
    save_mode: Literal["require_saved", "save", "discard"] = "require_saved",
) -> str:
    """Close an open Word document without allowing an implicit save prompt.

    ``require_saved`` is the safe default and refuses to close a document with
    unsaved changes. ``save`` saves in place before closing. ``discard`` closes
    without saving and must be requested explicitly.

    Args:
        filename: Document name or full path (None = active document).
        save_mode: Unsaved-change policy: require_saved, save, or discard.

    Returns:
        JSON with the closed path, save policy, and remaining open-document count.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)
        document_name = str(doc.Name)
        document_path = str(doc.FullName)
        saved_before = bool(doc.Saved)

        if save_mode == "require_saved" and not saved_before:
            return json.dumps(
                {
                    "error": "Document has unsaved changes; choose save_mode='save' "
                    "or save_mode='discard' explicitly",
                    "document": document_name,
                    "path": document_path,
                },
                ensure_ascii=False,
            )

        if save_mode == "save":
            doc.Save()

        # Always suppress Word's UI prompt. The requested policy was handled above.
        doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
        remaining_count = int(app.Documents.Count)
        active_document = None
        if remaining_count:
            active_document = str(app.ActiveDocument.FullName)

        return json.dumps(
            {
                "success": True,
                "document": document_name,
                "path": document_path,
                "save_mode": save_mode,
                "saved_before": saved_before,
                "remaining_open_documents": remaining_count,
                "active_document": active_document,
            },
            ensure_ascii=False,
        )
    except Exception as exc:
        return json.dumps({"error": str(exc)}, ensure_ascii=False)


@word_tool(title="Word Live Rename", domain="lifecycle", change="edit")
async def word_live_rename(
    new_path: str,
    filename: str | None = None,
) -> str:
    """Rename or move an open, saved Word document without leaving the old file.

    The destination must not already exist, its parent directory must exist, and
    its extension must match the source extension. Word saves the current document
    to the destination, the new path is verified, and only then is the old file
    removed. If old-file removal fails, the function attempts to restore the
    document to its original path and remove the new copy.

    Args:
        new_path: New local path for the document.
        filename: Document name or full path (None = active document).

    Returns:
        JSON with the original and new paths.
    """
    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)
        original_path = os.path.abspath(str(doc.FullName))
        destination = os.path.abspath(new_path)
        normalized_original = os.path.normcase(os.path.normpath(original_path))
        normalized_destination = os.path.normcase(os.path.normpath(destination))

        if normalized_destination == normalized_original:
            return json.dumps({"error": "New path is the current document path"})
        if not os.path.isfile(original_path):
            return json.dumps(
                {
                    "error": "The open document has no saved local source file; "
                    "use word_live_save(save_as=...) first"
                }
            )
        if not os.path.isdir(os.path.dirname(destination)):
            return json.dumps(
                {"error": f"Destination directory does not exist: {os.path.dirname(destination)}"}
            )
        if os.path.exists(destination):
            return json.dumps({"error": f"Destination already exists: {destination}"})

        original_extension = os.path.splitext(original_path)[1].casefold()
        destination_extension = os.path.splitext(destination)[1].casefold()
        if destination_extension != original_extension:
            return json.dumps(
                {
                    "error": "Rename must preserve the document extension; use "
                    "word_live_save(save_as=...) for format conversion"
                }
            )

        file_format = int(doc.SaveFormat)
        try:
            doc.SaveAs2(destination, FileFormat=file_format, AddToRecentFiles=False)
            actual_path = os.path.abspath(str(doc.FullName))
            if os.path.normcase(os.path.normpath(actual_path)) != normalized_destination:
                raise RuntimeError(
                    f"Word did not activate the requested renamed path: {actual_path}"
                )
            if not os.path.isfile(destination) or not bool(doc.Saved):
                raise RuntimeError("Word did not persist the renamed document")
            os.remove(original_path)
        except Exception as operation_error:
            rollback_error = None
            try:
                current_path = os.path.normcase(
                    os.path.normpath(os.path.abspath(str(doc.FullName)))
                )
                if current_path == normalized_destination:
                    previous_alerts = app.DisplayAlerts
                    try:
                        app.DisplayAlerts = 0
                        doc.SaveAs2(
                            original_path,
                            FileFormat=file_format,
                            AddToRecentFiles=False,
                        )
                    finally:
                        app.DisplayAlerts = previous_alerts
                if os.path.isfile(destination):
                    os.remove(destination)
            except Exception as exc:
                rollback_error = str(exc)
            if rollback_error is not None:
                return json.dumps(
                    {
                        "error": f"Rename failed: {operation_error}",
                        "rollback_error": rollback_error,
                        "document_path": str(doc.FullName),
                        "original_path": original_path,
                        "new_path": destination,
                    },
                    ensure_ascii=False,
                )
            return json.dumps(
                {
                    "error": f"Rename failed: {operation_error}; rename rolled back",
                    "document_path": str(doc.FullName),
                },
                ensure_ascii=False,
            )

        return json.dumps(
            {
                "success": True,
                "document": str(doc.Name),
                "original_path": original_path,
                "new_path": destination,
                "original_removed": not os.path.exists(original_path),
                "saved": bool(doc.Saved),
            },
            ensure_ascii=False,
        )
    except Exception as exc:
        return json.dumps({"error": str(exc)}, ensure_ascii=False)


@word_tool(title="Word Live Save", domain="lifecycle", change="edit")
async def word_live_save(
    filename: str | None = None,
    save_as: str | None = None,
    overwrite: bool = False,
) -> str:
    """Save an open Word document.

    Saves the document. Optionally saves to a new path with save_as.

    Args:
        filename: Document name or path (None = active document).
        save_as: Optional new file path to save as. If omitted, saves in place.
        overwrite: Allow save_as to replace an existing destination.

    Returns:
        JSON with save result.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app

        app = get_word_app()
        doc = find_document(app, filename)

        if save_as:
            save_path = os.path.abspath(save_as)
            # Determine format from extension
            ext = os.path.splitext(save_path)[1].lower()
            format_map = {
                ".docx": 16,  # wdFormatXMLDocument
                ".doc": 0,  # wdFormatDocument
                ".pdf": 17,  # wdFormatPDF
                ".rtf": 6,  # wdFormatRTF
                ".txt": 2,  # wdFormatText
            }
            if ext not in format_map:
                return json.dumps({"error": f"Unsupported save_as extension: {ext or '<none>'}"})
            if not os.path.isdir(os.path.dirname(save_path)):
                return json.dumps(
                    {"error": f"Destination directory does not exist: {os.path.dirname(save_path)}"}
                )
            same_path = os.path.normcase(os.path.normpath(save_path)) == os.path.normcase(
                os.path.normpath(str(doc.FullName))
            )
            if os.path.exists(save_path) and not same_path and not overwrite:
                return json.dumps({"error": f"Destination already exists: {save_path}"})
            file_format = format_map[ext]
            previous_alerts = app.DisplayAlerts
            try:
                if overwrite:
                    app.DisplayAlerts = 0
                doc.SaveAs2(save_path, FileFormat=file_format)
            finally:
                app.DisplayAlerts = previous_alerts
            return json.dumps(
                {
                    "success": True,
                    "document": doc.Name,
                    "saved_as": save_path,
                    "format": ext,
                },
                ensure_ascii=False,
            )
        else:
            doc.Save()
            return json.dumps(
                {
                    "success": True,
                    "document": doc.Name,
                    "path": doc.FullName,
                },
                ensure_ascii=False,
            )

    except Exception as e:
        return json.dumps({"error": str(e)})
