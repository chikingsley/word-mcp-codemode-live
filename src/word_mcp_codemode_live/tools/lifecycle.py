"""Open, save, inspect, and undo live Word document state."""

import os
from typing import Annotated, Literal

from pydantic import BaseModel, Field

from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session

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


class UndoHistoryResult(BaseModel):
    success: Literal[True] = True
    document: str
    undo_entries: list[str]
    count: int
    note: str | None = None


class PropertyChange(BaseModel):
    old: str | None
    new: str


class SetCorePropertiesResult(BaseModel):
    ok: bool
    document: str
    changed: dict[str, PropertyChange]
    errors: dict[str, str] | None = None


class OpenDocumentEntry(BaseModel):
    index: int
    name: str | None = None
    full_path: str | None = None
    saved: bool | None = None
    track_revisions: bool | None = None
    pages: int | None = None
    active: bool = False
    errors: list[str] | None = None


class ListOpenResult(BaseModel):
    success: Literal[True] = True
    count: int
    documents: list[OpenDocumentEntry]


class UndoResult(BaseModel):
    success: bool
    document: str
    times_requested: int
    undo_result: bool
    mcp_only: bool
    undo_entries: list[str]


class OpenResult(BaseModel):
    success: Literal[True] = True
    already_open: bool
    document: str
    path: str


class CloseResult(BaseModel):
    success: Literal[True] = True
    document: str
    path: str
    save_mode: Literal["require_saved", "save", "discard"]
    saved_before: bool
    remaining_open_documents: int
    active_document: str | None


class RenameResult(BaseModel):
    success: Literal[True] = True
    document: str
    original_path: str
    new_path: str
    original_removed: bool
    saved: bool


class SaveResult(BaseModel):
    success: Literal[True] = True
    document: str
    path: str | None = None
    saved_as: str | None = None
    format: str | None = None


@word_tool(title="Word Live Get Undo History", domain="lifecycle", change="read")
async def word_live_get_undo_history(
    filename: str | None = None,
) -> UndoHistoryResult:
    """[Windows only] Get the undo stack names from an open Word document.

    Uses Word's CommandBars to read the undo dropdown list. Each MCP tool call
    that was wrapped with undo_record will appear as "MCP: <tool name>".
    Degrades gracefully if the undo list is not accessible.

    Args:
        filename: Document name or path (None = active document).

    Returns:
        JSON with undo_entries list (most recent first) and count.
    """

    word_session.require_windows("Live Word tools")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    entries: list[str] = []
    try:
        # CommandBar control ID 128 = Undo split dropdown (Type=6).
        undo_control = app.CommandBars.FindControl(Type=6, Id=128)
        if undo_control is not None:
            entries = [str(undo_control.List(i)) for i in range(1, undo_control.ListCount + 1)]
    except Exception:
        # This undocumented API is unavailable in some Word versions.
        return UndoHistoryResult(
            document=str(doc.Name),
            undo_entries=[],
            count=0,
            note="Undo history not accessible in this Word version",
        )

    return UndoHistoryResult(document=str(doc.Name), undo_entries=entries, count=len(entries))


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
) -> SetCorePropertiesResult:
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
    word_session.require_windows("Live Word tools")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)
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
    changes: dict[str, PropertyChange] = {}
    errors: dict[str, str] = {}

    with word_session.undo_record(app, "MCP: Set Core Properties"):
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
                changes[key] = PropertyChange(old=old, new=value)
            except Exception as exc:
                errors[key] = str(exc)

    return SetCorePropertiesResult(
        ok=not errors,
        document=str(doc.Name),
        changed=changes,
        errors=errors or None,
    )


@word_tool(title="Word Live List Open", domain="lifecycle", change="read")
async def word_live_list_open() -> ListOpenResult:
    """[Windows only] List all documents currently open in Microsoft Word.

    Returns JSON with list of open documents including name, full_path,
    pages, saved status, and whether it is the active document.
    """

    word_session.require_windows("Live Word tools")
    app = word_session.get_word_app()

    # A broken ActiveDocument proxy should not block listing healthy documents.
    try:
        active_fullname = app.ActiveDocument.FullName if app.Documents.Count > 0 else None
    except Exception:
        active_fullname = None

    try:
        count = int(app.Documents.Count)
    except Exception as exc:
        raise RuntimeError(f"Could not enumerate Word's Documents collection: {exc}") from exc

    documents: list[OpenDocumentEntry] = []
    for index in range(1, count + 1):
        try:
            doc = app.Documents(index)
        except Exception as exc:
            documents.append(
                OpenDocumentEntry(
                    index=index,
                    name="<unavailable>",
                    errors=[f"could not access Documents({index}): {exc}"],
                )
            )
            continue

        errors: list[str] = []

        def _get(attr: str, transform=None, *, _doc=doc, _errors=errors):
            try:
                value = getattr(_doc, attr)
                return transform(value) if transform else value
            except Exception as exc:
                _errors.append(f"{attr}: {exc}")
                return None

        try:
            pages = int(doc.ComputeStatistics(2))  # wdStatisticPages
        except Exception as exc:
            pages = None
            errors.append(f"pages: {exc}")
        full_path = _get("FullName", str)
        documents.append(
            OpenDocumentEntry(
                index=index,
                name=_get("Name", str),
                full_path=full_path,
                saved=_get("Saved", bool),
                track_revisions=_get("TrackRevisions", bool),
                pages=pages,
                active=active_fullname is not None and full_path == active_fullname,
                errors=errors or None,
            )
        )

    return ListOpenResult(count=len(documents), documents=documents)


@word_tool(title="Word Live Undo", domain="lifecycle", change="edit")
async def word_live_undo(
    filename: str | None = None,
    times: Annotated[int, Field(ge=1)] = 1,
    mcp_only: bool = True,
) -> UndoResult:
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

    word_session.require_windows("Live Word editing")
    if times < 1:
        raise ValueError("times must be >= 1")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    checked_entries: list[str] = []
    if mcp_only:
        try:
            control = app.CommandBars.FindControl(Type=6, Id=128)
            if control is None or int(control.ListCount) < times:
                raise RuntimeError("Word's undo history cannot confirm the requested MCP entries")
            checked_entries = [str(control.List(index)) for index in range(1, times + 1)]
        except RuntimeError:
            raise
        except Exception as exc:
            raise RuntimeError("Word's undo history is unavailable; no undo was performed") from exc
        if any("mcp:" not in entry.casefold() for entry in checked_entries):
            raise RuntimeError(
                "The requested undo would include non-MCP user work: " + ", ".join(checked_entries)
            )

    result = bool(doc.Undo(times))
    return UndoResult(
        success=result,
        document=str(doc.Name),
        times_requested=times,
        undo_result=result,
        mcp_only=mcp_only,
        undo_entries=checked_entries,
    )


@word_tool(title="Word Live Open", domain="lifecycle", change="safe_write")
async def word_live_open(filename: str) -> OpenResult:
    """[Windows only] Open a document in Microsoft Word and make it active.

    Reuses a running Word instance when possible and starts a visible instance otherwise.

    Args:
        filename: Absolute path to an existing Word document.

    Returns:
        JSON with the opened document name and path.
    """
    word_session.require_windows("Opening Word documents")

    path = os.path.abspath(filename)
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Document not found: {path}")

    import win32com.client

    try:
        app = word_session.get_word_app()
    except RuntimeError:
        app = win32com.client.Dispatch("Word.Application")

    word_session.remember_word_app(app)
    app.Visible = True
    normalized_path = os.path.normcase(os.path.normpath(path))
    for index in range(1, app.Documents.Count + 1):
        candidate = app.Documents(index)
        candidate_path = os.path.normcase(os.path.normpath(candidate.FullName))
        if candidate_path == normalized_path:
            candidate.Activate()
            return OpenResult(
                already_open=True,
                document=str(candidate.Name),
                path=str(candidate.FullName),
            )

    document = app.Documents.Open(path)
    document.Activate()
    return OpenResult(already_open=False, document=str(document.Name), path=str(document.FullName))


@word_tool(title="Word Live Close", domain="lifecycle", change="edit")
async def word_live_close(
    filename: str | None = None,
    save_mode: Literal["require_saved", "save", "discard"] = "require_saved",
) -> CloseResult:
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
    word_session.require_windows("Live Word editing")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)
    document_name = str(doc.Name)
    document_path = str(doc.FullName)
    saved_before = bool(doc.Saved)

    if save_mode == "require_saved" and not saved_before:
        raise RuntimeError(
            "Document has unsaved changes; choose save_mode='save' "
            "or save_mode='discard' explicitly"
        )
    if save_mode == "save":
        doc.Save()

    # Always suppress Word's UI prompt. The requested policy was handled above.
    doc.Close(SaveChanges=0)  # wdDoNotSaveChanges
    remaining_count = int(app.Documents.Count)
    active_document = str(app.ActiveDocument.FullName) if remaining_count else None
    return CloseResult(
        document=document_name,
        path=document_path,
        save_mode=save_mode,
        saved_before=saved_before,
        remaining_open_documents=remaining_count,
        active_document=active_document,
    )


@word_tool(title="Word Live Rename", domain="lifecycle", change="edit")
async def word_live_rename(
    new_path: str,
    filename: str | None = None,
) -> RenameResult:
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
    word_session.require_windows("Live Word editing")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)
    original_path = os.path.abspath(str(doc.FullName))
    destination = os.path.abspath(new_path)
    normalized_original = os.path.normcase(os.path.normpath(original_path))
    normalized_destination = os.path.normcase(os.path.normpath(destination))

    if normalized_destination == normalized_original:
        raise ValueError("New path is the current document path")
    if not os.path.isfile(original_path):
        raise FileNotFoundError(
            "The open document has no saved local source file; "
            "use word_live_save(save_as=...) first"
        )
    destination_directory = os.path.dirname(destination)
    if not os.path.isdir(destination_directory):
        raise FileNotFoundError(f"Destination directory does not exist: {destination_directory}")
    if os.path.exists(destination):
        raise FileExistsError(f"Destination already exists: {destination}")

    original_extension = os.path.splitext(original_path)[1].casefold()
    destination_extension = os.path.splitext(destination)[1].casefold()
    if destination_extension != original_extension:
        raise ValueError(
            "Rename must preserve the document extension; use "
            "word_live_save(save_as=...) for format conversion"
        )

    file_format = int(doc.SaveFormat)
    try:
        doc.SaveAs2(destination, FileFormat=file_format, AddToRecentFiles=False)
        actual_path = os.path.abspath(str(doc.FullName))
        if os.path.normcase(os.path.normpath(actual_path)) != normalized_destination:
            raise RuntimeError(f"Word did not activate the requested renamed path: {actual_path}")
        if not os.path.isfile(destination) or not bool(doc.Saved):
            raise RuntimeError("Word did not persist the renamed document")
        os.remove(original_path)
    except Exception as operation_error:
        rollback_error = None
        try:
            current_path = os.path.normcase(os.path.normpath(os.path.abspath(str(doc.FullName))))
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
            raise RuntimeError(
                f"Rename failed: {operation_error}; rollback failed: {rollback_error}; "
                f"document path: {doc.FullName}; original path: {original_path}; "
                f"new path: {destination}"
            ) from operation_error
        raise RuntimeError(
            f"Rename failed: {operation_error}; rename rolled back; document path: {doc.FullName}"
        ) from operation_error

    return RenameResult(
        document=str(doc.Name),
        original_path=original_path,
        new_path=destination,
        original_removed=not os.path.exists(original_path),
        saved=bool(doc.Saved),
    )


@word_tool(title="Word Live Save", domain="lifecycle", change="edit")
async def word_live_save(
    filename: str | None = None,
    save_as: str | None = None,
    overwrite: bool = False,
) -> SaveResult:
    """Save an open Word document.

    Saves the document. Optionally saves to a new path with save_as.

    Args:
        filename: Document name or path (None = active document).
        save_as: Optional new file path to save as. If omitted, saves in place.
        overwrite: Allow save_as to replace an existing destination.

    Returns:
        JSON with save result.
    """

    word_session.require_windows("Live Word editing")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    if save_as:
        save_path = os.path.abspath(save_as)
        ext = os.path.splitext(save_path)[1].lower()
        format_map = {
            ".docx": 16,  # wdFormatXMLDocument
            ".doc": 0,  # wdFormatDocument
            ".pdf": 17,  # wdFormatPDF
            ".rtf": 6,  # wdFormatRTF
            ".txt": 2,  # wdFormatText
        }
        if ext not in format_map:
            raise ValueError(f"Unsupported save_as extension: {ext or '<none>'}")
        destination_directory = os.path.dirname(save_path)
        if not os.path.isdir(destination_directory):
            raise FileNotFoundError(
                f"Destination directory does not exist: {destination_directory}"
            )
        same_path = os.path.normcase(os.path.normpath(save_path)) == os.path.normcase(
            os.path.normpath(str(doc.FullName))
        )
        if os.path.exists(save_path) and not same_path and not overwrite:
            raise FileExistsError(f"Destination already exists: {save_path}")
        previous_alerts = app.DisplayAlerts
        try:
            if overwrite:
                app.DisplayAlerts = 0
            doc.SaveAs2(save_path, FileFormat=format_map[ext])
        finally:
            app.DisplayAlerts = previous_alerts
        return SaveResult(document=str(doc.Name), saved_as=save_path, format=ext)

    doc.Save()
    return SaveResult(document=str(doc.Name), path=str(doc.FullName))
