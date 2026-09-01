"""Insert, replace, and delete content in live Word documents."""

import json
import os
import re
import sys
from contextlib import contextmanager
from typing import Any

from word_mcp_codemode_live.defaults import DEFAULT_AUTHOR
from word_mcp_codemode_live.tools.metadata import word_tool

_INSERT_CHUNK_SIZE = 30000
_IMAGE_WRAP_STYLES = {
    "inline": None,
    "square": 0,
    "tight": 1,
    "behind": 3,
    "infront": 4,
    "topbottom": 2,
}
_IMAGE_BORDER_STYLES = {"none": 0, "single": 1, "double": 7, "dotted": 3, "dashed": 2, "thick": 6}
_IMAGE_ALIGNMENTS = {"left": 0, "center": 1, "right": 2}


def _normalized_word_text(text: str) -> str:
    return text.replace("\\r\\n", "\r").replace("\\r", "\r").replace("\\n", "\r")


def _text_chunks(text: str) -> list[str]:
    return [
        text[index : index + _INSERT_CHUNK_SIZE]
        for index in range(0, max(len(text), 1), _INSERT_CHUNK_SIZE)
    ]


def _insert_chunks(
    document: Any, app: Any, chunks: list[str], position: str, bookmark: str | None
) -> None:
    if bookmark is not None:
        word_range = document.Bookmarks(bookmark).Range
        for chunk in chunks:
            word_range.InsertAfter(chunk)
            word_range.Collapse(0)
        return
    if position == "start":
        for chunk in reversed(chunks):
            document.Range(0, 0).InsertBefore(chunk)
        return
    if position == "end":
        for chunk in chunks:
            end = document.Content.End - 1
            document.Range(end, end).InsertAfter(chunk)
        return
    if position == "cursor":
        for chunk in chunks:
            app.Selection.TypeText(chunk)
        return
    try:
        offset = int(position)
    except ValueError as exc:
        raise ValueError(
            f"Invalid position: {position}. Use 'start', 'end', 'cursor', or a character offset."
        ) from exc
    for chunk in reversed(chunks):
        document.Range(offset, offset).InsertBefore(chunk)


@contextmanager
def _replace_revision_state(app: Any, document: Any, track_changes: bool, replace_all: bool):
    previous_tracking = document.TrackRevisions
    previous_author = app.UserName
    if track_changes:
        document.TrackRevisions = True
        app.UserName = DEFAULT_AUTHOR
    elif replace_all and previous_tracking:
        document.TrackRevisions = False
    try:
        yield
    finally:
        document.TrackRevisions = previous_tracking
        if track_changes:
            app.UserName = previous_author


def _replace_matches(
    document: Any,
    find_text: str,
    replace_text: str,
    *,
    match_case: bool,
    match_whole_word: bool,
    use_wildcards: bool,
    replace_all: bool,
) -> int:
    processed = (
        replace_text.replace("^p", "\r")
        .replace("^t", "\t")
        .replace("^m", "\x0c")
        .replace("^s", "\u00a0")
    )
    word_range = document.Content.Duplicate
    word_range.Find.ClearFormatting()
    count = 0
    while word_range.Find.Execute(
        FindText=find_text,
        MatchCase=match_case,
        MatchWholeWord=match_whole_word if not use_wildcards else False,
        MatchWildcards=use_wildcards,
        Forward=True,
        Wrap=0,
    ):
        if word_range.Start == word_range.End:
            word_range.Start += 1
            word_range.End = document.Content.End
            continue
        word_range.Text = processed
        count += 1
        if not replace_all or count >= 50_000:
            break
        word_range.Collapse(0)
    return count


def _target_paragraph(
    document: Any, target_text: str | None, target_paragraph_index: int | None
) -> Any:
    if target_paragraph_index is not None:
        if not 1 <= target_paragraph_index <= document.Paragraphs.Count:
            raise ValueError(
                f"target_paragraph_index {target_paragraph_index} out of range "
                f"(1-{document.Paragraphs.Count})"
            )
        return document.Paragraphs(target_paragraph_index)
    for index in range(1, document.Paragraphs.Count + 1):
        paragraph = document.Paragraphs(index)
        if target_text in paragraph.Range.Text.rstrip("\r\x07"):
            return paragraph
    raise ValueError(f"No paragraph found containing {target_text!r}")


def _insert_paragraph_values(target: Any, paragraphs: list, position: str, style: Any) -> int:
    values = paragraphs if position == "after" else reversed(paragraphs)
    for text in values:
        word_range = target.Range.Duplicate
        word_range.Collapse(0 if position == "after" else 1)
        if position == "after":
            word_range.InsertParagraphAfter()
            word_range.Collapse(0)
        else:
            word_range.InsertParagraphBefore()
            word_range.Collapse(1)
        word_range.InsertAfter(text)
        word_range.Style = style
        if position == "after":
            word_range.Collapse(0)
    return len(paragraphs)


def _validate_image_options(options: dict[str, Any]) -> tuple[str, int]:
    path = os.path.abspath(options["image_path"])
    if not os.path.isfile(path):
        raise ValueError(f"Image file not found: {path}")
    choices = (
        ("wrapping", _IMAGE_WRAP_STYLES),
        ("border_style", _IMAGE_BORDER_STYLES),
        ("alignment", _IMAGE_ALIGNMENTS),
    )
    for name, allowed in choices:
        value = options[name]
        if value is not None and value.casefold() not in allowed:
            raise ValueError(f"Unknown {name}: {value}")
    dimensions = ("width_inches", "height_inches", "width_pt", "height_pt", "border_width_pt")
    if any(options[name] is not None and options[name] <= 0 for name in dimensions):
        raise ValueError("Image dimensions and border_width_pt must be positive")
    from word_mcp_codemode_live.core.word_values import rgb_hex_to_word

    color = (
        rgb_hex_to_word(options["border_color"], field_name="border_color")
        if options["border_color"]
        else 0
    )
    return path, color


def _image_range(document: Any, paragraph_index: int | None, position: str) -> Any:
    if paragraph_index is not None:
        if not 1 <= paragraph_index <= document.Paragraphs.Count:
            raise ValueError(
                f"paragraph_index {paragraph_index} out of range (1-{document.Paragraphs.Count})"
            )
        word_range = document.Paragraphs(paragraph_index).Range
        word_range.Collapse(1)
        return word_range
    if position == "start":
        return document.Range(0, 0)
    if position == "end":
        word_range = document.Range()
        word_range.Collapse(0)
        return word_range
    try:
        offset = int(position)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid position: {position}") from exc
    if not 0 <= offset < document.Content.End:
        raise ValueError(f"position must be between 0 and {document.Content.End - 1}")
    return document.Range(offset, offset)


def _image_dimensions(options: dict[str, Any]) -> tuple[float | None, float | None]:
    width = options["width_pt"]
    if width is None and options["width_inches"] is not None:
        width = float(options["width_inches"]) * 72.0
    height = options["height_pt"]
    if height is None and options["height_inches"] is not None:
        height = float(options["height_inches"]) * 72.0
    return width, height


def _resize_image(shape: Any, width: float | None, height: float | None) -> None:
    if width is not None and height is not None:
        shape.Width, shape.Height = width, height
    elif width is not None:
        ratio = shape.Height / shape.Width
        shape.Width, shape.Height = width, width * ratio
    elif height is not None:
        ratio = shape.Width / shape.Height
        shape.Height, shape.Width = height, height * ratio


def _apply_floating_image(shape: Any, document: Any, options: dict[str, Any], color: int) -> None:
    shape.WrapFormat.Type = _IMAGE_WRAP_STYLES[options["wrapping"].casefold()]
    _apply_shape_line(shape.Line, options, color)
    alignment = options["alignment"]
    if alignment is None:
        return
    shape.RelativeHorizontalPosition = 0
    shape.RelativeVerticalPosition = 2
    setup = document.PageSetup
    text_width = setup.PageWidth - setup.LeftMargin - setup.RightMargin
    positions = {
        "left": 0,
        "right": max(0, text_width - shape.Width),
        "center": max(0, (text_width - shape.Width) / 2),
    }
    shape.Left = positions[alignment.casefold()]


def _apply_shape_line(line: Any, options: dict[str, Any], color: int) -> None:
    style = options["border_style"]
    if style is None:
        return
    if style.casefold() == "none":
        line.Visible = False
        return
    line.Visible = True
    line.DashStyle = {"dotted": 3, "dashed": 4}.get(style.casefold(), 1)
    line.Weight = float(options["border_width_pt"] or 1.0)
    line.ForeColor.RGB = color
    if style.casefold() == "double":
        line.Style = 3


def _apply_inline_image(shape: Any, options: dict[str, Any], color: int) -> list[str]:
    failures: list[str] = []
    style = options["border_style"]
    if style is not None:
        for border_id in (-1, -2, -3, -4):
            try:
                border = shape.Borders(border_id)
                border.LineStyle = _IMAGE_BORDER_STYLES[style.casefold()]
                if style.casefold() != "none":
                    border.LineWidth = float(options["border_width_pt"] or 1.0)
                    border.Color = color
            except Exception as exc:
                failures.append(f"border {border_id}: {exc}")
    if options["alignment"] is not None:
        shape.Range.ParagraphFormat.Alignment = _IMAGE_ALIGNMENTS[options["alignment"].casefold()]
    return failures


@word_tool(
    title="Word Live Insert Page Break",
    domain="content",
    change="edit",
    batchable=True,
)
async def word_live_insert_page_break(
    filename: str | None = None,
    position: str = "end",
    paragraph_index: int | None = None,
    character_offset: int | None = None,
) -> dict[str, object]:
    """Insert a native Word page break at one unambiguous document position.

    ``paragraph_index`` is one-based and inserts before that paragraph.
    ``character_offset`` is a zero-based Word character position. When neither
    is supplied, ``position`` must be ``start`` or ``end``.
    """
    if sys.platform != "win32":
        raise RuntimeError("Live editing is only available on Windows")
    if paragraph_index is not None and character_offset is not None:
        raise ValueError("Provide paragraph_index or character_offset, not both")
    if position not in {"start", "end"}:
        raise ValueError("position must be 'start' or 'end'")

    from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

    app = get_word_app()
    document = find_document(app, filename)
    content_end = max(0, int(document.Content.End) - 1)

    if paragraph_index is not None:
        if paragraph_index < 1 or paragraph_index > int(document.Paragraphs.Count):
            raise ValueError(
                f"paragraph_index {paragraph_index} is outside 1-{int(document.Paragraphs.Count)}"
            )
        target = document.Paragraphs(paragraph_index).Range.Duplicate
        target.Collapse(1)  # wdCollapseStart
        resolved = {"kind": "paragraph", "paragraph_index": paragraph_index}
    elif character_offset is not None:
        if character_offset < 0 or character_offset > content_end:
            raise ValueError(f"character_offset must be between 0 and {content_end}")
        target = document.Range(character_offset, character_offset)
        resolved = {"kind": "character", "character_offset": character_offset}
    else:
        offset = 0 if position == "start" else content_end
        target = document.Range(offset, offset)
        resolved = {"kind": position, "character_offset": offset}

    pages_before = int(document.ComputeStatistics(2))
    insertion_offset = int(target.Start)
    with undo_record(app, "MCP: Insert Page Break"):
        target.InsertBreak(7)  # wdPageBreak
    document.Repaginate()

    return {
        "success": True,
        "document": str(document.Name),
        "insertion_offset": insertion_offset,
        "target": resolved,
        "pages_before": pages_before,
        "pages_after": int(document.ComputeStatistics(2)),
    }


@word_tool(title="Word Live Insert Text", domain="content", change="edit", batchable=True)
async def word_live_insert_text(
    filename: str | None = None,
    text: str = "",
    position: str = "end",
    bookmark: str | None = None,
    track_changes: bool = False,
) -> str:
    """Insert text into an open Word document.

    Automatically chunks large text (>30K chars) to avoid Word COM limits.

    Args:
        filename: Document name or path (None = active document).
        text: Text to insert (no length limit — auto-chunked if needed).
        position: "start", "end", "cursor", or character offset as string.
        bookmark: Insert after a named bookmark (overrides position).
        track_changes: Track the insertion as a revision.

    Returns:
        JSON with result info.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import (
            find_document,
            get_word_app,
            revision_tracking,
            undo_record,
        )

        app = get_word_app()
        doc = find_document(app, filename)

        if position == "cursor":
            selection_document = app.Selection.Range.Document
            if str(selection_document.FullName).casefold() != str(doc.FullName).casefold():
                return json.dumps(
                    {
                        "error": "The Word cursor is in a different document than filename. "
                        "Activate the requested document or use an explicit position."
                    }
                )

        text = _normalized_word_text(text)

        # Reject control bytes (notably \x07 cell separator) — inserting
        # these outside a real table creates invalid document state that
        # subsequent Find/Replace and table operations cannot recover from.
        from word_mcp_codemode_live.utils.text_safety import reject_control_chars

        reject_control_chars("text", text)
        if bookmark and not doc.Bookmarks.Exists(bookmark):
            raise ValueError(f"Bookmark {bookmark!r} not found")
        chunks = _text_chunks(text)

        with undo_record(app, "MCP: Insert Text"):
            with revision_tracking(app, doc, track_changes, DEFAULT_AUTHOR):
                _insert_chunks(doc, app, chunks, position, bookmark)

        result = {
            "success": True,
            "document": doc.Name,
            "text_length": len(text),
            "position": position,
            "tracked": track_changes,
        }
        if len(chunks) > 1:
            result["chunks_used"] = len(chunks)
        return json.dumps(result)

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Replace Text", domain="content", change="edit", batchable=True)
async def word_live_replace_text(
    filename: str | None = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = False,
    match_whole_word: bool = False,
    use_wildcards: bool = False,
    replace_all: bool = True,
    track_changes: bool = False,
) -> str:
    """[Windows only] Find and replace text in an open Word document.

    Uses Word's native Find & Replace, which works across tracked change boundaries
    (unlike manual delete+insert). Supports Word special characters when use_wildcards=True:
    ^m (manual page break), ^t (tab), ^p (paragraph mark), ^s (non-breaking space), and Word wildcard syntax.

    Args:
        filename: Document name or path (None = active document).
        find_text: Text to find. With use_wildcards=True, supports ^m, ^t, ^p, ^s and Word wildcards.
        replace_text: Replacement text. Use "" to delete matches.
        match_case: Case-sensitive search.
        match_whole_word: Match whole words only (ignored when use_wildcards=True).
        use_wildcards: Enable Word wildcards and special characters.
        replace_all: Replace all occurrences (True) or just the first one (False).
        track_changes: Track replacements as revisions.

    Returns:
        JSON with count of replacements made.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not find_text:
        return json.dumps({"error": "find_text is required"})

    if len(find_text) > 255:
        return json.dumps(
            {
                "error": f"find_text is {len(find_text)} chars (Word limit: 255). "
                "Break into smaller find/replace pairs."
            }
        )
    if len(replace_text) > 255:
        return json.dumps(
            {
                "error": f"replace_text is {len(replace_text)} chars (Word limit: 255). "
                "Break into smaller find/replace pairs."
            }
        )

    # Reject control bytes (notably \x07 cell separator) that can corrupt
    # Find/Replace and have historically caused full-document data loss.
    from word_mcp_codemode_live.utils.text_safety import reject_control_chars

    try:
        reject_control_chars("find_text", find_text)
        reject_control_chars("replace_text", replace_text)
    except ValueError as e:
        return json.dumps({"error": str(e)})

    if replace_all and track_changes:
        return json.dumps(
            {
                "error": "replace_all=True with track_changes=True causes an infinite loop "
                "(tracked deletions stay visible to Find, triggering endless re-replacement). "
                "Use replace_all=False — each unique text only needs one replacement."
            }
        )

    try:
        from word_mcp_codemode_live.core.word_com import (
            find_document,
            get_word_app,
            undo_record,
        )

        app = get_word_app()
        doc = find_document(app, filename)

        with undo_record(app, "MCP: Replace Text"):
            with _replace_revision_state(app, doc, track_changes, replace_all):
                count = _replace_matches(
                    doc,
                    find_text,
                    replace_text,
                    match_case=match_case,
                    match_whole_word=match_whole_word,
                    use_wildcards=use_wildcards,
                    replace_all=replace_all,
                )

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "find_text": find_text,
                "replace_text": replace_text,
                "replacements": count,
                "replace_all": replace_all,
                "tracked": track_changes,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(
    title="Word Live Insert Paragraphs",
    domain="content",
    change="edit",
    batchable=True,
)
async def word_live_insert_paragraphs(
    filename: str | None = None,
    paragraphs: list | None = None,
    target_text: str | None = None,
    target_paragraph_index: int | None = None,
    position: str = "after",
    style: str | None = None,
    track_changes: bool = False,
) -> str:
    """[Windows only] Insert one or more paragraphs near a target paragraph in an open Word document.

    Targets by text match or one-based paragraph index, matching word_live_get_text.
    Inserts all paragraphs in a single undo record.

    Args:
        filename: Document name or path (None = active document).
        paragraphs: List of paragraph texts to insert. Each string becomes one Word paragraph.
        target_text: Text to search for (first matching paragraph). Mutually exclusive with target_paragraph_index.
        target_paragraph_index: One-based paragraph index (as returned by word_live_get_text).
        position: 'before' or 'after' the target paragraph (default 'after').
        style: Style name for inserted paragraphs. None = "Normal" (avoids inheriting heading styles).
        track_changes: Track insertions as revisions.

    Returns:
        JSON with result info including count of paragraphs inserted.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not paragraphs or not isinstance(paragraphs, list):
        return json.dumps({"error": "paragraphs must be a non-empty list of strings"})

    if target_text is None and target_paragraph_index is None:
        return json.dumps({"error": "Provide either target_text or target_paragraph_index"})

    if target_text is not None and target_paragraph_index is not None:
        return json.dumps({"error": "Provide target_text or target_paragraph_index, not both"})

    if position not in ("before", "after"):
        return json.dumps({"error": f"position must be 'before' or 'after', got '{position}'"})

    try:
        from word_mcp_codemode_live.core.word_com import (
            find_document,
            get_word_app,
            revision_tracking,
            undo_record,
        )

        app = get_word_app()
        doc = find_document(app, filename)

        target_para = _target_paragraph(doc, target_text, target_paragraph_index)
        resolved_style = style if style else "Normal"
        try:
            word_style = doc.Styles(resolved_style)
        except Exception as exc:
            raise ValueError(f"Word style not found: {resolved_style}") from exc

        with undo_record(app, "MCP: Insert Paragraphs"):
            with revision_tracking(app, doc, track_changes, DEFAULT_AUTHOR):
                inserted = _insert_paragraph_values(target_para, paragraphs, position, word_style)

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "paragraphs_inserted": inserted,
                "position": position,
                "style": resolved_style,
                "tracked": track_changes,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Delete Text", domain="content", change="edit", batchable=True)
async def word_live_delete_text(
    filename: str | None = None,
    start: int | None = None,
    end: int | None = None,
    track_changes: bool = False,
) -> str:
    """Delete text from an open Word document.

    Args:
        filename: Document name or path.
        start: Start character position.
        end: End character position.
        track_changes: Track deletion as a revision.

    Returns:
        JSON with deleted text info.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if start is None or end is None:
        return json.dumps({"error": "Both 'start' and 'end' character positions are required"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        app = get_word_app()
        doc = find_document(app, filename)
        if start < 0 or end <= start or end > doc.Content.End:
            return json.dumps(
                {"error": f"Range {start}-{end} is outside document bounds 0-{doc.Content.End}"}
            )
        intersecting_tables = [
            index
            for index in range(1, doc.Tables.Count + 1)
            if doc.Tables(index).Range.Start < end and doc.Tables(index).Range.End > start
        ]
        if intersecting_tables:
            return json.dumps(
                {
                    "error": "word_live_delete_text does not delete table structure. "
                    f"The range intersects table(s) {intersecting_tables}; use word_live_modify_table."
                }
            )
        rng = doc.Range(start, end)
        deleted_text = rng.Text

        with undo_record(app, "MCP: Delete Text"):
            prev_tracking = doc.TrackRevisions
            prev_author = app.UserName
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR

            try:
                rng.Delete()
            finally:
                if track_changes:
                    doc.TrackRevisions = prev_tracking
                    app.UserName = prev_author

        preview = deleted_text
        if len(preview) > 100:
            preview = preview[:100] + "..."

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "deleted_text": preview,
                "range": f"{start}-{end}",
                "tracked": track_changes,
            }
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Insert Image", domain="content", change="edit", batchable=True)
async def word_live_insert_image(
    filename: str | None = None,
    image_path: str = "",
    paragraph_index: int | None = None,
    position: str = "end",
    width_inches: float | None = None,
    height_inches: float | None = None,
    width_pt: float | None = None,
    height_pt: float | None = None,
    alignment: str | None = None,
    wrapping: str | None = None,
    border_style: str | None = None,
    border_width_pt: float | None = None,
    border_color: str | None = None,
    link_to_file: bool = False,
) -> str:
    """Insert an image into an open Word document.

    The image can be placed at a specific paragraph, at the start or end,
    or at a character offset position.

    Args:
        filename: Document name or path (None = active document).
        image_path: Full path to the image file (PNG, JPG, BMP, etc.).
        paragraph_index: 1-indexed paragraph to insert before (image goes before the paragraph).
        position: "start", "end", or character offset as string. Only used if paragraph_index is None.
        width_inches: Optional width in inches (aspect ratio maintained if only one dimension given).
        height_inches: Optional height in inches.
        width_pt: Optional width in points (1 inch = 72 pt). Overrides width_inches if both given.
        height_pt: Optional height in points. Overrides height_inches if both given.
        alignment: Paragraph alignment for the image: "left", "center", "right". Default: unchanged.
        wrapping: Text wrapping style: "inline" (default), "square", "tight", "behind",
            "infront", "topbottom". Non-inline converts to a floating Shape.
        border_style: Border style around the image: "single", "double", "dotted", "dashed",
            "thick", "none". Default: no border.
        border_width_pt: Border line width in points (e.g. 1.0, 2.0). Default: 1.0.
        border_color: Border color as "#RRGGBB" hex string. Default: black (#000000).
        link_to_file: If True, links to the file instead of embedding it.

    Returns:
        JSON with image insertion result.
    """

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    if not image_path:
        return json.dumps({"error": "image_path is required"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        options = {
            "image_path": image_path,
            "width_inches": width_inches,
            "height_inches": height_inches,
            "width_pt": width_pt,
            "height_pt": height_pt,
            "alignment": alignment,
            "wrapping": wrapping,
            "border_style": border_style,
            "border_width_pt": border_width_pt,
            "border_color": border_color,
        }
        abs_path, parsed_border_color = _validate_image_options(options)
        app = get_word_app()
        doc = find_document(app, filename)
        rng = _image_range(doc, paragraph_index, position)
        final_w, final_h = _image_dimensions(options)

        with undo_record(app, "MCP: Insert Image"):
            inline_shape = rng.InlineShapes.AddPicture(
                FileName=abs_path,
                LinkToFile=link_to_file,
                SaveWithDocument=not link_to_file,
            )

            _resize_image(inline_shape, final_w, final_h)

            result_width = inline_shape.Width
            result_height = inline_shape.Height
            result_wrapping = "inline"
            warnings: list[str] = []

            if wrapping is not None and _IMAGE_WRAP_STYLES[wrapping.casefold()] is not None:
                float_shape = inline_shape.ConvertToShape()
                _apply_floating_image(float_shape, doc, options, parsed_border_color)
                result_wrapping = (wrapping or "inline").lower()
                result_width = float_shape.Width
                result_height = float_shape.Height
            else:
                warnings = _apply_inline_image(inline_shape, options, parsed_border_color)

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "image": os.path.basename(abs_path),
                "width_pt": result_width,
                "height_pt": result_height,
                "alignment": alignment or "unchanged",
                "wrapping": result_wrapping,
                "border": border_style or "none",
                "linked": link_to_file,
                "warnings": warnings,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})


@word_tool(title="Word Live Insert Equation", domain="content", change="edit", batchable=True)
async def word_live_insert_equation(
    filename: str | None = None,
    equation: str = "",
    paragraph_index: int | None = None,
    position: str = "end",
    display_mode: bool = False,
) -> str:
    """Insert a mathematical equation into a Word document using UnicodeMath syntax.

    LaTeX-like commands (e.g. \\int, \\sum, \\alpha) are automatically converted to
    Unicode math symbols before insertion, ensuring proper rendering.

    Args:
        filename: Document name (uses active document if None).
        equation: Equation text in UnicodeMath syntax. Examples:
            Simple: "x^2 + y^2 = z^2", "E = mc^2"
            Fractions: "(a+b)/(c+d)"
            Square root: "\\sqrt(x^2+y^2)"
            Greek letters: "\\alpha + \\beta = \\gamma"
            Integrals: "\\int_0^\\infty e^(-x^2) dx"
            Summation: "\\sum_(i=1)^n i^2"
            Matrix: "\\matrix(a&b@c&d)"
            Taylor series: "f(x) = \\sum_(n=0)^\\infty (f^((n))(a))/(n!) (x-a)^n"
        paragraph_index: Insert after this paragraph (1-based). None = use position.
        position: "start" or "end" of document. Ignored if paragraph_index given.
        display_mode: If True, equation is centered on its own line (display style).
            If False, equation is inline with surrounding text.

    Returns:
        JSON with success status and equation details.
    """
    # LaTeX-like command to Unicode math symbol mapping.
    # Word's COM OMaths.Add + BuildUp doesn't process autocorrect entries,
    # so we must pre-convert commands like \int, \sum to their Unicode equivalents.
    UNICODE_MATH = {
        # Greek lowercase
        r"\alpha": "\u03b1",
        r"\beta": "\u03b2",
        r"\gamma": "\u03b3",
        r"\delta": "\u03b4",
        r"\epsilon": "\u03b5",
        r"\varepsilon": "\u03b5",
        r"\zeta": "\u03b6",
        r"\eta": "\u03b7",
        r"\theta": "\u03b8",
        r"\vartheta": "\u03d1",
        r"\iota": "\u03b9",
        r"\kappa": "\u03ba",
        r"\lambda": "\u03bb",
        r"\mu": "\u03bc",
        r"\nu": "\u03bd",
        r"\xi": "\u03be",
        r"\pi": "\u03c0",
        r"\rho": "\u03c1",
        r"\sigma": "\u03c3",
        r"\varsigma": "\u03c2",
        r"\tau": "\u03c4",
        r"\upsilon": "\u03c5",
        r"\phi": "\u03c6",
        r"\varphi": "\u03d5",
        r"\chi": "\u03c7",
        r"\psi": "\u03c8",
        r"\omega": "\u03c9",
        # Greek uppercase
        r"\Gamma": "\u0393",
        r"\Delta": "\u0394",
        r"\Theta": "\u0398",
        r"\Lambda": "\u039b",
        r"\Xi": "\u039e",
        r"\Pi": "\u03a0",
        r"\Sigma": "\u03a3",
        r"\Upsilon": "\u03a5",
        r"\Phi": "\u03a6",
        r"\Psi": "\u03a8",
        r"\Omega": "\u03a9",
        # Operators / big operators
        r"\int": "\u222b",
        r"\iint": "\u222c",
        r"\iiint": "\u222d",
        r"\oint": "\u222e",
        r"\sum": "\u2211",
        r"\prod": "\u220f",
        r"\coprod": "\u2210",
        # Roots and radicals
        r"\sqrt": "\u221a",
        r"\cbrt": "\u221b",
        # Calculus / analysis
        r"\partial": "\u2202",
        r"\nabla": "\u2207",
        r"\infty": "\u221e",
        # Logic / set theory
        r"\forall": "\u2200",
        r"\exists": "\u2203",
        r"\nexists": "\u2204",
        r"\in": "\u2208",
        r"\notin": "\u2209",
        r"\subset": "\u2282",
        r"\supset": "\u2283",
        r"\subseteq": "\u2286",
        r"\supseteq": "\u2287",
        r"\cup": "\u222a",
        r"\cap": "\u2229",
        r"\emptyset": "\u2205",
        r"\neg": "\u00ac",
        r"\land": "\u2227",
        r"\lor": "\u2228",
        # Arithmetic / relations
        r"\pm": "\u00b1",
        r"\mp": "\u2213",
        r"\times": "\u00d7",
        r"\div": "\u00f7",
        r"\cdot": "\u22c5",
        r"\leq": "\u2264",
        r"\geq": "\u2265",
        r"\neq": "\u2260",
        r"\approx": "\u2248",
        r"\equiv": "\u2261",
        r"\cong": "\u2245",
        r"\sim": "\u223c",
        r"\propto": "\u221d",
        r"\ll": "\u226a",
        r"\gg": "\u226b",
        # Arrows
        r"\rightarrow": "\u2192",
        r"\leftarrow": "\u2190",
        r"\leftrightarrow": "\u2194",
        r"\Rightarrow": "\u21d2",
        r"\Leftarrow": "\u21d0",
        r"\Leftrightarrow": "\u21d4",
        r"\uparrow": "\u2191",
        r"\downarrow": "\u2193",
        r"\mapsto": "\u21a6",
        # Dots
        r"\cdots": "\u22ef",
        r"\ldots": "\u2026",
        r"\vdots": "\u22ee",
        r"\ddots": "\u22f1",
        # Miscellaneous
        r"\angle": "\u2220",
        r"\degree": "\u00b0",
        r"\star": "\u22c6",
        r"\circ": "\u2218",
        r"\bullet": "\u2022",
        r"\diamond": "\u22c4",
        r"\triangle": "\u25b3",
        r"\hbar": "\u210f",
        r"\ell": "\u2113",
        r"\Re": "\u211c",
        r"\Im": "\u2124",
        r"\aleph": "\u2135",
        # Matrix (Word UnicodeMath uses ■ for matrix)
        r"\matrix": "\u25a0",
        r"\pmatrix": "\u25a0",
        # Function names (these stay as text but without backslash)
        r"\lim": "lim",
        r"\sin": "sin",
        r"\cos": "cos",
        r"\tan": "tan",
        r"\sec": "sec",
        r"\csc": "csc",
        r"\cot": "cot",
        r"\arcsin": "arcsin",
        r"\arccos": "arccos",
        r"\arctan": "arctan",
        r"\sinh": "sinh",
        r"\cosh": "cosh",
        r"\tanh": "tanh",
        r"\log": "log",
        r"\ln": "ln",
        r"\exp": "exp",
        r"\det": "det",
        r"\dim": "dim",
        r"\ker": "ker",
        r"\min": "min",
        r"\max": "max",
        r"\inf": "inf",
        r"\sup": "sup",
        r"\gcd": "gcd",
        r"\arg": "arg",
        r"\mod": "mod",
    }

    if sys.platform != "win32":
        return json.dumps({"error": "Live editing is only available on Windows"})

    try:
        from word_mcp_codemode_live.core.word_com import find_document, get_word_app, undo_record

        app = get_word_app()
        doc = find_document(app, filename)

        if not equation or not equation.strip():
            return json.dumps({"error": "equation text is required"})

        with undo_record(app, "MCP: Insert Equation"):
            # Determine insertion range
            if paragraph_index is not None:
                if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                    return json.dumps(
                        {
                            "error": f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                        }
                    )
                rng = doc.Paragraphs(paragraph_index).Range
                rng.Collapse(0)  # After the paragraph
                rng.InsertParagraphAfter()
                rng.Collapse(0)
            elif position == "start":
                rng = doc.Paragraphs(1).Range
                rng.Collapse(1)  # Before first paragraph
                rng.InsertParagraphBefore()
                rng = doc.Paragraphs(1).Range
                rng.Collapse(1)
            else:  # "end"
                rng = doc.Content
                rng.Collapse(0)  # After last content
                rng.InsertParagraphAfter()
                rng.Collapse(0)

            # Convert LaTeX-like commands to Unicode math symbols.
            # Sort by length descending so longer matches take priority
            # (e.g. \iint before \int, \infty before \in).
            # Use negative lookahead (?![a-zA-Z]) to avoid partial matches.
            _commands = sorted(UNICODE_MATH.keys(), key=len, reverse=True)
            _pattern = "|".join(re.escape(c) for c in _commands)
            _pattern = f"({_pattern})(?![a-zA-Z])"
            eq_text = re.sub(_pattern, lambda m: UNICODE_MATH[m.group(1)], equation)

            # Insert the converted equation text
            rng.Text = eq_text

            # Convert to OMath
            doc.OMaths.Add(rng)
            omath = doc.OMaths(doc.OMaths.Count)

            # Set display mode (centered on own line) vs inline
            if display_mode:
                omath.Type = 1  # wdOMathDisplay
            else:
                omath.Type = 0  # wdOMathInline

            # Build up the equation (render UnicodeMath to formatted equation)
            omath.BuildUp()

        return json.dumps(
            {
                "success": True,
                "document": doc.Name,
                "equation": equation,
                "display_mode": display_mode,
                "omath_count": doc.OMaths.Count,
            },
            ensure_ascii=False,
        )

    except Exception as e:
        return json.dumps({"error": str(e)})
