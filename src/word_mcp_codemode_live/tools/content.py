"""Insert, replace, and delete content in live Word documents."""

import os
from contextlib import contextmanager
from typing import Annotated, Any, Literal

from pydantic import BaseModel, Field

from word_mcp_codemode_live.defaults import DEFAULT_AUTHOR
from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.validation import reject_control_chars
from word_mcp_codemode_live.word import images as word_images
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word.equations import to_unicode_math
from word_mcp_codemode_live.word.ranges import character_range

_INSERT_CHUNK_SIZE = 30000


class DeleteTextResult(BaseModel):
    """Structured result for a text deletion."""

    success: Literal[True] = True
    document: str
    deleted_text: str
    range: str
    tracked: bool


class InsertTextResult(BaseModel):
    success: Literal[True] = True
    document: str
    text_length: int
    position: str
    tracked: bool
    chunks_used: int | None = None


class ReplaceTextResult(BaseModel):
    success: Literal[True] = True
    document: str
    find_text: str
    replace_text: str
    replacements: int
    replace_all: bool
    tracked: bool


class InsertParagraphsResult(BaseModel):
    success: Literal[True] = True
    document: str
    paragraphs_inserted: int
    position: Literal["before", "after"]
    style: str
    tracked: bool


class InsertImageResult(BaseModel):
    success: Literal[True] = True
    document: str
    image: str
    width_pt: float
    height_pt: float
    alignment: str
    wrapping: str
    border: str
    linked: bool
    warnings: list[str]


class InsertEquationResult(BaseModel):
    success: Literal[True] = True
    document: str
    equation: str
    display_mode: bool
    omath_count: int


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
    word_session.require_windows("Live Word editing")
    if paragraph_index is not None and character_offset is not None:
        raise ValueError("Provide paragraph_index or character_offset, not both")
    if position not in {"start", "end"}:
        raise ValueError("position must be 'start' or 'end'")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
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
    with word_session.undo_record(app, "MCP: Insert Page Break"):
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
) -> InsertTextResult:
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

    word_session.require_windows("Live Word editing")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    if position == "cursor":
        selection_document = app.Selection.Range.Document
        if str(selection_document.FullName).casefold() != str(doc.FullName).casefold():
            raise RuntimeError(
                "The Word cursor is in a different document than filename. "
                "Activate the requested document or use an explicit position."
            )

    text = _normalized_word_text(text)

    # Reject control bytes (notably \x07 cell separator) — inserting
    # these outside a real table creates invalid document state that
    # subsequent Find/Replace and table operations cannot recover from.
    reject_control_chars("text", text)
    if bookmark and not doc.Bookmarks.Exists(bookmark):
        raise ValueError(f"Bookmark {bookmark!r} not found")
    chunks = _text_chunks(text)

    with word_session.undo_record(app, "MCP: Insert Text"):
        with word_session.revision_tracking(app, doc, track_changes, DEFAULT_AUTHOR):
            _insert_chunks(doc, app, chunks, position, bookmark)

    return InsertTextResult(
        document=str(doc.Name),
        text_length=len(text),
        position=position,
        tracked=track_changes,
        chunks_used=len(chunks) if len(chunks) > 1 else None,
    )


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
) -> ReplaceTextResult:
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

    word_session.require_windows("Live Word editing")

    if not find_text:
        raise ValueError("find_text is required")

    if len(find_text) > 255:
        raise ValueError(
            f"find_text is {len(find_text)} chars (Word limit: 255). "
            "Break into smaller find/replace pairs."
        )
    if len(replace_text) > 255:
        raise ValueError(
            f"replace_text is {len(replace_text)} chars (Word limit: 255). "
            "Break into smaller find/replace pairs."
        )

    # Reject control bytes (notably \x07 cell separator) that can corrupt
    # Find/Replace and have historically caused full-document data loss.
    reject_control_chars("find_text", find_text)
    reject_control_chars("replace_text", replace_text)

    if replace_all and track_changes:
        raise ValueError(
            "replace_all=True with track_changes=True causes an infinite loop "
            "(tracked deletions stay visible to Find, triggering endless re-replacement). "
            "Use replace_all=False — each unique text only needs one replacement."
        )

    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    with word_session.undo_record(app, "MCP: Replace Text"):
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

    return ReplaceTextResult(
        document=str(doc.Name),
        find_text=find_text,
        replace_text=replace_text,
        replacements=count,
        replace_all=replace_all,
        tracked=track_changes,
    )


@word_tool(
    title="Word Live Insert Paragraphs",
    domain="content",
    change="edit",
    batchable=True,
)
async def word_live_insert_paragraphs(
    filename: str | None = None,
    paragraphs: list[str] | None = None,
    target_text: str | None = None,
    target_paragraph_index: Annotated[int, Field(ge=1)] | None = None,
    position: Literal["before", "after"] = "after",
    style: str | None = None,
    track_changes: bool = False,
) -> InsertParagraphsResult:
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

    word_session.require_windows("Live Word editing")

    if not paragraphs or not isinstance(paragraphs, list):
        raise ValueError("paragraphs must be a non-empty list of strings")

    if target_text is None and target_paragraph_index is None:
        raise ValueError("Provide either target_text or target_paragraph_index")

    if target_text is not None and target_paragraph_index is not None:
        raise ValueError("Provide target_text or target_paragraph_index, not both")

    if position not in ("before", "after"):
        raise ValueError(f"position must be 'before' or 'after', got '{position}'")

    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    target_para = _target_paragraph(doc, target_text, target_paragraph_index)
    resolved_style = style if style else "Normal"
    try:
        word_style = doc.Styles(resolved_style)
    except Exception as exc:
        raise ValueError(f"Word style not found: {resolved_style}") from exc

    with word_session.undo_record(app, "MCP: Insert Paragraphs"):
        with word_session.revision_tracking(app, doc, track_changes, DEFAULT_AUTHOR):
            inserted = _insert_paragraph_values(target_para, paragraphs, position, word_style)

    return InsertParagraphsResult(
        document=str(doc.Name),
        paragraphs_inserted=inserted,
        position=position,
        style=resolved_style,
        tracked=track_changes,
    )


@word_tool(title="Word Live Delete Text", domain="content", change="edit", batchable=True)
async def word_live_delete_text(
    filename: str | None = None,
    start: Annotated[int, Field(ge=0)] | None = None,
    end: Annotated[int, Field(ge=0)] | None = None,
    track_changes: bool = False,
) -> DeleteTextResult:
    """Delete text from an open Word document.

    Args:
        filename: Document name or path.
        start: Start character position.
        end: End character position.
        track_changes: Track deletion as a revision.

    Returns:
        Structured information about the deleted text.
    """
    word_session.require_windows("Live Word editing")

    if start is None or end is None:
        raise ValueError("Both start and end character positions are required")

    app = word_session.get_word_app()
    document = word_session.find_document(app, filename)
    resolved = character_range(document, start, end)
    intersecting_tables = [
        index
        for index in range(1, int(document.Tables.Count) + 1)
        if int(document.Tables(index).Range.Start) < end
        and int(document.Tables(index).Range.End) > start
    ]
    if intersecting_tables:
        raise ValueError(
            "word_live_delete_text does not delete table structure. "
            f"The range intersects table(s) {intersecting_tables}; use word_live_modify_table."
        )

    deleted_text = str(resolved.com_range.Text)
    with word_session.undo_record(app, "MCP: Delete Text"):
        with word_session.revision_tracking(app, document, track_changes, DEFAULT_AUTHOR):
            resolved.com_range.Delete()

    preview = deleted_text[:100] + ("..." if len(deleted_text) > 100 else "")
    return DeleteTextResult(
        document=str(document.Name),
        deleted_text=preview,
        range=resolved.label,
        tracked=track_changes,
    )


@word_tool(title="Word Live Insert Image", domain="content", change="edit", batchable=True)
async def word_live_insert_image(
    filename: str | None = None,
    image_path: str = "",
    paragraph_index: Annotated[int, Field(ge=1)] | None = None,
    position: str = "end",
    width_inches: Annotated[float, Field(gt=0)] | None = None,
    height_inches: Annotated[float, Field(gt=0)] | None = None,
    width_pt: Annotated[float, Field(gt=0)] | None = None,
    height_pt: Annotated[float, Field(gt=0)] | None = None,
    alignment: Literal["left", "center", "right"] | None = None,
    wrapping: Literal["inline", "square", "tight", "behind", "infront", "topbottom"] | None = None,
    border_style: Literal["none", "single", "double", "dotted", "dashed", "thick"] | None = None,
    border_width_pt: Annotated[float, Field(gt=0)] | None = None,
    border_color: str | None = None,
    link_to_file: bool = False,
) -> InsertImageResult:
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

    word_session.require_windows("Live Word editing")

    if not image_path:
        raise ValueError("image_path is required")

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
    abs_path, parsed_border_color = word_images.validate_options(options)
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)
    rng = word_images.insertion_range(doc, paragraph_index, position)
    final_w, final_h = word_images.dimensions(options)

    with word_session.undo_record(app, "MCP: Insert Image"):
        inline_shape = rng.InlineShapes.AddPicture(
            FileName=abs_path,
            LinkToFile=link_to_file,
            SaveWithDocument=not link_to_file,
        )

        word_images.resize(inline_shape, final_w, final_h)

        result_width = float(inline_shape.Width)
        result_height = float(inline_shape.Height)
        result_wrapping = "inline"
        warnings: list[str] = []

        if wrapping is not None and word_images.WRAP_STYLES[wrapping.casefold()] is not None:
            float_shape = inline_shape.ConvertToShape()
            word_images.apply_floating(float_shape, doc, options, parsed_border_color)
            result_wrapping = wrapping.lower()
            result_width = float(float_shape.Width)
            result_height = float(float_shape.Height)
        else:
            warnings = word_images.apply_inline(inline_shape, options, parsed_border_color)

    return InsertImageResult(
        document=str(doc.Name),
        image=os.path.basename(abs_path),
        width_pt=result_width,
        height_pt=result_height,
        alignment=alignment or "unchanged",
        wrapping=result_wrapping,
        border=border_style or "none",
        linked=link_to_file,
        warnings=warnings,
    )


@word_tool(title="Word Live Insert Equation", domain="content", change="edit", batchable=True)
async def word_live_insert_equation(
    filename: str | None = None,
    equation: str = "",
    paragraph_index: Annotated[int, Field(ge=1)] | None = None,
    position: Literal["start", "end"] = "end",
    display_mode: bool = False,
) -> InsertEquationResult:
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

    word_session.require_windows("Live Word editing")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    if not equation or not equation.strip():
        raise ValueError("equation text is required")

    with word_session.undo_record(app, "MCP: Insert Equation"):
        # Determine insertion range
        if paragraph_index is not None:
            if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
                raise ValueError(
                    f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})"
                )
            rng = doc.Paragraphs(paragraph_index).Range
            rng.Collapse(0)
            rng.InsertParagraphAfter()
            rng.Collapse(0)
        elif position == "start":
            rng = doc.Paragraphs(1).Range
            rng.Collapse(1)
            rng.InsertParagraphBefore()
            rng = doc.Paragraphs(1).Range
            rng.Collapse(1)
        else:
            rng = doc.Content
            rng.Collapse(0)
            rng.InsertParagraphAfter()
            rng.Collapse(0)

            # Convert LaTeX-like commands to Unicode math symbols.
            # Sort by length descending so longer matches take priority
            # (e.g. \iint before \int, \infty before \in).
            # Use negative lookahead (?![a-zA-Z]) to avoid partial matches.
        equation_text = to_unicode_math(equation)

        # Insert the converted equation text
        rng.Text = equation_text

        # Convert to OMath
        doc.OMaths.Add(rng)
        omath = doc.OMaths(doc.OMaths.Count)

        # Set display mode (centered on own line) vs inline
        omath.Type = 1 if display_mode else 0

        # Build up the equation (render UnicodeMath to formatted equation)
        omath.BuildUp()

    return InsertEquationResult(
        document=str(doc.Name),
        equation=equation,
        display_mode=display_mode,
        omath_count=int(doc.OMaths.Count),
    )
