"""Create, format, and modify tables in live Word documents."""

import logging
import re
from typing import Annotated, Any, Literal

from pydantic import Field

from word_mcp_codemode_live.defaults import DEFAULT_AUTHOR
from word_mcp_codemode_live.tools.metadata import word_tool
from word_mcp_codemode_live.word import session as word_session
from word_mcp_codemode_live.word import tables as table_com
from word_mcp_codemode_live.word.values import rgb_hex_to_word

logger = logging.getLogger(__name__)

_BORDER_STYLES = {"none": 0, "single": 1, "double": 7, "dotted": 3, "dashed": 2, "thick": 6}
_BORDER_IDS = (-1, -2, -3, -4, -5, -6, -7, -8)
_AUTOFIT_VALUES = {"window": 2, "content": 1, "fixed": 0}
_TABLE_ALIGNMENTS = {"left": 0, "center": 1, "right": 2}
_PARAGRAPH_ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "justify": 3}


def _validate_table_data(data: list | None, rows: int, cols: int) -> None:
    if data is None:
        return
    if not isinstance(data, list) or any(not isinstance(row, list) for row in data):
        raise ValueError("data must be a two-dimensional list")
    if len(data) > rows or any(len(row) > cols for row in data):
        raise ValueError("data dimensions exceed rows or cols")


def _insertion_range(document: Any, position: str) -> Any:
    if position == "start":
        return document.Range(0, 0)
    if position == "end":
        end = document.Content.End - 1
        return document.Range(end, end)
    try:
        offset = int(position)
    except ValueError as exc:
        raise ValueError(f"Invalid position: {position}") from exc
    if not 0 <= offset < document.Content.End:
        raise ValueError(f"position must be between 0 and {document.Content.End - 1}")
    _reject_existing_table_offset(document, offset)
    _reject_orphan_separator(document, offset)
    return document.Range(offset, offset)


def _reject_existing_table_offset(document: Any, offset: int) -> None:
    for table in document.Tables:
        try:
            start, end = table.Range.Start, table.Range.End
        except Exception as exc:
            logger.debug("Could not inspect an existing table range: %s", exc)
            continue
        if start <= offset <= end:
            raise ValueError(
                f"position offset {offset} falls within an existing table at range "
                f"[{start}, {end}]. Choose an offset outside any table, or use "
                "position='end'/'start'."
            )


def _reject_orphan_separator(document: Any, offset: int) -> None:
    if offset <= 0:
        return
    try:
        probe = document.Range(offset - 1, offset).Text or ""
    except Exception:
        probe = ""
    if probe == "\x07":
        raise ValueError(
            f"position offset {offset} sits immediately after an orphan cell separator (\\x07). "
            "Run word_live_modify_table operation='delete_table' with scrub_orphans=True "
            "(the default) on the prior table, then inspect the affected range with "
            "word_live_get_text."
        )


def _populate_table(table: Any, data: list | None) -> None:
    for row_index, row in enumerate(data or [], start=1):
        for column_index, value in enumerate(row, start=1):
            table.Cell(row_index, column_index).Range.Text = str(value)


def _resolve_table(document: Any, table_index: int) -> tuple[Any, int]:
    count = int(document.Tables.Count)
    if count == 0:
        raise ValueError("Document has no tables")
    index = table_index if table_index > 0 else count
    if not 1 <= index <= count:
        raise ValueError(f"Table index {table_index} out of range (1-{count})")
    return document.Tables(index), index


def _cell_coordinates(table: Any, row: int, column: int):
    rows = range(1, int(table.Rows.Count) + 1) if row == 0 else (row,)
    columns = range(1, int(table.Columns.Count) + 1) if column == 0 else (column,)
    return ((row_index, column_index) for row_index in rows for column_index in columns)


def _validate_cell_entry(table: Any, entry: list, *, allow_zero: bool, name: str) -> None:
    if len(entry) != 3:
        raise ValueError(f"Invalid {name} entry: {entry}")
    row, column = int(entry[0]), int(entry[1])
    minimum = 0 if allow_zero else 1
    if not minimum <= row <= table.Rows.Count or not minimum <= column <= table.Columns.Count:
        raise ValueError(f"Invalid {name} entry: {entry}")


def _validate_table_format(table: Any, options: dict[str, Any]) -> None:
    named_options = (
        ("border_style", options["border_style"], _BORDER_STYLES),
        ("autofit", options["autofit"], _AUTOFIT_VALUES),
        ("table_alignment", options["table_alignment"], _TABLE_ALIGNMENTS),
    )
    for name, value, choices in named_options:
        if value is not None and value.casefold() not in choices:
            raise ValueError(f"Unknown {name}: {value}")
    if options["left_indent_points"] is not None and options["left_indent_points"] < 0:
        raise ValueError("left_indent_points cannot be negative")
    if options["font_size"] is not None and options["font_size"] <= 0:
        raise ValueError("font_size must be greater than zero")
    widths = options["column_widths"]
    if widths is not None and (
        len(widths) > table.Columns.Count or any(float(width) <= 0 for width in widths)
    ):
        raise ValueError("column_widths must contain positive values for existing columns")
    _validate_table_cell_options(table, options)


def _validate_table_cell_options(table: Any, options: dict[str, Any]) -> None:
    for entry in options["cell_bold"] or []:
        _validate_cell_entry(table, entry, allow_zero=False, name="cell_bold")
    for entry in options["cell_alignment"] or []:
        _validate_cell_entry(table, entry, allow_zero=True, name="cell_alignment")
        if str(entry[2]).casefold() not in _PARAGRAPH_ALIGNMENTS:
            raise ValueError(f"Invalid cell alignment: {entry[2]}")
    for entry in options["cell_shading"] or []:
        _validate_cell_entry(table, entry, allow_zero=True, name="cell_shading")
        if not re.fullmatch(r"#?[0-9A-Fa-f]{6}", str(entry[2])):
            raise ValueError(f"Invalid shading color: {entry[2]}")
    cell_options = ("column_widths", "cell_bold", "cell_alignment", "cell_shading")
    if not bool(table.Uniform) and any(options[name] is not None for name in cell_options):
        raise ValueError("Cell/column formatting requires a uniform table")


def _apply_table_properties(table: Any, options: dict[str, Any], actions: list[str]) -> None:
    assignments = (
        ("table_alignment", table.Rows, "Alignment", _TABLE_ALIGNMENTS),
        ("left_indent_points", table.Rows, "LeftIndent", None),
        ("font_name", table.Range.Font, "Name", None),
        ("font_size", table.Range.Font, "Size", None),
    )
    for name, target, attribute, mapping in assignments:
        value = options[name]
        if value is None:
            continue
        converted = mapping[value.casefold()] if mapping else value
        setattr(target, attribute, converted)
        actions.append(f"{name}={value}")
    if options["autofit"] is not None:
        table.AutoFitBehavior(_AUTOFIT_VALUES[options["autofit"].casefold()])
        actions.append(f"autofit={options['autofit']}")


def _apply_table_borders(table: Any, style: str | None, actions: list[str]) -> list[str]:
    if style is None:
        return []
    failures = []
    for border_id in _BORDER_IDS:
        try:
            table.Borders(border_id).LineStyle = _BORDER_STYLES[style.casefold()]
        except Exception as exc:
            failures.append(f"border {border_id}: {exc}")
    actions.append(f"borders={style}")
    return failures


def _apply_table_cells(table: Any, options: dict[str, Any], actions: list[str]) -> None:
    if options["column_widths"] is not None:
        for index, width in enumerate(options["column_widths"], start=1):
            table.Columns(index).Width = float(width)
        actions.append(f"column_widths={options['column_widths']}")
    for row, column, value in options["cell_bold"] or []:
        table.Cell(int(row), int(column)).Range.Font.Bold = bool(value)
    if options["cell_bold"] is not None:
        actions.append(f"cell_bold={len(options['cell_bold'])} cells")
    _apply_cell_alignment(table, options["cell_alignment"], actions)
    _apply_cell_shading(table, options["cell_shading"], actions)


def _apply_cell_alignment(table: Any, entries: list | None, actions: list[str]) -> None:
    for row, column, alignment in entries or []:
        value = _PARAGRAPH_ALIGNMENTS[str(alignment).casefold()]
        for row_index, column_index in _cell_coordinates(table, int(row), int(column)):
            table.Cell(row_index, column_index).Range.ParagraphFormat.Alignment = value
    if entries is not None:
        actions.append(f"cell_alignment={len(entries)} entries")


def _apply_cell_shading(table: Any, entries: list | None, actions: list[str]) -> None:
    for row, column, color in entries or []:
        value = rgb_hex_to_word(str(color), field_name="cell shading color")
        for row_index, column_index in _cell_coordinates(table, int(row), int(column)):
            table.Cell(row_index, column_index).Shading.BackgroundPatternColor = value
    if entries is not None:
        actions.append(f"cell_shading={len(entries)} entries")


def _execute_write_operation(table_com: Any, table: Any, operation: str, values: dict[str, Any]):
    if operation == "set_cell":
        if values["row"] is None or values["col"] is None or values["text"] is None:
            raise ValueError("set_cell requires row, col, and text")
        return table_com.set_cell(
            table,
            values["row"],
            values["col"],
            values["text"],
            accept_revisions=values["accept_revisions"],
        )
    if operation == "set_row":
        if values["row"] is None or not values["cells"]:
            raise ValueError("set_row requires row and cells (list of values)")
        return table_com.set_row(
            table, values["row"], values["cells"], accept_revisions=values["accept_revisions"]
        )
    if operation == "set_range":
        if not values["cells"]:
            raise ValueError("set_range requires cells (2D list of values)")
        return table_com.set_range(
            table,
            values["cells"],
            start_row=values["start_row"] or 1,
            start_col=values["start_col"] or 1,
            accept_revisions=values["accept_revisions"],
        )
    raise ValueError(f"Unknown table write operation: {operation}")


def _execute_table_operation(table_com: Any, table: Any, operation: str, values: dict[str, Any]):
    if operation in {"set_cell", "set_row", "set_range"}:
        return _execute_write_operation(table_com, table, operation, values)
    if operation == "add_column":
        return table_com.add_column(table, values["before_col"], values["header"], values["cells"])
    if operation == "delete_column":
        if values["col"] is None:
            raise ValueError("delete_column requires col")
        return table_com.delete_column(table, values["col"])
    if operation == "add_row":
        return table_com.add_row(table, values["before_row"], values["cells"])
    if operation == "delete_row":
        if values["row"] is None:
            raise ValueError("delete_row requires row")
        return table_com.delete_row(table, values["row"])
    if operation == "merge_cells":
        coordinates = [values[name] for name in ("start_row", "start_col", "end_row", "end_col")]
        if any(value is None for value in coordinates):
            raise ValueError("merge_cells requires start_row, start_col, end_row, end_col")
        return table_com.merge_cells(table, *coordinates)
    if operation == "autofit":
        return table_com.autofit(table, values["autofit_mode"])
    if operation == "delete_table":
        return table_com.delete_table(table, scrub_orphans=values["scrub_orphans"])
    raise ValueError(
        f"Unknown operation {operation!r}. Use: get_info, set_cell, set_row, set_range, "
        "add_column, delete_column, add_row, delete_row, merge_cells, autofit, delete_table"
    )


@word_tool(title="Word Live Add Table", domain="tables", change="edit", batchable=True)
async def word_live_add_table(
    filename: str | None = None,
    rows: Annotated[int, Field(ge=1)] = 2,
    cols: Annotated[int, Field(ge=1)] = 2,
    position: str = "end",
    data: list | None = None,
    style: str = "Table Grid",
    autofit: Literal["window", "content", "fixed"] | None = "window",
    track_changes: bool = False,
) -> dict[str, Any]:
    """Add a table to an open Word document.

    Args:
        filename: Document name or path.
        rows: Number of rows.
        cols: Number of columns.
        position: "start", "end", or character offset.
        data: Optional 2D list of cell data.
        style: Table style name. Default "Table Grid" (bordered).
            Use None or "" for no style.
        autofit: "window" (fit page width, default), "content" (fit cell content),
            "fixed" (fixed widths), or None for legacy behavior (no autofit).
        track_changes: Track as revision.

    Returns:
        JSON with result info.
    """

    word_session.require_windows("Live Word editing")

    if rows < 1 or cols < 1:
        raise ValueError("rows and cols must be at least 1")
    autofit_map = {
        "window": (1, 2),
        "content": (1, 1),
        "fixed": (0, 0),
    }
    if autofit is not None and autofit.casefold() not in autofit_map:
        raise ValueError("autofit must be window, content, fixed, or null")

    _validate_table_data(data, rows, cols)
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    if style:
        try:
            resolved_style = doc.Styles(style)
        except Exception as exc:
            raise ValueError(f"Word table style not found: {style}") from exc
    else:
        resolved_style = None

    rng = _insertion_range(doc, position)

    with word_session.undo_record(app, "MCP: Add Table"):
        with word_session.revision_tracking(app, doc, track_changes, DEFAULT_AUTHOR):
            if autofit:
                default_behavior, autofit_behavior = autofit_map[autofit.casefold()]
                table = doc.Tables.Add(rng, rows, cols, default_behavior, autofit_behavior)
            else:
                table = doc.Tables.Add(rng, rows, cols)

                # Apply table style
            if resolved_style is not None:
                table.Style = resolved_style
            _populate_table(table, data)

    return {
        "success": True,
        "document": str(doc.Name),
        "rows": rows,
        "cols": cols,
        "position": position,
        "style": style or None,
        "autofit": autofit or None,
        "tracked": track_changes,
    }


@word_tool(title="Word Live Format Table", domain="tables", change="edit", batchable=True)
async def word_live_format_table(
    filename: str | None = None,
    table_index: int = -1,
    border_style: str | None = None,
    cell_bold: list | None = None,
    cell_alignment: list | None = None,
    column_widths: list | None = None,
    table_alignment: str | None = None,
    cell_shading: list | None = None,
    autofit: str | None = None,
    left_indent_points: float | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
) -> dict[str, Any]:
    """Format a table in an open Word document via COM.

    Supports border removal, cell formatting, column sizing, and table alignment.
    Use table_index=-1 for the last table, 1 for the first, etc.

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index, or -1 for the last table.
        border_style: Border style for all edges: "none", "single", "double", "dotted",
            "dashed", "thick". "none" removes all borders.
        cell_bold: List of [row, col, bold] entries (1-indexed) to set bold on cell text.
            Example: [[1, 1, true], [1, 2, true]] bolds row 1 cells.
        cell_alignment: List of [row, col, alignment] entries. alignment: "left", "center",
            "right", "justify". Row 0 = all rows, Col 0 = all cols.
        column_widths: List of column widths in points (1-indexed order).
            Example: [200, 200] sets col 1 to 200pt, col 2 to 200pt.
        table_alignment: Table alignment on page: "left", "center", "right".
        cell_shading: List of [row, col, color_hex] entries. color_hex as "#RRGGBB".
            Row 0 = all rows. Example: [[1, 0, "#DDDDDD"]] shades entire row 1.
        autofit: "window" (fit to page width), "content" (fit to cell content),
            "fixed" (fixed column widths).
        left_indent_points: Distance from the left margin to the table edge, in points.
        font_name: Font family for all table text.
        font_size: Font size in points for all table text.

    Returns:
        JSON with result info.
    """

    word_session.require_windows("Live Word editing")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)
    tbl, idx = _resolve_table(doc, table_index)
    options = {
        "border_style": border_style,
        "cell_bold": cell_bold,
        "cell_alignment": cell_alignment,
        "column_widths": column_widths,
        "table_alignment": table_alignment,
        "cell_shading": cell_shading,
        "autofit": autofit,
        "left_indent_points": left_indent_points,
        "font_name": font_name,
        "font_size": font_size,
    }
    _validate_table_format(tbl, options)
    actions: list[str] = []

    with word_session.undo_record(app, "MCP: Format Table"):
        warnings = _apply_table_borders(tbl, border_style, actions)
        _apply_table_properties(tbl, options, actions)
        _apply_table_cells(tbl, options, actions)

    return {
        "success": True,
        "document": str(doc.Name),
        "table_index": idx,
        "rows": int(tbl.Rows.Count),
        "cols": int(tbl.Columns.Count),
        "actions": actions,
        "warnings": warnings,
    }


@word_tool(title="Word Live Modify Table", domain="tables", change="edit", batchable=True)
async def word_live_modify_table(
    filename: str | None = None,
    table_index: Annotated[int, Field(ge=1)] = 1,
    operation: Literal[
        "get_info",
        "set_cell",
        "set_row",
        "set_range",
        "add_column",
        "delete_column",
        "add_row",
        "delete_row",
        "merge_cells",
        "autofit",
        "delete_table",
    ] = "get_info",
    row: int | None = None,
    col: int | None = None,
    text: str | None = None,
    before_row: int | None = None,
    before_col: int | None = None,
    header: str | None = None,
    cells: list | None = None,
    start_row: int | None = None,
    start_col: int | None = None,
    end_row: int | None = None,
    end_col: int | None = None,
    autofit_mode: str = "content",
    accept_revisions: bool = False,
    track_changes: bool = False,
    scrub_orphans: bool = True,
) -> dict[str, Any]:
    """[Windows only] Modify a table in an open Word document.

    Operations: get_info, set_cell, set_row, set_range, add_column, delete_column,
    add_row, delete_row, merge_cells, autofit, delete_table.
    All row/col indices are 1-based (Word COM standard).

    Args:
        filename: Document name or path (None = active document).
        table_index: 1-based table index (default 1).
        operation: One of: get_info, set_cell, set_row, set_range, add_column,
            delete_column, add_row, delete_row, merge_cells, autofit, delete_table.
        row: Row index for set_cell, set_row, delete_row.
        col: Column index for set_cell, delete_column.
        text: Text for set_cell.
        before_row: Insert row before this index (add_row). None = append at end.
        before_col: Insert column before this index (add_column). None = append at end.
        header: Header text for new column (add_column, placed in row 1).
        cells: List of cell values for set_row (1D) or set_range (2D). None values skip that cell.
            Also used for new row/column values (add_row, add_column).
        start_row: Start row for merge_cells or set_range (default 1).
        start_col: Start column for merge_cells or set_range (default 1).
        end_row: End row for merge_cells.
        end_col: End column for merge_cells.
        autofit_mode: 'content', 'window', or 'fixed' (autofit operation).
        accept_revisions: For set_cell/set_row/set_range — accept tracked changes before writing
            (prevents layered text from old revisions persisting underneath new content).
        track_changes: Track modifications as revisions.
        scrub_orphans: For delete_table — scan the deletion site for orphan
            cell-separator (\\x07) bytes and remove them. Default True.

    Returns:
        JSON with operation result.
    """

    word_session.require_windows("Live Word editing")
    app = word_session.get_word_app()
    doc = word_session.find_document(app, filename)

    # Per-call validation: re-read Tables.Count fresh in case a prior
    # MCP call (especially delete_table) reduced or zeroed the count.
    try:
        table_count = int(doc.Tables.Count)
    except Exception as exc:
        raise RuntimeError(f"Could not enumerate document tables: {exc}") from exc

    if table_count == 0:
        raise ValueError("Document has no tables")

    if not (1 <= table_index <= table_count):
        raise ValueError(
            f"table_index {table_index} out of range. Document has "
            f"{table_count} table(s) (valid range: 1..{table_count}). "
            "If a prior delete_table reduced the count, call word_live_get_info to refresh."
        )

    table = doc.Tables(table_index)
    op = operation.lower()

    # get_info is read-only — no undo record needed
    if op == "get_info":
        result = table_com.get_info(table)
        result["document"] = str(doc.Name)
        result["table_index"] = table_index
        return result

    values = {
        "row": row,
        "col": col,
        "text": text,
        "before_row": before_row,
        "before_col": before_col,
        "header": header,
        "cells": cells,
        "start_row": start_row,
        "start_col": start_col,
        "end_row": end_row,
        "end_col": end_col,
        "autofit_mode": autofit_mode,
        "accept_revisions": accept_revisions,
        "scrub_orphans": scrub_orphans,
    }
    with word_session.undo_record(app, "MCP: Modify Table"):
        with word_session.revision_tracking(app, doc, track_changes, DEFAULT_AUTHOR):
            result = _execute_table_operation(table_com, table, op, values)

    result["success"] = True
    result["document"] = str(doc.Name)
    result["table_index"] = table_index
    result["operation"] = op
    result["tracked"] = track_changes
    return result
