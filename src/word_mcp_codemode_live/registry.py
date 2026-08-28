"""Registration of the public MCP tool surface."""

from collections.abc import Callable
from typing import Any

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

from word_mcp_codemode_live.tools import (
    comment_tools,
    comment_write_tools,
    content_tools,
    document_tools,
    extended_document_tools,
    footnote_tools,
    format_tools,
    hyperlink_tools,
    layout_tools,
    live_batch_tools,
    live_layout_tools,
    live_read_tools,
    live_tools,
    page_capture_tools,
    tracked_changes_tools,
)

ToolFunction = Callable[..., Any]
ToolSpec = tuple[ToolFunction, str, ToolAnnotations]

TOOL_SPECS: tuple[ToolSpec, ...] = (
    (
        document_tools.create_document,
        "create_document",
        ToolAnnotations(title="Create Word Document", destructiveHint=True),
    ),
    (
        document_tools.copy_document,
        "copy_document",
        ToolAnnotations(title="Copy Word Document", destructiveHint=True),
    ),
    (
        document_tools.get_document_info,
        "get_document_info",
        ToolAnnotations(title="Get Document Info", readOnlyHint=True),
    ),
    (
        document_tools.get_document_text,
        "get_document_text",
        ToolAnnotations(title="Get Document Text", readOnlyHint=True),
    ),
    (
        document_tools.get_document_outline,
        "get_document_outline",
        ToolAnnotations(title="Get Document Outline", readOnlyHint=True),
    ),
    (
        document_tools.list_available_documents,
        "list_available_documents",
        ToolAnnotations(title="List Available Documents", readOnlyHint=True),
    ),
    (
        document_tools.get_document_xml_tool,
        "get_document_xml",
        ToolAnnotations(title="Get Document XML", readOnlyHint=True),
    ),
    (
        content_tools.insert_header_near_text_tool,
        "insert_header_near_text",
        ToolAnnotations(title="Insert Header Near Text", readOnlyHint=False, destructiveHint=False),
    ),
    (
        content_tools.insert_line_or_paragraph_near_text_tool,
        "insert_line_or_paragraph_near_text",
        ToolAnnotations(title="Insert Line Near Text", readOnlyHint=False, destructiveHint=False),
    ),
    (
        content_tools.insert_numbered_list_near_text_tool,
        "insert_numbered_list_near_text",
        ToolAnnotations(title="Insert List Near Text", readOnlyHint=False, destructiveHint=False),
    ),
    (
        content_tools.add_paragraph,
        "add_paragraph",
        ToolAnnotations(title="Add Paragraph", readOnlyHint=False, destructiveHint=False),
    ),
    (
        content_tools.add_heading,
        "add_heading",
        ToolAnnotations(title="Add Heading", readOnlyHint=False, destructiveHint=False),
    ),
    (
        content_tools.add_picture,
        "add_picture",
        ToolAnnotations(title="Add Picture", readOnlyHint=False, destructiveHint=False),
    ),
    (
        content_tools.add_table,
        "add_table",
        ToolAnnotations(title="Add Table", readOnlyHint=False, destructiveHint=False),
    ),
    (
        content_tools.add_page_break,
        "add_page_break",
        ToolAnnotations(title="Add Page Break", readOnlyHint=False, destructiveHint=False),
    ),
    (
        content_tools.delete_paragraph,
        "delete_paragraph",
        ToolAnnotations(title="Delete Paragraph", destructiveHint=True),
    ),
    (
        content_tools.search_and_replace,
        "search_and_replace",
        ToolAnnotations(title="Search and Replace", destructiveHint=True),
    ),
    (
        format_tools.create_custom_style,
        "create_custom_style",
        ToolAnnotations(title="Create Custom Style", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.format_text,
        "format_text",
        ToolAnnotations(title="Format Text", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.format_table,
        "format_table",
        ToolAnnotations(title="Format Table", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.set_table_cell_shading,
        "set_table_cell_shading",
        ToolAnnotations(title="Set Table Cell Shading", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.apply_table_alternating_rows,
        "apply_table_alternating_rows",
        ToolAnnotations(
            title="Apply Alternating Row Colors", readOnlyHint=False, destructiveHint=False
        ),
    ),
    (
        format_tools.highlight_table_header,
        "highlight_table_header",
        ToolAnnotations(title="Highlight Table Header", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.merge_table_cells,
        "merge_table_cells",
        ToolAnnotations(title="Merge Table Cells", readOnlyHint=False, destructiveHint=True),
    ),
    (
        format_tools.merge_table_cells_horizontal,
        "merge_table_cells_horizontal",
        ToolAnnotations(title="Merge Cells Horizontally", readOnlyHint=False, destructiveHint=True),
    ),
    (
        format_tools.merge_table_cells_vertical,
        "merge_table_cells_vertical",
        ToolAnnotations(title="Merge Cells Vertically", readOnlyHint=False, destructiveHint=True),
    ),
    (
        format_tools.set_table_cell_alignment,
        "set_table_cell_alignment",
        ToolAnnotations(title="Set Cell Alignment", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.set_table_alignment_all,
        "set_table_alignment_all",
        ToolAnnotations(title="Set Table Alignment", readOnlyHint=False, destructiveHint=False),
    ),
    (
        footnote_tools.add_footnote_to_document,
        "add_footnote_to_document",
        ToolAnnotations(title="Add Footnote", readOnlyHint=False, destructiveHint=False),
    ),
    (
        footnote_tools.add_footnote_after_text,
        "add_footnote_after_text",
        ToolAnnotations(title="Add Footnote After Text", readOnlyHint=False, destructiveHint=False),
    ),
    (
        footnote_tools.add_footnote_before_text,
        "add_footnote_before_text",
        ToolAnnotations(
            title="Add Footnote Before Text", readOnlyHint=False, destructiveHint=False
        ),
    ),
    (
        footnote_tools.add_footnote_enhanced,
        "add_footnote_enhanced",
        ToolAnnotations(title="Add Footnote Enhanced", readOnlyHint=False, destructiveHint=False),
    ),
    (
        footnote_tools.add_endnote_to_document,
        "add_endnote_to_document",
        ToolAnnotations(title="Add Endnote", readOnlyHint=False, destructiveHint=False),
    ),
    (
        footnote_tools.customize_footnote_style,
        "customize_footnote_style",
        ToolAnnotations(
            title="Customize Footnote Style", readOnlyHint=False, destructiveHint=False
        ),
    ),
    (
        footnote_tools.delete_footnote_from_document,
        "delete_footnote_from_document",
        ToolAnnotations(title="Delete Footnote", destructiveHint=True),
    ),
    (
        footnote_tools.add_footnote_robust_tool,
        "add_footnote_robust",
        ToolAnnotations(title="Add Footnote Robust", readOnlyHint=False, destructiveHint=False),
    ),
    (
        footnote_tools.validate_footnotes_tool,
        "validate_document_footnotes",
        ToolAnnotations(title="Validate Footnotes", readOnlyHint=True),
    ),
    (
        footnote_tools.delete_footnote_robust_tool,
        "delete_footnote_robust",
        ToolAnnotations(title="Delete Footnote Robust", destructiveHint=True),
    ),
    (
        extended_document_tools.get_paragraph_text_from_document,
        "get_paragraph_text_from_document",
        ToolAnnotations(title="Get Paragraph Text", readOnlyHint=True),
    ),
    (
        extended_document_tools.find_text_in_document,
        "find_text_in_document",
        ToolAnnotations(title="Find Text", readOnlyHint=True),
    ),
    (
        extended_document_tools.get_highlighted_text_from_document,
        "get_highlighted_text",
        ToolAnnotations(title="Get Highlighted Text", readOnlyHint=True),
    ),
    (
        extended_document_tools.convert_to_pdf,
        "convert_to_pdf",
        ToolAnnotations(title="Convert to PDF", destructiveHint=True),
    ),
    (
        content_tools.replace_paragraph_block_below_header_tool,
        "replace_paragraph_block_below_header",
        ToolAnnotations(
            title="Replace Block Below Header", readOnlyHint=False, destructiveHint=True
        ),
    ),
    (
        content_tools.replace_block_between_manual_anchors_tool,
        "replace_block_between_manual_anchors",
        ToolAnnotations(
            title="Replace Block Between Anchors", readOnlyHint=False, destructiveHint=True
        ),
    ),
    (
        comment_tools.get_all_comments,
        "get_all_comments",
        ToolAnnotations(title="Get All Comments", readOnlyHint=True),
    ),
    (
        comment_tools.get_comments_by_author,
        "get_comments_by_author",
        ToolAnnotations(title="Get Comments by Author", readOnlyHint=True),
    ),
    (
        comment_tools.get_comments_for_paragraph,
        "get_comments_for_paragraph",
        ToolAnnotations(title="Get Comments for Paragraph", readOnlyHint=True),
    ),
    (
        comment_write_tools.add_comment,
        "add_comment",
        ToolAnnotations(title="Add Comment", readOnlyHint=False, destructiveHint=False),
    ),
    (
        hyperlink_tools.manage_hyperlinks,
        "manage_hyperlinks",
        ToolAnnotations(title="Manage Hyperlinks", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.set_table_column_width,
        "set_table_column_width",
        ToolAnnotations(title="Set Column Width", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.set_table_column_widths,
        "set_table_column_widths",
        ToolAnnotations(title="Set Column Widths", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.set_table_width,
        "set_table_width",
        ToolAnnotations(title="Set Table Width", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.auto_fit_table_columns,
        "auto_fit_table_columns",
        ToolAnnotations(title="Auto-Fit Table Columns", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.format_table_cell_text,
        "format_table_cell_text",
        ToolAnnotations(title="Format Cell Text", readOnlyHint=False, destructiveHint=False),
    ),
    (
        format_tools.set_table_cell_padding,
        "set_table_cell_padding",
        ToolAnnotations(title="Set Cell Padding", readOnlyHint=False, destructiveHint=False),
    ),
    (
        tracked_changes_tools.track_replace,
        "track_replace",
        ToolAnnotations(title="Track Replace", destructiveHint=True),
    ),
    (
        tracked_changes_tools.track_insert,
        "track_insert",
        ToolAnnotations(title="Track Insert", destructiveHint=True),
    ),
    (
        tracked_changes_tools.track_delete,
        "track_delete",
        ToolAnnotations(title="Track Delete", destructiveHint=True),
    ),
    (
        tracked_changes_tools.list_tracked_changes,
        "list_tracked_changes",
        ToolAnnotations(title="List Tracked Changes", readOnlyHint=True),
    ),
    (
        tracked_changes_tools.accept_tracked_changes,
        "accept_tracked_changes",
        ToolAnnotations(title="Accept Tracked Changes", destructiveHint=True),
    ),
    (
        tracked_changes_tools.reject_tracked_changes,
        "reject_tracked_changes",
        ToolAnnotations(title="Reject Tracked Changes", destructiveHint=True),
    ),
    (
        page_capture_tools.word_live_capture_pages,
        "word_live_capture_pages",
        ToolAnnotations(title="Render Live Word Pages", readOnlyHint=True),
    ),
    (
        live_batch_tools.word_live_edit_batch,
        "word_live_edit_batch",
        ToolAnnotations(title="Batch Edit and Verify Live Word", destructiveHint=True),
    ),
    (
        live_tools.word_live_insert_text,
        "word_live_insert_text",
        ToolAnnotations(title="Word Live Insert Text", destructiveHint=True),
    ),
    (
        live_tools.word_live_format_text,
        "word_live_format_text",
        ToolAnnotations(title="Word Live Format Text", destructiveHint=True),
    ),
    (
        live_tools.word_live_replace_text,
        "word_live_replace_text",
        ToolAnnotations(title="Word Live Replace Text", destructiveHint=True),
    ),
    (
        live_tools.word_live_insert_paragraphs,
        "word_live_insert_paragraphs",
        ToolAnnotations(title="Word Live Insert Paragraphs", destructiveHint=True),
    ),
    (
        live_tools.word_live_add_table,
        "word_live_add_table",
        ToolAnnotations(title="Word Live Add Table", destructiveHint=True),
    ),
    (
        live_tools.word_live_format_table,
        "word_live_format_table",
        ToolAnnotations(title="Word Live Format Table", destructiveHint=True),
    ),
    (
        live_tools.word_live_modify_table,
        "word_live_modify_table",
        ToolAnnotations(title="Word Live Modify Table", destructiveHint=True),
    ),
    (
        live_tools.word_live_delete_text,
        "word_live_delete_text",
        ToolAnnotations(title="Word Live Delete Text", destructiveHint=True),
    ),
    (
        live_tools.word_live_apply_list,
        "word_live_apply_list",
        ToolAnnotations(title="Word Live Apply List", destructiveHint=True),
    ),
    (
        live_tools.word_live_setup_heading_numbering,
        "word_live_setup_heading_numbering",
        ToolAnnotations(title="Word Live Setup Heading Numbering", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_get_text,
        "word_live_get_text",
        ToolAnnotations(title="Word Live Get Text", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_take_snapshot,
        "word_live_take_snapshot",
        ToolAnnotations(title="Word Live Take Snapshot", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_get_diff,
        "word_live_get_diff",
        ToolAnnotations(title="Word Live Get Diff", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_snapshot_status,
        "word_live_snapshot_status",
        ToolAnnotations(title="Word Live Snapshot Status", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_get_paragraph_format,
        "word_live_get_paragraph_format",
        ToolAnnotations(title="Word Live Get Paragraph Format", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_get_info,
        "word_live_get_info",
        ToolAnnotations(title="Word Live Get Info", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_set_core_properties,
        "word_live_set_core_properties",
        ToolAnnotations(title="Word Live Set Core Properties", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_list_open,
        "word_live_list_open",
        ToolAnnotations(title="Word Live List Open", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_find_text,
        "word_live_find_text",
        ToolAnnotations(title="Word Live Find Text", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_get_comments,
        "word_live_get_comments",
        ToolAnnotations(title="Word Live Get Comments", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_add_comment,
        "word_live_add_comment",
        ToolAnnotations(title="Word Live Add Comment", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_reply_to_comment,
        "word_live_reply_to_comment",
        ToolAnnotations(title="Word Live Reply to Comment", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_resolve_comment,
        "word_live_resolve_comment",
        ToolAnnotations(title="Word Live Resolve Comment", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_delete_comment,
        "word_live_delete_comment",
        ToolAnnotations(title="Word Live Delete Comment", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_list_revisions,
        "word_live_list_revisions",
        ToolAnnotations(title="Word Live List Revisions", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_accept_revisions,
        "word_live_accept_revisions",
        ToolAnnotations(title="Word Live Accept Revisions", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_reject_revisions,
        "word_live_reject_revisions",
        ToolAnnotations(title="Word Live Reject Revisions", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_get_page_text,
        "word_live_get_page_text",
        ToolAnnotations(title="Word Live Get Page Text", readOnlyHint=True),
    ),
    (
        live_read_tools.word_live_get_undo_history,
        "word_live_get_undo_history",
        ToolAnnotations(title="Word Live Get Undo History", readOnlyHint=True),
    ),
    (
        live_tools.word_live_undo,
        "word_live_undo",
        ToolAnnotations(title="Word Live Undo", destructiveHint=True),
    ),
    (
        live_tools.word_live_open,
        "word_live_open",
        ToolAnnotations(title="Word Live Open", destructiveHint=False),
    ),
    (
        live_tools.word_live_go_to_page,
        "word_live_go_to_page",
        ToolAnnotations(title="Word Live Go To Page", readOnlyHint=True),
    ),
    (
        live_tools.word_live_save,
        "word_live_save",
        ToolAnnotations(title="Word Live Save", destructiveHint=True),
    ),
    (
        live_tools.word_live_toggle_track_changes,
        "word_live_toggle_track_changes",
        ToolAnnotations(title="Word Live Toggle Track Changes", destructiveHint=True),
    ),
    (
        live_tools.word_live_insert_image,
        "word_live_insert_image",
        ToolAnnotations(title="Word Live Insert Image", destructiveHint=True),
    ),
    (
        live_tools.word_live_insert_cross_reference,
        "word_live_insert_cross_reference",
        ToolAnnotations(title="Word Live Insert Cross Reference", destructiveHint=True),
    ),
    (
        live_tools.word_live_list_cross_reference_items,
        "word_live_list_cross_reference_items",
        ToolAnnotations(title="Word Live List Cross Reference Items", readOnlyHint=True),
    ),
    (
        live_tools.word_live_insert_equation,
        "word_live_insert_equation",
        ToolAnnotations(title="Word Live Insert Equation", destructiveHint=True),
    ),
    (
        live_read_tools.word_live_diagnose_layout,
        "word_live_diagnose_layout",
        ToolAnnotations(title="Word Live Diagnose Layout", readOnlyHint=True),
    ),
    (
        live_layout_tools.word_live_set_page_layout,
        "word_live_set_page_layout",
        ToolAnnotations(title="Word Live Set Page Layout", destructiveHint=True),
    ),
    (
        live_layout_tools.word_live_add_header_footer,
        "word_live_add_header_footer",
        ToolAnnotations(title="Word Live Add Header/Footer", destructiveHint=True),
    ),
    (
        live_layout_tools.word_live_add_page_numbers,
        "word_live_add_page_numbers",
        ToolAnnotations(title="Word Live Add Page Numbers", destructiveHint=True),
    ),
    (
        live_layout_tools.word_live_add_section_break,
        "word_live_add_section_break",
        ToolAnnotations(title="Word Live Add Section Break", destructiveHint=True),
    ),
    (
        live_layout_tools.word_live_set_paragraph_spacing,
        "word_live_set_paragraph_spacing",
        ToolAnnotations(title="Word Live Set Paragraph Spacing", destructiveHint=True),
    ),
    (
        live_layout_tools.word_live_add_bookmark,
        "word_live_add_bookmark",
        ToolAnnotations(title="Word Live Add Bookmark", destructiveHint=True),
    ),
    (
        live_layout_tools.word_live_add_watermark,
        "word_live_add_watermark",
        ToolAnnotations(title="Word Live Add Watermark", destructiveHint=True),
    ),
    (
        layout_tools.set_page_layout,
        "set_page_layout",
        ToolAnnotations(title="Set Page Layout", destructiveHint=True),
    ),
    (
        layout_tools.add_header_footer,
        "add_header_footer",
        ToolAnnotations(title="Add Header/Footer", destructiveHint=True),
    ),
    (
        layout_tools.add_page_numbers,
        "add_page_numbers",
        ToolAnnotations(title="Add Page Numbers", destructiveHint=True),
    ),
    (
        layout_tools.add_section_break,
        "add_section_break",
        ToolAnnotations(title="Add Section Break", destructiveHint=True),
    ),
    (
        layout_tools.set_paragraph_spacing,
        "set_paragraph_spacing",
        ToolAnnotations(title="Set Paragraph Spacing", destructiveHint=True),
    ),
    (
        layout_tools.add_bookmark,
        "add_bookmark",
        ToolAnnotations(title="Add Bookmark", destructiveHint=True),
    ),
    (
        layout_tools.add_watermark,
        "add_watermark",
        ToolAnnotations(title="Add Watermark", destructiveHint=True),
    ),
    (
        content_tools.add_table_of_contents,
        "add_table_of_contents",
        ToolAnnotations(title="Add Table of Contents", destructiveHint=True),
    ),
    (
        document_tools.merge_documents,
        "merge_documents",
        ToolAnnotations(title="Merge Documents", destructiveHint=True),
    ),
)


def register_tools(server: FastMCP) -> FastMCP:
    """Register every public tool on *server* and return it."""
    for function, name, annotations in TOOL_SPECS:
        server.tool(name=name, annotations=annotations)(function)
    return server
