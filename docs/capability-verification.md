# Capability verification results

Generated: 2026-08-31 15:48:47 -0700

This report records actual calls through `fastmcp.Client` using its in-memory
transport. Live cases used a private, hidden Microsoft Word COM instance and
disposable documents. It does not claim stdio/HTTP transport verification.
Each tool receives one explicit representative case; PASS does not mean every
argument combination or multi-operation branch has been verified.

- Registered tools: 78
- Cases run: 78
- Passed: 78
- Failed: 0
- Not run: 0
- Stateful workflows passed: 2 / 2
- Word version: 16.0
- Python: 3.14.7 (main, Aug 25 2026, 14:01:25) [MSC v.1944 64 bit (AMD64)]
- Package version: 0.1.1
- Git commit: 7c3c69caced32a51cec0ef33bc3e4a2e454a41e0
- Git dirty: True
- Source fingerprint: `e488d9b8b832514f63510192b2cb1b85b08410cce1f9b6db2033b419df07626d`
- Verifier SHA-256: `33a09c01554ea8a116fdad4214efa17d354a02c5445879fb85ab7a38a5fba0aa`

| Tool | Status | MCP error | Nested error | Postcondition |
| --- | --- | --- | --- | --- |
| `word_live_edit_batch` | PASS | False |  | document contains 'Batch verified'=True; MCP images=1 |
| `word_live_capture_pages` | PASS | False |  | captured_pages=[1]; MCP images=1 |
| `word_live_add_comment` | PASS | False |  | Word comment count=2 |
| `word_live_delete_comment` | PASS | False |  | Word comment count=0 |
| `word_live_get_comments` | PASS | False |  | payload threads=1; raw Word comments=2 |
| `word_live_reply_to_comment` | PASS | False |  | Word reply count=1 |
| `word_live_set_comment_status` | PASS | False |  | payload resolved=True; Word Done=True |
| `word_live_delete_text` | PASS | False |  | document excludes 'Gamma marker'=True |
| `word_live_insert_equation` | PASS | False |  | Word equation count=1 |
| `word_live_insert_image` | PASS | False |  | Word inline_shapes=1; shapes=0 |
| `word_live_insert_page_break` | PASS | False |  | Word pages=2; page break present=True |
| `word_live_insert_paragraphs` | PASS | False |  | document contains 'Inserted paragraph'=True |
| `word_live_insert_text` | PASS | False |  | document contains 'Inserted marker'=True |
| `word_live_replace_text` | PASS | False |  | document contains 'Alpha replaced'=True |
| `word_live_insert_cross_reference` | PASS | False |  | Word main-story fields=1 |
| `word_live_list_cross_reference_targets` | PASS | False |  | native heading targets=1 |
| `convert_to_pdf` | PASS | False |  | PDF exists=True; bytes=113719 |
| `word_live_list_fields` | PASS | False |  | listed fields=2 |
| `word_live_unlink_fields` | PASS | False |  | unlinked fields=2; Word main fields=0 |
| `word_live_update_fields` | PASS | False |  | updated fields=1; result='8/31/2026' |
| `copy_document` | PASS | False |  | copied file exists=True |
| `create_document` | PASS | False |  | created file exists=True |
| `get_document_info` | PASS | False |  | document='C:\\Users\\user18\\AppData\\Local\\Temp\\word-mcp-capabilities-xm0h4uyr\\closed-source.docx'; body_paragraphs=3 |
| `list_available_documents` | PASS | False |  | listed count=3 |
| `word_live_format_text` | PASS | False |  | Word bold=-1; size=15.0 |
| `word_live_get_paragraph_format` | PASS | False |  | payload success=True |
| `word_live_edit_headers_footers` | PASS | False |  | Word footer fields=2; text='Page 1 of 1\r' |
| `word_live_get_headers_footers` | PASS | False |  | returned native header and footer text |
| `word_live_inspect_highlighted_text` | PASS | False |  | highlighted ranges=1 |
| `word_live_add_hyperlink` | PASS | False |  | Word hyperlinks=1 |
| `word_live_list_hyperlinks` | PASS | False |  | payload hyperlinks=1 |
| `word_live_remove_hyperlink` | PASS | False |  | Word hyperlinks=0; text preserved=True |
| `word_live_update_hyperlink` | PASS | False |  | Word address='https://example.com/updated' |
| `word_live_find_text` | PASS | False |  | matches=1 |
| `word_live_get_info` | PASS | False |  | payload success=True |
| `word_live_get_page_text` | PASS | False |  | page result contains Alpha marker |
| `word_live_get_text` | PASS | False |  | Alpha marker present in returned paragraphs |
| `word_live_add_bookmark` | PASS | False |  | Word bookmark exists=True |
| `word_live_add_section_break` | PASS | False |  | Word section count=2 |
| `word_live_add_watermark` | PASS | False |  | Word header shape count=1 |
| `word_live_set_page_layout` | PASS | False |  | Word orientation=1; left_margin=57.599998474121094 |
| `word_live_set_paragraph_spacing` | PASS | False |  | Word space_after=18.0; alignment=1 |
| `word_live_inspect_layout` | PASS | False |  | pages=3; sections=2; manual_offsets=[29]; expected_manual_offsets=[29]; section_offsets=[15]; section2 start/end pages=2/3 |
| `word_live_close` | PASS | False |  | closed=True; saved file exists=True |
| `word_live_get_undo_history` | PASS | False |  | payload success=True |
| `word_live_list_open` | PASS | False |  | reported open documents=1 |
| `word_live_open` | PASS | False |  | Word opened exact path=True |
| `word_live_rename` | PASS | False |  | active path=C:\Users\user18\AppData\Local\Temp\word-mcp-capabilities-xm0h4uyr\rename-destination.docx; destination exists=True; source removed=True; content preserved=True |
| `word_live_save` | PASS | False |  | Word Saved=True |
| `word_live_set_core_properties` | PASS | False |  | Word title='Verified title' |
| `word_live_undo` | PASS | False |  | document excludes 'Undo target marker'=True |
| `word_live_insert_file` | PASS | False |  | preserved={'sections': True, 'styles': True, 'fields': True, 'comments': True, 'footnotes': True, 'endnotes': True, 'revisions': True}; inserted_range={'start': 73, 'end': 194} |
| `word_live_navigate` | PASS | False |  | selection={'char_start': 28, 'char_end': 28, 'collapsed': True, 'start_page': 2, 'end_page': 2, 'active_end_page': 2}; Word Saved=True; visible=True |
| `word_live_edit_footnotes_endnotes` | PASS | False |  | Word footnote count=2 |
| `word_live_get_note_configuration` | PASS | False |  | configured footnotes=1; Word footnotes=1 |
| `word_live_list_footnotes_endnotes` | PASS | False |  | payload footnotes=1; Word footnotes=1 |
| `word_live_set_note_configuration` | PASS | False |  | Word footnote starting=1; rule=1; style=2; location=1; separator='Capability separator\r\r' |
| `word_live_apply_list` | PASS | False |  | Word list type=2 |
| `word_live_inspect_heading_numbering` | PASS | False |  | headings=1; heading styles=9 |
| `word_live_setup_heading_numbering` | PASS | False |  | numbered headings=1; Heading 1 format='Article %1'; Word list type=4 |
| `word_live_inspect_document_outline` | PASS | False |  | outline entries=1 |
| `word_live_accept_revisions` | PASS | False |  | Word revision count=0 |
| `word_live_list_revisions` | PASS | False |  | payload revisions=1; Word revisions=1 |
| `word_live_reject_revisions` | PASS | False |  | Word revisions=0; marker present=False |
| `word_live_toggle_track_changes` | PASS | False |  | Word TrackRevisions=True |
| `word_live_create_document_snapshot` | PASS | False |  | snapshot exists=True; Word Saved=True; section headers=['Inherited capability header', 'Inherited capability header'] |
| `word_live_diff_document_snapshots` | PASS | False |  | paragraph_change_groups=1; component_leaf_changes=0; paragraph_operations=1 |
| `word_live_create_custom_style` | PASS | False |  | Word custom character style was created with bold formatting |
| `word_live_delete_custom_style` | PASS | False |  | Word style still exists=False |
| `word_live_list_custom_styles` | PASS | False |  | custom styles=1 |
| `word_live_update_custom_style` | PASS | False |  | Word style size=14.0; italic=-1 |
| `word_live_add_table` | PASS | False |  | Word table count=2 |
| `word_live_format_table` | PASS | False |  | Word table alignment=1; font_size=10.0 |
| `word_live_modify_table` | PASS | False |  | Word cell text='Changed cell\r\x07' |
| `word_live_create_table_of_contents` | PASS | False |  | Word TOCs=1 |
| `word_live_delete_table_of_contents` | PASS | False |  | Word TOCs after delete=0 |
| `word_live_list_tables_of_contents` | PASS | False |  | payload TOCs=1; Word TOCs=1 |
| `word_live_update_table_of_contents` | PASS | False |  | Word TOCs after update=1 |

## Stateful workflow results

| Workflow | Status | Steps | Evidence | Error |
| --- | --- | --- | --- | --- |
| stateful single-document edit, save, and reopen | PASS | 14/14 | All cumulative text, formatting, comment, endnote, footer field, bookmark, property, table, save, public close-with-save, reopen, and implicit active-document checks passed; immediate Word Saved flag=True |  |
| multiple open documents and target selection | PASS | 6/6 | Active-document, unique-basename, full-path, isolated mutation, listing, and activation-switch checks passed across three simultaneously open documents |  |

Full arguments, response excerpts, timings, and errors are in [`capability-verification.json`](capability-verification.json).
