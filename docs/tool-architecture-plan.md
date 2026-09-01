# Tool architecture and capability plan

Status: active implementation roadmap. Registration and source-layout migration are
complete; capability recovery is in progress.

## Purpose

The repository needs a tool organization based on stable Word capabilities rather
than the history of individual implementations. The immediate goals are to:

- make every public tool belong to one clear domain;
- keep the MCP adapter, validation, Word operation, and result close enough to read
  as one vertical slice;
- make each tool definition the source of its registration metadata and batch eligibility;
- distinguish an intentionally unsupported workflow from a capability that was
  merely lost during cleanup;
- permit missing capabilities to return without rebuilding catch-all modules; and
- preserve current public names while the internal structure is migrated.

## First principles

1. **Organize by user capability.** A module answers a recognizable Word question
   such as "comments," "tables," or "headers and footers." Names such as
   `extended_document_tools` and catch-alls such as `live_tools` are not domains.
2. **Keep one vertical slice together.** Tool input models, output models, validation,
   and domain-specific COM logic start in the same domain module. Extract code only
   after it is shared or independently complex.
3. **Share infrastructure, not every noun.** Connection management, document
   resolution, ranges, transactions, constants, and rendering belong below the tool
   layer. Do not create matching `models`, `services`, and `adapters` trees for every
   domain.
4. **The tool definition is authoritative.** Registration, annotations, Code Mode
   discovery, batch eligibility, inventory generation, and contract tests derive
   from metadata beside the function. Do not maintain a second catalog of the same
   tools. FastMCP's provider is the runtime registry.
5. **The wire contract is typed.** Tools return dictionaries, dataclasses, Pydantic
   models, or an explicit `ToolResult`; they do not return JSON encoded inside a
   string. Correctable failures raise tool errors rather than returning successful
   `{"error": ...}` payloads.
6. **Public names and Python layout are separate decisions.** Moving a function to a
   coherent module does not require changing its MCP name.
7. **Native Word behavior is the default.** File-based DOCX operations remain limited
   to lifecycle and inspection cases that have a clear reason not to use Word.
8. **No silent capability loss.** Every inherited or proposed capability receives a
   disposition: active, superseded, backlog, internal-only, or rejected with a
   rationale.

## Current baseline

The worktree currently registers 78 reviewed tools. It also contains two non-tool
page-rendering helpers. On 2026-08-31, all 78 tools passed one explicit representative
case through `fastmcp.Client`; live cases used a private Microsoft Word 16.0 COM
instance and disposable documents. The evidence is recorded in
[`capability-verification.md`](capability-verification.md) and
[`capability-verification.json`](capability-verification.json). A 14-step cumulative
workflow also verified implicit active-document editing, multiple interacting
artifacts, explicit save, close without saving, reopen, and persistence. A separate
six-step workflow verified active, basename, and full-path targeting across three
simultaneously open documents without mutation leakage. This establishes a tested
baseline, not exhaustive coverage of every argument combination or every operation
exposed by consolidated tools. The current public tool inventory is below; target
modules describe the intended internal home and do not rename the MCP tool.

### Closed-file lifecycle and export

Target modules: `tools/files.py`, `tools/export.py`

| Public tool | Status | Target module |
| --- | --- | --- |
| `create_document` | active | `tools/files.py` |
| `copy_document` | active | `tools/files.py` |
| `get_document_info` | active | `tools/files.py` |
| `list_available_documents` | active | `tools/files.py` |
| `convert_to_pdf` | active | `tools/export.py` |

### Live lifecycle

Target modules: `tools/lifecycle.py`, `tools/navigation.py`

| Public tool | Status |
| --- | --- |
| `word_live_list_open` | active |
| `word_live_open` | active |
| `word_live_save` | active |
| `word_live_close` | active; require-saved default with explicit save/discard policy |
| `word_live_rename` | active; verified destination, no overwrite, old path removed |
| `word_live_undo` | active |
| `word_live_get_undo_history` | active |
| `word_live_set_core_properties` | active; keep here unless properties grow into a cohesive domain |
| `word_live_navigate` | active; validated page or UTF-16 main-story range selection without saved-state mutation |

### Live inspection

Target modules: `tools/inspection.py`, `tools/snapshots.py`

| Public tool | Status |
| --- | --- |
| `word_live_get_text` | active |
| `word_live_get_info` | active |
| `word_live_find_text` | active |
| `word_live_get_page_text` | active |
| `word_live_create_document_snapshot` | active; persisted versioned semantic baseline with integrity metadata |
| `word_live_diff_document_snapshots` | active; deterministic persisted-baseline comparison, not visual equivalence |

### Live content

Target modules: `tools/content.py`, `tools/merge.py`

| Public tool | Status |
| --- | --- |
| `word_live_insert_text` | active |
| `word_live_replace_text` | active |
| `word_live_insert_paragraphs` | active |
| `word_live_delete_text` | active |
| `word_live_insert_image` | active |
| `word_live_insert_equation` | active |
| `word_live_insert_page_break` | active; native page break with explicit range targeting |
| `word_live_insert_file` | active; native `Range.InsertFile` with explicit target and Word-managed conflicts |

### Live formatting and numbering

Target modules: `tools/formatting.py`, `tools/numbering.py`

| Public tool | Status | Target module |
| --- | --- | --- |
| `word_live_format_text` | active | `formatting.py` |
| `word_live_get_paragraph_format` | active | `formatting.py` |
| `word_live_set_paragraph_spacing` | active | `formatting.py` |
| `word_live_apply_list` | active | `numbering.py` |
| `word_live_inspect_heading_numbering` | active | `numbering.py` |
| `word_live_setup_heading_numbering` | active; preserves existing schemes unless replacement is explicit | `numbering.py` |

### Live styles

Target module: `tools/styles.py`

| Public tool | Status |
| --- | --- |
| `word_live_list_custom_styles` | active |
| `word_live_create_custom_style` | active |
| `word_live_update_custom_style` | active; custom styles only |
| `word_live_delete_custom_style` | active; custom styles only |

### Structural inspection

Target modules: `tools/outline.py`, `tools/highlights.py`,
`tools/layout_diagnostics.py`

| Public tool | Status | Target module |
| --- | --- | --- |
| `word_live_inspect_document_outline` | active | `outline.py` |
| `word_live_inspect_highlighted_text` | active; traverses populated Word stories | `highlights.py` |
| `word_live_inspect_layout` | active; reports objective geometry and pagination controls | `layout_diagnostics.py` |

### Live tables

Target module: `tools/tables.py`

| Public tool | Status |
| --- | --- |
| `word_live_add_table` | active |
| `word_live_format_table` | active |
| `word_live_modify_table` | active; split-by-operation should be evaluated during contract redesign |

### Live comments

Target module: `tools/comments.py`

| Public tool | Status |
| --- | --- |
| `word_live_get_comments` | active |
| `word_live_add_comment` | active |
| `word_live_reply_to_comment` | active |
| `word_live_delete_comment` | active |
| `word_live_set_comment_status` | active; native top-level thread resolve/reopen |

### Live revisions

Target module: `tools/revisions.py`

| Public tool | Status |
| --- | --- |
| `word_live_toggle_track_changes` | active |
| `word_live_list_revisions` | active |
| `word_live_accept_revisions` | active |
| `word_live_reject_revisions` | active |

### Live footnotes and endnotes

Target module: `tools/notes.py`

| Public tool | Status |
| --- | --- |
| `word_live_list_footnotes_endnotes` | active |
| `word_live_edit_footnotes_endnotes` | active; operation-specific request models are required |
| `word_live_get_note_configuration` | active |
| `word_live_set_note_configuration` | active; native numbering, placement, and separator controls |

Footnotes and endnotes remain together because Word exposes parallel collections
with shared targeting and conversion behavior. The shorter internal domain name is
`notes`; the explicit public names may remain for discoverability.

### Live headers, footers, and layout

Target modules: `tools/headers_footers.py`, `tools/layout.py`

| Public tool | Status | Target module |
| --- | --- | --- |
| `word_live_get_headers_footers` | active | `headers_footers.py` |
| `word_live_edit_headers_footers` | active; replace the flat 19-argument contract | `headers_footers.py` |
| `word_live_set_page_layout` | active | `layout.py` |
| `word_live_add_section_break` | active | `layout.py` |
| `word_live_add_watermark` | active | `layout.py` |
| `word_live_add_bookmark` | active; move to references when that domain is restored | `layout.py` temporarily |

Headers and footers remain together because both are section stories with the same
primary, first-page, and even-page variants, linkage rules, and page-number fields.

### Live fields and references

Target modules: `tools/fields.py`, `tools/toc.py`,
`tools/hyperlinks.py`, `tools/cross_references.py`

| Public tool | Status |
| --- | --- |
| `word_live_list_fields` | active; all populated Word stories |
| `word_live_update_fields` | active; authoritative field-update path |
| `word_live_unlink_fields` | active; explicit conversion of field results to static content |
| `word_live_list_tables_of_contents` | active |
| `word_live_create_table_of_contents` | active; native Word TOC field |
| `word_live_update_table_of_contents` | active |
| `word_live_delete_table_of_contents` | active |
| `word_live_list_hyperlinks` | active |
| `word_live_add_hyperlink` | active |
| `word_live_update_hyperlink` | active |
| `word_live_remove_hyperlink` | active; preserves visible text |
| `word_live_list_cross_reference_targets` | active; native transient list positions |
| `word_live_insert_cross_reference` | active; native Word cross-reference field |

### Verification and orchestration

Target modules: `tools/capture.py`, `tools/batch.py`

| Public tool | Status | Target module |
| --- | --- | --- |
| `word_live_capture_pages` | active | `capture.py` |
| `word_live_edit_batch` | active | `batch.py` |

`render_word_pages` and `rendered_pages_result` are internal helpers, not public tool
candidates.

## Capability recovery ledger

Removal from the inherited 118-tool catalog is not, by itself, a final product
decision. This ledger groups missing functionality by intended disposition.

### Superseded by current native live capabilities

These old names should not return as separate tools unless a concrete user workflow
cannot be expressed by the replacement.

| Inherited capability | Disposition |
| --- | --- |
| Closed-file paragraph insertion, headings, pictures, tables, deletion, and replacement | superseded by live content tools |
| Closed-file text/table formatting wrappers | superseded by live formatting and table tools |
| Closed-file comments and raw-OOXML tracked changes | superseded by native live comments and revisions |
| Numerous robust/enhanced/add-before/add-after footnote variants | superseded by the native consolidated note tools |
| Separate header/footer and page-number tools | superseded by native header/footer inspection and editing |
| Duplicate closed-file layout, section, bookmark, and watermark wrappers | superseded by live layout tools |
| `get_comments_by_author` and paragraph comment filters | caller-side filtering of structured comment results unless scale proves a dedicated query is needed |
| Separate table width, padding, alignment, shading, merge, row, and column wrappers | consolidate under a coherent table contract; restore dedicated tools only where schema clarity requires them |
| `track_insert`, `track_delete`, and `track_replace` | use native live mutations with tracked changes; verify complete coverage before marking fully superseded |

### Recovery backlog: expected Word capabilities

These are real user-facing gaps. Names and exact contracts remain design decisions.

| Capability | Candidate domain | Priority | Acceptance requirement |
| --- | --- | --- | --- |
| Native table of contents create/update/inspect | references | completed | genuine Word fields, style/level controls, update and real-Word verification |
| Hyperlink list/add/edit/remove | references | completed | native hyperlink objects, range targeting, safe text-preserving removal |
| Cross-reference discovery/insert | references | completed | native list positions and real Word fields; generic field tools own updates |
| Heading and multilevel numbering setup/inspect | numbering | completed | native list templates, existing-number preservation, real-document tests |
| Insert page break | content | completed | exact range targeting, Undo support, and real-Word verification |
| Resolve/reopen Modern Comments | comments | completed | native top-level thread state, persistence, Undo, and real-Word verification |
| Create/update/delete custom styles | styles | completed | distinguish style definitions from direct formatting; round-trip inspection |
| Document outline/headings inspection | inspection | completed | stable heading hierarchy with positions and page information |
| Highlighted-text inspection | inspection | completed | body, tables, headers/footers, notes, and color reporting as applicable |
| Footnote/endnote numbering, placement, separator, and style controls | notes | completed | inspect-before-edit and native Word verification |
| Layout diagnostics | inspection | completed | objective findings with page/range evidence; no unsupported heuristic claims |
| Field list/update operations | references | completed | identifies field type, story, location, and update result |
| Field unlink operation | references | completed | explicit destructive semantics and real-Word verification |
| Native merge/insert-file workflow | content | completed | native `Range.InsertFile`; verified sections, styles, fields, comments, notes, revisions, and explicit native-Undo definition-residue semantics |
| Snapshot and diff | inspection | completed | versioned persisted JSON, integrity/source identity, explicit overwrite and cross-document semantics |
| Navigate Word UI to page/range | lifecycle | completed | validated rendered-page or UTF-16 range targeting with exact selection and saved-state verification |

Priority means planning order, not authorization to implement.

### Internal-only or rejected public surface

| Capability | Disposition |
| --- | --- |
| Raw document XML | internal diagnostic API or test helper, not a normal public MCP tool |
| Lossy paragraph-by-paragraph document merge | rejected |
| Fake table of contents made from copied heading text | rejected |
| Raw-OOXML mutations while Word owns the open document | rejected |
| Multiple tools differentiated only by `robust`, `enhanced`, or target-position suffixes | rejected naming pattern; use validated request variants |

## Target source tree

Create modules as capabilities are migrated; do not scaffold empty packages merely
to resemble this diagram.

```text
src/word_mcp_codemode_live/
|-- main.py                     # server assembly and provider configuration
|-- tools/
|   |-- files.py
|   |-- export.py
|   |-- lifecycle.py
|   |-- inspection.py
|   |-- content.py
|   |-- formatting.py
|   |-- numbering.py
|   |-- styles.py
|   |-- outline.py
|   |-- highlights.py
|   |-- tables.py
|   |-- comments.py
|   |-- revisions.py
|   |-- notes.py
|   |-- headers_footers.py
|   |-- layout.py
|   |-- layout_diagnostics.py
|   |-- fields.py
|   |-- toc.py
|   |-- hyperlinks.py
|   |-- cross_references.py
|   |-- capture.py
|   `-- batch.py
`-- word/
    |-- application.py         # COM attachment and lifetime
    |-- documents.py           # active/named/full-path document resolution
    |-- transactions.py        # Undo record and rollback semantics
    |-- ranges.py               # text, paragraph, page, and story targeting
    |-- constants.py            # named Word constants instead of magic integers
    `-- rendering.py            # Word PDF export and page rasterization internals
```

Domain-specific request/result models remain beside their tool functions. A shared
model moves to a common module only when at least two domains use the same contract.
The tool package is deliberately flat: the project and public names already identify
live Word operations, while `files.py` and `export.py` clearly identify the small
closed-file surface. A one-child `live/` package added hierarchy without another
peer execution package to distinguish.

## Tool definitions and discovery

There must not be separate hand-maintained catalog and registry files. Each public
function carries its FastMCP name, title, annotations, domain tags, and batch
eligibility beside its implementation. FastMCP's provider discovers and registers
those definitions.

```python
@word_tool(
    title="Word Live Replace Text",
    domain="content",
    change="edit",
    batchable=True,
)
async def word_live_replace_text(...):
    ...
```

`word_tool` is a thin project decorator around FastMCP's standalone `@tool`. It
attaches metadata once; it must not require another tuple or dictionary entry. Its
runtime batch index is populated automatically from decorated functions. Filesystem
discovery and the Code Mode transform are covered by repository tests and the stdio
verification script.

Current change kinds:

- `read`
- `safe_write`
- `edit`

MCP annotations are derived deliberately from this metadata. In particular, not
every write is automatically destructive, and not every reversible Word edit is
safe to include in a document Undo batch.

The definitions provide:

- the functions registered by FastMCP;
- tool annotations and titles;
- the batch allowlist;
- domain tags used by Code Mode search;
- a deterministic inventory for documentation; and
- contract tests that detect duplicate or stranded public tools.

Backlog entries stay in this document or a separate product ledger; they must not be
fake tool definitions or placeholder tools.

## Naming policy

### During internal migration

- Preserve active public MCP names unless a reviewed rename materially improves the
  contract.
- Rename Python modules and internal functions freely when tests cover the public
  contract.
- Do not add aliases merely to make internal names symmetrical.

### When a versioned public naming pass is approved

- Live tools continue to use `word_live_<verb>_<object>`.
- Decide together whether closed-file tools become `word_file_<verb>_<object>`;
  until then, keep their existing names.
- Use concrete verbs: `get`, `list`, `find`, `insert`, `set`, `replace`, `delete`,
  `accept`, `reject`, `capture`, `export`.
- Avoid `manage`, `modify`, `enhanced`, `robust`, `extended`, and generic `tools` in
  public names.
- A consolidated tool is appropriate only when its operation variants have one
  cohesive result and can be represented by validated discriminated request models.

## Migration plan

### Phase 0: approve the inventory

- Confirm the active tool inventory before each recovery wave.
- Review the recovery backlog and priorities.
- Record explicit rejection rationales for capabilities that should never return.
- Do not perform broad module moves until this baseline is accepted.

### Phase 1: establish co-located tool definitions

- Added the thin `word_tool` decorator with domain, change-kind, and batchability
  metadata.
- Registration metadata now lives beside all public functions.
- The batch index is generated automatically from decorated functions.
- `registry.py` was deleted; no `catalog.py` was added.
- Add tests for unique names, annotations, domain membership, and batch eligibility.
- Replace exact-count-only testing with a checked-in contract inventory plus focused
  schema assertions.

### Phase 2: migrate one reference vertical slice

- Use notes as the first slice because it already has a coherent native Word model.
- Move it to `tools/notes.py`.
- Introduce validated operation-specific input models and structured results.
- Raise proper tool errors and preserve one Undo record.
- Verify with fake-COM unit tests and a real Word integration document.
- Treat the resulting structure and tests as the template for later domains.

### Phase 3: split the catch-all modules without behavior changes

Completed capability layout:

1. `live/headers_footers.py`, `live/comments.py`, and `live/revisions.py`;
2. `tools/lifecycle.py` and `tools/inspection.py`;
3. `live/tables.py`, `live/formatting.py`, and `live/numbering.py`;
4. `live/content.py`, `live/layout.py`, `live/capture.py`, and `live/batch.py`;
5. `files.py` and `export.py`.

The catch-all modules were removed after provider discovery confirmed a unique tool
names. Public names were retained because they still describe their operations;
internal legacy filenames were not retained for compatibility.

### Phase 4: normalize tool behavior domain by domain

- Replace JSON strings with structured return types.
- Replace returned error strings with raised tool errors.
- Mark truly required arguments as required.
- Replace free-form enum strings with `Literal`, enums, or constrained fields.
- Replace large flat signatures with cohesive request models.
- Name Word constants and centralize common range/transaction behavior.
- Decide how synchronous COM work is serialized without blocking or violating COM
  apartment rules.

This is intentionally separate from the physical split so failures can be attributed
to either movement or behavior change.

### Phase 5: recover missing capabilities

- Recovery through P2 is complete: native fields, TOCs, hyperlinks,
  cross-references, page breaks, heading numbering, Modern Comment status,
  custom styles, structural inspection, note configuration, layout inspection,
  field unlinking, native file insertion, persisted semantic snapshot/diff, and
  verified non-mutating Word navigation are implemented.
- Continue implementing capabilities one vertical slice at a time.
- Require native Word behavior, Undo/rollback semantics where applicable, structured
  results, and real-document verification before registration.
- Reassess future additions based on actual document workflows rather than inherited code.

### Phase 6: remove compatibility structure

- The old document, extended-document, live-layout, live-read, and live-edit
  compatibility modules have been removed.
- Update README and changelog language after the new structure is real.
- Consider public renames only as a separately reviewed versioned change.

## Definition of done for a migrated domain

A domain is migrated only when:

- all its active public tools live in the target module;
- public names and descriptions are intentional;
- inputs are typed and constrained;
- results are structured and have stable schemas;
- failures produce `is_error=True` rather than successful error strings;
- mutation annotations and batchability match real behavior;
- document selection and range indexes follow repository-wide conventions;
- mutations have tested Undo or explicit non-batchable rationale;
- fake-COM unit tests cover validation and edge cases;
- at least one real Word integration test verifies the native artifact when the
  operation is not reliably represented by fakes; and
- no duplicate registration, duplicate batch entry, or stranded public function
  remains.

## Decisions still requiring agreement

1. Which recovery backlog items are required for the next release rather than later?
2. Should `word_live_modify_table` remain an operation switch or become several
   focused table tools?
3. Should note mutations remain one discriminated tool or become add/delete/convert
   tools?
4. Should header/footer mutation remain one discriminated tool or split content,
   linkage, formatting, and numbering operations?
5. Are closed-file public names retained indefinitely or renamed in a versioned
   release?
6. Is UI navigation an intended product capability or an implementation side effect?
7. What is the supported concurrency model for Word COM calls?

These decisions should be made from representative user workflows and generated MCP
schemas, not file-name symmetry.
