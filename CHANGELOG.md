# Changelog

All notable changes to this project are documented here.

## [0.1.2] - 2026-09-01

### Changed

- Rebuilt the inherited 118-name surface as 78 reviewed, capability-based tools.
- Upgraded and pinned the runtime to FastMCP 4.0.0 and MCP SDK 2.x, including
  migration to FastMCP 4's public tool imports and snake-case protocol fields.
- Flattened the one-child `tools/live/` package into domain modules directly under
  `tools/`; live behavior remains explicit in public names and tool metadata.
- Limited closed-file operations to safe lifecycle tasks: create, copy, list,
  body-only metadata, and native Word PDF export.
- Replaced separate destructive header/footer and page-number tools with
  `word_live_get_headers_footers` and `word_live_edit_headers_footers`.
- Standardized public paragraph, table, note, section, and page indexes on
  one-based Word conventions.
- Made batch rollback conditional on positively identifying the batch's own
  latest Word Undo entry.

### Added

- Safe live-document close with explicit save, discard, or require-saved policy.
- Verified live-document rename/move that refuses destination overwrite and preserves
  the document format.
- Native field inspection/update across populated Word stories.
- Native table-of-contents create/list/update/delete operations.
- Native hyperlink list/add/update/remove operations with text-preserving unlink.
- Native cross-reference target discovery and insertion without heuristic IDs.
- Native page-break insertion with paragraph or character targeting.
- Native Heading 1-9 multilevel numbering inspection and guarded setup.
- Native top-level comment-thread resolve and reopen support.
- Custom paragraph/character style list, create, update, and delete operations.
- Document outline and cross-story highlighted-text inspection.
- Native footnote/endnote numbering, placement, and separator configuration.
- Objective section geometry, page-break, and paragraph pagination-control inspection.
- Explicit native field unlinking with unsupported-field preflight checks.
- Native `Range.InsertFile` document insertion with explicit targeting, Word-managed
  conflict semantics, explicit native-Undo residue reporting, and preservation
  verification for sections, styles, fields, comments, notes, and revisions.
- Persisted, versioned semantic document snapshots with integrity checks, explicit
  overwrite behavior, sensitive-content warnings, and deterministic snapshot diffing.
- Validated page and UTF-16 range navigation that reports final Word selection state
  without changing application visibility or the document's saved state.
- Header/footer inspection for every section and primary, first-page, and
  even-page story.
- Header/footer templates with `{page}`, `{pages}`, and `{section_pages}` Word
  fields; Arabic, Roman, and letter numbering; restart/start-at controls;
  section linkage; alignment; font; size; bold; italic; and color.
- Targeted section-break insertion by character offset or paragraph index.
- Regression coverage for ambiguous open-document names, lifecycle overwrite
  guards, native note editing, threaded-comment deletion, batch safety, and
  header/footer validation.

### Fixed

- Full-path live operations no longer select the wrong same-named open document;
  ambiguous basename-only requests now fail with candidate paths.
- `word_live_delete_text` no longer deletes past its requested range after table
  offsets shift; ranges intersecting tables are rejected explicitly.
- Cursor insertion now verifies that Word's selection belongs to the requested
  document.
- Paragraph-index chaining between `word_live_get_text` and
  `word_live_insert_paragraphs` now targets the same paragraph.
- Styles are applied before explicit text-format overrides, with all relevant
  formatting inputs validated before mutation.
- Threaded comments are deleted with Word's recursive thread operation.
- Table, list, layout, image, save-as, and page-number inputs now reject invalid
  values instead of silently defaulting or partially mutating a document.
- DOCX extension checks now handle uppercase `.DOCX` correctly.

### Removed

- The one-time catalog audit command, test, and workflow scaffold.
- Lossy fake table-of-contents and document-merge implementations.
- Unsafe raw-OOXML tracked changes and broken offline comments.
- Duplicate closed-file content, formatting, layout, table, and note wrappers.
- Volatile in-memory snapshot/diff state, heuristic layout diagnosis, unreliable
  cross-reference discovery, and the inherited risky heading-numbering and
  unverified Modern Comments helpers; verified replacements now exist.
- Inherited macOS/JXA compatibility code and obsolete container/service files.

## [0.1.0]

- Forked and renamed the distribution as `word-mcp-codemode-live`.
- Migrated to a uv-native `src/` layout with FastMCP Code Mode as the default.

[0.1.2]: https://github.com/chikingsley/word-mcp-codemode-live/compare/7c3c69c...v0.1.2
[0.1.0]: https://github.com/chikingsley/word-mcp-codemode-live/releases/tag/v0.1.0
