---
name: word-ai
description: Safe Microsoft Word DOCX editing through the Word AI MCP server and Office.js bridge. Use when Codex needs to inspect, modify, validate, diff, audit, or roll back Word `.docx` files; when a user asks to edit the currently open Word document; when content controls, OOXML, PatchSet, Office.js, document structure preservation, or Word delivery documents are involved.
---

# Word AI

Use Word AI to edit existing Word documents without rebuilding them. The default contract is: read broadly, write narrowly, prefer content controls, require PatchSet preconditions, dry-run before file writes, validate and diff after writes, and report audit/rollback details.

## Mode Selection

Choose one mode before editing:

| Situation | Use | Notes |
|---|---|---|
| User gives a `.docx` path or asks to edit a file on disk | Offline `docx_*` tools | Formal file writes create a new DOCX by default. |
| User says “current/open Word document”, “Word session”, “taskpane”, or wants edits inside Word | Live `word_session_*` tools | Requires the Office.js taskpane connected to the local bridge. |
| User asks for visual/render QA only | Word AI validation first, optional OfficeCLI read-only rendering | Rendering is evidence, not the write path. |
| No content-control anchor exists for a broad edit | Add/recommend anchors first | Use high-risk operations only after assess + dry-run and explicit scope. |

Never silently fall back from live Word session editing to offline file editing. If no active session exists, say that Word must have the taskpane open and connected, then offer offline file editing only when the user provides a file path.

For file paths outside the repository, use Word AI only when the path is inside the configured primary `--root` or an explicit `--allow-root` directory. Common user document folders such as Downloads, Documents, and Desktop should be added as allowed roots in Codex config instead of copying source files into the repository.

## Offline Engine

Offline `docx_*` PatchSet transactions default to the .NET Open XML backend when available. Python is the MCP facade, read/index layer, Office.js bridge runtime, and fallback/reference path.

- Default: `WORD_AI_ENGINE=auto`, which selects .NET native executable, .NET DLL, or local .NET project before falling back to Python.
- Production: prefer `WORD_AI_ENGINE=dotnet` so a missing .NET backend fails fast.
- Development fallback: use `engine: "python"` only for comparison or when .NET is unavailable.

## Offline DOCX Workflow

Use this for files on disk:

1. `docx_health_check`.
2. `docx_map`, `docx_list_anchors`, `docx_list_content_controls`, `docx_search_text`, or table/field/image/comment readers as needed.
3. Read only the target scope with `docx_read_content_control`, `docx_read_anchor`, `docx_read_paragraph`, `docx_read_table_cell`, `docx_read_table`, or `docx_read_heading_section`.
4. Build a strict PatchSet with `source_sha256` when available and target-level `expected_old_sha256` or `expected_old_text`.
5. `docx_assess_patchset`.
6. `docx_dry_run_patchset` or `docx_preflight_patchset`.
7. `docx_backup`.
8. `docx_apply_patchset` to a new output file unless overwrite was explicitly approved.
9. `docx_validate` or `docx_compare_structure`, then `docx_text_diff`.

Preferred operation order:

1. `replace_content_control_text`
2. `replace_text_in_content_control`
3. `append_content_control_text` / `prepend_content_control_text`
4. `replace_table_cell_text`
5. `replace_paragraph_text`
6. `insert_paragraph_after` / `insert_paragraph_before`
7. `append_table_row`
8. `wrap_paragraph_with_content_control`
9. `add_comment`

## Live Word Session Workflow

Use this for the document currently open in Microsoft Word:

1. `word_session_list`.
2. If needed, `word_session_refresh` or `word_session_snapshot`.
3. `word_session_read_content_control` for each target tag.
4. Build a strict PatchSet using live `text_sha256` values as `expected_old_sha256`.
5. `word_session_preview_patchset`.
6. `word_session_apply_patchset`.
7. Return the live audit, touched content-control tags, validation result, and rollback PatchSet.
8. Use `word_session_rollback` only when the user asks to revert.

Live session writes are intentionally narrower than offline file writes. They only support content-control text operations:

- `replace_content_control_text`
- `append_content_control_text`
- `prepend_content_control_text`
- `replace_text_in_content_control`

Do not use live session tools for fields, styles, numbering, relationship edits, images, table structure, or headers/footers.

## PatchSet Requirements

Use this shape:

```json
{
  "schema_version": "2.0",
  "strict": true,
  "source_sha256": "<offline-source-docx-sha256-if-file-mode>",
  "reason": "user-requested edit",
  "guard": {
    "require_preconditions": true,
    "allow_overwrite": false
  },
  "operations": [
    {
      "op": "replace_content_control_text",
      "tag": "WORD-AI:SRS:1.0:overview",
      "expected_old_sha256": "<target-text-sha256>",
      "text": "New text",
      "preserve_style": true,
      "allow_complex_content": false
    }
  ]
}
```

For high-risk operations, include `expected_old_sha256` or `expected_old_text`. For table operations, first call `docx_list_tables` and `docx_read_table` or `docx_read_table_cell`.

## OfficeCLI Auxiliary Backend

OfficeCLI may be used only as optional read-only or low-risk evidence. It is not the default write path.

Allowed OfficeCLI commands:

- `officecli view <file> html`
- `officecli view <file> screenshot`
- `officecli view <file> issues`
- `officecli query <file> <selector> --json`
- `officecli validate <file>`
- `officecli help ...` when checking syntax for the allowed commands

When the Word AI MCP server exposes OfficeCLI wrappers, prefer those wrappers over direct shell commands:

- `officecli_view_html`
- `officecli_view_screenshot`
- `officecli_view_issues`
- `officecli_query`
- `officecli_validate`

`officecli_view_screenshot` writes a PNG sidecar and should be treated like other sidecar export tools.

Forbidden by default:

- `officecli set`
- `officecli add`
- `officecli remove`
- `officecli move`
- `officecli swap`
- `officecli batch`
- `officecli raw-set`
- `officecli create` for replacing an existing deliverable
- `officecli merge`, `dump`, or any command that mutates the DOCX

Use OfficeCLI rendering to catch layout or visual issues after Word AI validation, especially overflow, placeholder leakage, or obvious formatting problems. Treat OfficeCLI output as advisory; the official structural gate remains Word AI `docx_validate` / `docx_compare_structure`.

## Help-First Rules

- For Word AI tools, follow the repository docs: `docs/TOOL_CONTRACT.md`, `docs/CODEX_TOOL_CATALOG.md`, and `AGENTS.md`.
- For OfficeCLI auxiliary reads, run `officecli help` instead of guessing selectors or command flags.
- Prefer semantic anchors: content-control tag > heading anchor > bookmark > paraId paragraph > paragraph index > raw search match.
- Do not expose whole long documents to the model when a target anchor or section can be read directly.

## Stop Conditions

Stop and report risk before writing when:

- The target contains images, fields, comments, tracked changes, nested tables, or nested content controls.
- `docx_assess_patchset` reports errors or unapproved high-risk changes.
- Dry-run, live preview, apply validation, or diff checks fail.
- The requested change requires styles, numbering, relationships, headers, footers, fields, image relationships, section breaks, or full-document rebuilds.
- The user asks to overwrite the source file without explicit approval and backup.

## Completion Report

After a successful edit, report:

- Mode used: offline `docx_*` or live `word_session_*`.
- Anchors, tags, tables, paragraphs, or objects touched.
- Whether only authorized DOCX parts or live content controls changed.
- Stability of content-control tags, tables, images, fields, comments, revisions, and relationships.
- Output DOCX path and audit JSON path for offline edits, or live audit and rollback PatchSet for Word session edits.
- Diff summary and any residual risk.
