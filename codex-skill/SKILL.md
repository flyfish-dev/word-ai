---
name: word-ai
description: Safe Microsoft Word DOCX editing through the Word AI MCP server and Office.js bridge. Use when Codex needs to inspect, modify, validate, diff, audit, or roll back Word `.docx` files; when a user asks to edit the currently open Word document; when content controls, OOXML, PatchSet, Office.js, document structure preservation, or Word delivery documents are involved.
---

# Word AI

This is a compatibility copy of the formal Word AI Codex Skill. The canonical project copy lives in `skills/word-ai/SKILL.md`.

## Choose The Editing Mode

| Situation | Use | Rule |
| --- | --- | --- |
| The user gives a `.docx` file path or wants a new output file | Offline `docx_*` tools | Default path for delivery documents and batch work. |
| The user says the document is open in Word, asks for the current Word document, or needs visible Word session editing | Live `word_session_*` tools | Requires the Office.js taskpane connected to the local bridge. |
| The user needs visual/render evidence only | Word AI validation first, optional OfficeCLI read-only commands | Rendering is evidence, not the write path. |

Never silently fall back from a requested live Word session to offline file editing. If no live session exists, ask the user to load the Office.js add-in and connect the bridge.

For file paths outside the repository, use Word AI only when the path is inside the configured primary `--root` or an explicit `--allow-root` directory. Add common user document folders such as Downloads, Documents, and Desktop to Codex config instead of copying source files into the repository.

## Offline Engine

Offline `docx_*` PatchSet transactions default to the .NET Open XML backend when available. Python is the MCP facade, read/index layer, Office.js bridge runtime, and fallback/reference path.

- Default: `WORD_AI_ENGINE=auto`, which selects .NET native executable, .NET DLL, or local .NET project before falling back to Python.
- Production: prefer `WORD_AI_ENGINE=dotnet` so a missing .NET backend fails fast.
- Development fallback: use `engine: "python"` only for comparison or when .NET is unavailable.

## Offline DOCX Workflow

1. Run `docx_health_check`.
2. Locate anchors with `docx_map`, `docx_list_anchors`, or `docx_list_content_controls`.
3. Read target text using `docx_read_content_control`, `docx_read_anchor`, `docx_read_table`, `docx_read_table_cell`, or a similarly scoped read tool; record `text_sha256`.
4. Build a strict PatchSet with `source_sha256` when available, `expected_old_sha256`, and `guard.require_preconditions=true`.
5. Run `docx_assess_patchset`.
6. Run `docx_dry_run_patchset`.
7. Run `docx_backup`.
8. Run `docx_apply_patchset` to a new output path.
9. Run `docx_validate`, `docx_compare_structure`, and `docx_text_diff`.

Preferred operations are `replace_content_control_text`, `append_content_control_text`, `prepend_content_control_text`, then `replace_text_in_content_control`. Use paragraph and table operations only after assessing risk and reading the exact target scope.

## Live Word Session Workflow

Use live tools only after the taskpane has registered a session:

1. `word_session_list`.
2. `word_session_snapshot` or `word_session_refresh`.
3. `word_session_read_content_control` for target content-control tags.
4. Build the same guarded PatchSet shape used offline.
5. `word_session_preview_patchset`.
6. `word_session_apply_patchset` only after preview passes and approval is available.
7. Report the returned audit and rollback PatchSet. Use `word_session_rollback` only by explicit request or approved recovery.

Live writes must stay limited to content-control text operations unless the implementation wraps broader operations in the Word AI PatchSet, dry-run, audit, rollback, and approval gates.

## PatchSet Shape

```json
{
  "schema_version": "2.0",
  "strict": true,
  "source_sha256": "<optional source sha256>",
  "reason": "user-requested edit",
  "guard": {
    "require_preconditions": true,
    "allow_overwrite": false
  },
  "operations": [
    {
      "op": "replace_content_control_text",
      "tag": "WORD-AI:SRS:1.0:overview",
      "expected_old_sha256": "<target text sha256>",
      "text": "New text",
      "preserve_style": true,
      "allow_complex_content": false
    }
  ]
}
```

## OfficeCLI Auxiliary Backend

OfficeCLI can be used only as optional read-only or low-risk evidence:

- `officecli view <file> html`
- `officecli view <file> screenshot`
- `officecli view <file> issues`
- `officecli query <file> <selector> --json`
- `officecli validate <file>`
- `officecli help ...` for allowed command syntax

When available, use the Word AI MCP wrappers instead of direct shell commands: `officecli_view_html`, `officecli_view_screenshot`, `officecli_view_issues`, `officecli_query`, and `officecli_validate`. `officecli_view_screenshot` writes a PNG sidecar and should be treated like other sidecar export tools.

Do not use OfficeCLI mutation commands by default: `officecli set`, `officecli add`, `officecli remove`, `officecli move`, `officecli swap`, `officecli batch`, `officecli raw-set`, `officecli create` for replacing a deliverable, `officecli merge`, `dump`, or any command that mutates a DOCX. These may be considered only if wrapped inside Word AI PatchSet, dry-run, audit, rollback, and explicit approval.

Borrow OfficeCLI ideas such as schema/help-first usage, semantic paths, watch/render evidence, template merge concepts, and dump/batch inspection, but keep Word AI's PatchSet safety model as the authority.

## Completion Report

Tell the user which anchors or objects changed, whether only allowed DOCX parts changed, whether tags/tables/images/fields/comments/revisions stayed stable, and provide the output DOCX path, audit JSON path, and diff summary.
