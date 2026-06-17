# Codex Tool Catalog v0.7

本文件由 `word_ai_mcp.server` 的 `tools/list` 生成，当前可发现工具数：**63**。

## 工具列表

1. `docx_health_check` — Read-only. Conservative stability report: duplicate tags, complex content controls, fields, tracked changes, comments, table previews and recommended safety policy.
2. `docx_inspect` — Read-only. Inspect counts, anchors and optional visible text.
3. `docx_package_manifest` — Read-only. DOCX ZIP manifest with part sizes and optional SHA-256 hashes.
4. `docx_list_parts` — Read-only. List all package parts with sizes and hashes.
5. `docx_part_hashes` — Read-only. Compact DOCX part-hash map for before/after invariance checks.
6. `docx_structural_fingerprint` — Read-only. Strong structure fingerprint: package parts, content-control hashes, paragraph hashes, table hashes, outline and health report.
7. `docx_map` — Read-only. Codex-friendly map: health, headings, anchors, paragraph chunks, table summaries and hashes.
8. `docx_get_outline` — Read-only. Heading outline with section paragraph ranges for large-document chunking.
9. `docx_outline` — Read-only alias for docx_get_outline; kept for Codex prompt ergonomics.
10. `docx_list_headings` — Read-only. Heading-only view with previews and hashes.
11. `docx_read_heading_section` — Read-only. Read one heading section by heading anchor_id or exact heading text.
12. `docx_list_anchors` — Read-only. List stable editing anchors: content controls, headings, bookmarks and paraId paragraphs.
13. `docx_read_anchor` — Read-only. Read text behind an anchor_id from docx_list_anchors.
14. `docx_search_text` — Read-only. Search visible paragraph text; returns paraId, paragraph index, scope and hashes.
15. `docx_list_paragraphs` — Read-only. Paragraph inventory with style, paraId, heading level, scope, preview, hash and complexity.
16. `docx_read_paragraph` — Read-only. Read one paragraph by paragraph_index or paraId and return hash/complexity.
17. `docx_list_content_controls` — Read-only. List content controls with tag, alias, id, preview, text hash and complexity.
18. `docx_read_content_control` — Read-only. Read content-control text by w:tag. Preferred primitive for safe edits.
19. `docx_extract_plain_text` — Read-only. Extract visible body text for coarse review; never use to regenerate DOCX.
20. `docx_export_plain_text` — Write-sidecar only. Export visible body text to .txt; does not modify DOCX.
21. `docx_list_tables` — Read-only. List table dimensions, scope tags, complexity and preview rows.
22. `docx_read_table` — Read-only. Read a table matrix by table_index and optional content-control scope_tag.
23. `docx_read_table_cell` — Read-only. Read a single table cell and return hash for write preconditions.
24. `docx_export_table_csv` — Write-sidecar only. Export table to CSV; does not modify DOCX. Uses default sidecar naming when out_path is omitted.
25. `docx_table_to_csv` — Write-sidecar only. Export table to CSV using default naming when out_path is omitted.
26. `docx_list_styles` — Read-only. List Word styles from styles.xml for style-aware generation.
27. `docx_list_numbering` — Read-only. Inspect numbering.xml numId/abstractNum mappings.
28. `docx_list_sections` — Read-only. Inspect section properties, margins, page size and header/footer refs.
29. `docx_list_fields` — Read-only. List field instructions such as TOC, REF, PAGEREF and SEQ.
30. `docx_list_images` — Read-only. List drawing/pict objects with relationship IDs and alt metadata.
31. `docx_list_hyperlinks` — Read-only. List hyperlinks with relationship targets, preview and hashes.
32. `docx_list_comments` — Read-only. Extract comments.xml previews, author/date/id and hashes.
33. `docx_extract_comments` — Read-only. Extract full comment metadata and text.
34. `docx_list_bookmarks` — Read-only. List bookmarks and paths for cross-reference preservation.
35. `docx_list_headers_footers` — Read-only. Inspect header/footer parts and visible text without modifying them.
36. `docx_list_notes` — Read-only. Inspect footnotes/endnotes when present.
37. `docx_list_revisions` — Read-only. List tracked-change nodes with previews and hashes.
38. `docx_assess_patchset` — Read-only. Resolve PatchSet targets and report risks, touched objects and precondition gaps. Accepts `engine=auto|dotnet|python`; default auto selects .NET when available.
39. `docx_plan_patchset` — Read-only. Alias-grade PatchSet planning assessment. Accepts `engine=auto|dotnet|python`.
40. `docx_preflight_patchset` — Write-temp. Dry-run PatchSet to a temporary copy, validate, and remove output unless keep_output=true. Accepts `engine=auto|dotnet|python`.
41. `docx_dry_run_patchset` — Write-temp. Same as preflight; useful for explicit Codex dry-run stage. Accepts `engine=auto|dotnet|python`.
42. `docx_write_index` — Write-sidecar only. Write .wordai index JSON; does not modify DOCX.
43. `docx_backup` — Write-sidecar only. Create timestamped backup; does not modify original DOCX.
44. `docx_restore_backup` — Write. Restore a backup to target_path.
45. `docx_rollback` — Write. Restore from backup and optionally back up replaced file first.
46. `docx_apply_patchset` — Write. Apply constrained PatchSet to a new DOCX plus audit JSON; validation gates final commit. Accepts `engine=auto|dotnet|python`; default auto selects .NET when available.
47. `docx_validate` — Read-only. Validate structural invariants. Supply touched_* for intentional edits. Accepts `engine=auto|dotnet|python`; default auto selects .NET when available.
48. `docx_compare_structure` — Read-only. Same validation report phrased as structural comparison. Accepts `engine=auto|dotnet|python`.
49. `docx_text_diff` — Read-only. Unified visible-text diff for human review after validation.
50. `word_session_list` — Read-only. List active Office.js taskpane sessions registered by an open Word document.
51. `word_session_snapshot` — Read-only. Return the latest content-control snapshot for an open Word session. If session_id is omitted, uses the most recently active session.
52. `word_session_refresh` — Read-only. Ask the Office.js taskpane to refresh its open-document content-control snapshot.
53. `word_session_read_content_control` — Read-only. Ask the Office.js taskpane to read content-control text from the currently open Word document by tag.
54. `word_session_preview_patchset` — Read-only. Ask the Office.js taskpane to validate and preview a PatchSet against the currently open Word document without modifying it.
55. `word_session_apply_patchset` — Write. Apply a supported content-control PatchSet to the currently open Word document through Office.js, with hash preconditions, audit, and rollback PatchSet.
56. `word_session_wrap_selection` — Write. Ask the Office.js taskpane to wrap the current Word selection in a content control with a stable tag/title.
57. `word_session_rollback` — Write. Roll back a previous word_session_apply_patchset command by applying its generated rollback PatchSet to the open Word document.
58. `word_session_command_status` — Read-only. Return the current status/result/error for a queued Office.js session command.
59. `officecli_view_html` — Optional OfficeCLI auxiliary read. Render a DOCX to a bounded HTML snapshot for visual evidence. Does not modify the DOCX.
60. `officecli_view_screenshot` — Optional OfficeCLI auxiliary sidecar export. Render a DOCX screenshot/PNG to output_path. Does not modify the DOCX.
61. `officecli_view_issues` — Optional OfficeCLI auxiliary read. Run view issues with JSON output for formatting/content/structure evidence.
62. `officecli_query` — Optional OfficeCLI auxiliary read. Run CSS-like query with --json. Use only for inspection, never mutation.
63. `officecli_validate` — Optional OfficeCLI auxiliary read. Validate the DOCX with OfficeCLI --json as extra evidence after Word AI validation.
