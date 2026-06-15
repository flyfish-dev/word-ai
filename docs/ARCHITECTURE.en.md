# Word AI Architecture

Word AI is not a document generator. It is a document transaction system for AI-assisted editing of existing Word documents.

## Goals

- Preserve the original DOCX package structure.
- Avoid full-document rebuilds and Markdown/HTML round-trips.
- Keep writes constrained to explicit anchors.
- Make every write auditable, reversible, and structurally validated.
- Give agents rich read tools while forcing formal writes through PatchSet.
- Support both offline DOCX file transactions and live Word session edits through Office.js.

## Components

### MCP Client

Codex, OpenAI Agents, or another MCP client reads document structure, identifies target anchors, asks the model to generate replacement text, and submits a constrained PatchSet.

### MCP Server

The Python MCP server exposes 63 tools:

- package inspection and health checks
- heading, paragraph, bookmark, and content-control navigation
- table, style, numbering, field, image, hyperlink, comment, note, and revision inspection
- PatchSet assessment, dry-run, backup, apply, validation, rollback, and diff
- live Word session listing, snapshot, content-control reads, PatchSet preview/apply, selection wrapping, rollback, and command status
- optional OfficeCLI auxiliary evidence: HTML rendering, screenshot sidecar export, issues, query, and validation

All paths are scoped to the configured root directory.

### Python OOXML Engine

The Python engine implements the local-first reference behavior:

- read DOCX package parts
- resolve content controls, paragraphs, and tables
- apply targeted XML edits
- preserve non-target package parts
- validate structural invariants
- write audit JSON

### .NET Open XML SDK Engine

The .NET 8 engine provides a production-oriented Open XML SDK implementation. It supports the same PatchSet model and uses typed Open XML APIs plus OpenXmlValidator checks.

### Office.js Taskpane

The Word add-in is the session layer:

- wrap current selection as a content control
- assign stable tags and titles
- list open-document content controls
- connect to the local Office bridge
- build, assess, dry-run, and apply PatchSets
- register the currently open Word document as a live session
- poll queued Codex commands from the local bridge
- apply supported operations to the currently open Word document after a live preflight and SHA-256 precondition checks
- return an audit object and rollback PatchSet for live session writes

### Office Bridge

The local HTTP bridge is a convenience layer for the taskpane. It does not bypass MCP rules. File-level writes still flow through:

```text
health -> assess -> dry-run -> backup -> apply -> validate -> diff
```

Bridge endpoints:

- `/office/read`
- `/office/build-patchset`
- `/office/assess-patchset`
- `/office/preview-patchset`
- `/office/apply-patchset`
- `/office/session/register`
- `/office/session/heartbeat`
- `/office/session/poll`
- `/office/session/result`
- `/office/session/list`

All `/office/*` POST requests require a local token.

Live Word session flow:

```text
Codex MCP client
  -> word_session_* tool
  -> .wordai/sessions/commands/*.json
  -> Office.js taskpane polling loop
  -> Word.run(...)
  -> /office/session/result
  -> Codex receives audit / rollback / error
```

## PatchSet Lifecycle

```text
Source DOCX
  -> inspect and read target anchor
  -> model generates PatchSet
  -> assess risks and preconditions
  -> dry-run to temporary candidate
  -> validate structure
  -> backup source
  -> apply to a new output DOCX
  -> write audit JSON
  -> compare structure and text diff
```

Live Word session writes use the same PatchSet shape but do not create a new DOCX file. They target the document currently open in Word and are limited to content-control text operations. `word_session_apply_patchset` performs a live preflight first, re-checks `expected_old_sha256`, applies through Office.js, and returns audit plus rollback PatchSet.

## Safety Boundaries

Word AI blocks or reports risk for:

- missing source hash when preconditions are required
- missing target old-text hash for high-risk operations
- complex content controls with tables, fields, comments, images, or tracked changes
- unauthorized package part changes
- unexpected content-control/table/paragraph drift outside touched targets
- tracked changes, field, image, comment, and table count drift unless explicitly authorized

## Production Recommendations

- Use the .NET Open XML SDK engine as the authoritative batch editing engine.
- Store document indexes in SQLite, PostgreSQL, or another durable store for large document sets.
- Use render/PDF diff for release-grade validation.
- Use Word COM, Aspose, or Syncfusion for field refresh and PDF export where Office.js APIs are insufficient.
- Put write tools behind human approval, RBAC, audit logging, and network controls.
