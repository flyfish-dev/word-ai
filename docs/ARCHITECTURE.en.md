# Word AI Architecture

Word AI is not a document generator. It is a document transaction system for AI-assisted editing of existing Word documents.

## Goals

- Preserve the original DOCX package structure.
- Avoid full-document rebuilds and Markdown/HTML round-trips.
- Keep writes constrained to explicit anchors.
- Make every write auditable, reversible, and structurally validated.
- Give agents rich read tools while forcing formal writes through PatchSet.

## Components

### MCP Client

Codex, OpenAI Agents, or another MCP client reads document structure, identifies target anchors, asks the model to generate replacement text, and submits a constrained PatchSet.

### MCP Server

The Python MCP server exposes 49 tools:

- package inspection and health checks
- heading, paragraph, bookmark, and content-control navigation
- table, style, numbering, field, image, hyperlink, comment, note, and revision inspection
- PatchSet assessment, dry-run, backup, apply, validation, rollback, and diff

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
- apply supported operations to the currently open Word document after SHA-256 precondition checks

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

All `/office/*` POST requests require a local token.

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

