from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from pathlib import Path
from typing import Any, Callable

from .ooxml import (
    apply_patchset,
    assess_patchset,
    backup_docx,
    build_document_map,
    compare_structure,
    diff_text,
    dry_run_patchset,
    export_plain_text,
    export_table_csv,
    extract_comments,
    extract_plain_text,
    get_content_control_text,
    get_outline,
    health_check,
    inspect_docx,
    list_anchors,
    list_bookmarks,
    list_comments,
    list_content_controls,
    list_docx_parts,
    list_fields,
    list_headings,
    list_headers_footers,
    list_hyperlinks,
    list_images,
    list_notes,
    list_numbering,
    list_paragraphs,
    list_revisions,
    list_sections,
    list_styles,
    list_tables,
    package_manifest,
    part_hashes,
    plan_patchset,
    read_anchor,
    read_heading_section,
    read_paragraph,
    read_table,
    read_table_cell,
    restore_backup,
    rollback_docx,
    search_text,
    structural_fingerprint,
    table_to_csv,
    validate_structure,
    write_sidecar_index,
)

JSON = dict[str, Any]


def _text_payload(obj: Any) -> list[dict[str, str]]:
    text = obj if isinstance(obj, str) else json.dumps(obj, ensure_ascii=False, indent=2)
    return [{"type": "text", "text": text}]


def _schema(props: dict[str, Any], required: list[str] | None = None) -> JSON:
    return {"type": "object", "properties": props, "required": required or [], "additionalProperties": False}


class WordAiMcpServer:
    def __init__(self, root: str | None = None, allow_write: bool = True):
        self.root = Path(root or os.getcwd()).resolve()
        self.allow_write = allow_write
        self.tools: dict[str, Callable[[JSON], Any]] = {
            # Package / health / map
            "docx_health_check": self.tool_docx_health_check,
            "docx_inspect": self.tool_docx_inspect,
            "docx_package_manifest": self.tool_docx_package_manifest,
            "docx_list_parts": self.tool_docx_list_parts,
            "docx_part_hashes": self.tool_docx_part_hashes,
            "docx_structural_fingerprint": self.tool_docx_structural_fingerprint,
            "docx_map": self.tool_docx_map,
            # Navigation / retrieval
            "docx_get_outline": self.tool_docx_get_outline,
            "docx_outline": self.tool_docx_get_outline,
            "docx_list_headings": self.tool_docx_list_headings,
            "docx_read_heading_section": self.tool_docx_read_heading_section,
            "docx_list_anchors": self.tool_docx_list_anchors,
            "docx_read_anchor": self.tool_docx_read_anchor,
            "docx_search_text": self.tool_docx_search_text,
            "docx_list_paragraphs": self.tool_docx_list_paragraphs,
            "docx_read_paragraph": self.tool_docx_read_paragraph,
            "docx_list_content_controls": self.tool_docx_list_content_controls,
            "docx_read_content_control": self.tool_docx_read_content_control,
            "docx_extract_plain_text": self.tool_docx_extract_plain_text,
            "docx_export_plain_text": self.tool_docx_export_plain_text,
            # Tables / styles / complex Word objects
            "docx_list_tables": self.tool_docx_list_tables,
            "docx_read_table": self.tool_docx_read_table,
            "docx_read_table_cell": self.tool_docx_read_table_cell,
            "docx_export_table_csv": self.tool_docx_export_table_csv,
            "docx_table_to_csv": self.tool_docx_table_to_csv,
            "docx_list_styles": self.tool_docx_list_styles,
            "docx_list_numbering": self.tool_docx_list_numbering,
            "docx_list_sections": self.tool_docx_list_sections,
            "docx_list_fields": self.tool_docx_list_fields,
            "docx_list_images": self.tool_docx_list_images,
            "docx_list_hyperlinks": self.tool_docx_list_hyperlinks,
            "docx_list_comments": self.tool_docx_list_comments,
            "docx_extract_comments": self.tool_docx_extract_comments,
            "docx_list_bookmarks": self.tool_docx_list_bookmarks,
            "docx_list_headers_footers": self.tool_docx_list_headers_footers,
            "docx_list_notes": self.tool_docx_list_notes,
            "docx_list_revisions": self.tool_docx_list_revisions,
            # Patch / write / validation lifecycle
            "docx_assess_patchset": self.tool_docx_assess_patchset,
            "docx_plan_patchset": self.tool_docx_plan_patchset,
            "docx_preflight_patchset": self.tool_docx_preflight_patchset,
            "docx_dry_run_patchset": self.tool_docx_dry_run_patchset,
            "docx_write_index": self.tool_docx_write_index,
            "docx_backup": self.tool_docx_backup,
            "docx_restore_backup": self.tool_docx_restore_backup,
            "docx_rollback": self.tool_docx_rollback,
            "docx_apply_patchset": self.tool_docx_apply_patchset,
            "docx_validate": self.tool_docx_validate,
            "docx_compare_structure": self.tool_docx_compare_structure,
            "docx_text_diff": self.tool_docx_text_diff,
        }

    def _resolve_path(self, p: str) -> str:
        path = Path(p)
        if not path.is_absolute():
            path = self.root / path
        path = path.resolve()
        try:
            path.relative_to(self.root)
        except ValueError:
            raise PermissionError(f"Path is outside allowed root: {path}")
        return str(path)

    def _maybe_resolve(self, p: str | None) -> str | None:
        return self._resolve_path(p) if p else None

    def _ensure_write(self) -> None:
        if not self.allow_write:
            raise PermissionError("Server is running in read-only mode")

    def list_tools(self) -> list[JSON]:
        strp = {"type": "string"}
        intp = {"type": "integer", "minimum": 1}
        boolp = {"type": "boolean"}
        patchset = {"$ref": "https://word-ai.local/schemas/patchset.schema.json"}
        validate_props = {
            "source_docx": strp,
            "target_docx": strp,
            "strict": {"type": "boolean", "default": True},
            "touched_content_control_tags": {"type": "array", "items": strp},
            "touched_para_ids": {"type": "array", "items": strp},
            "touched_paragraph_indices": {"type": "array", "items": {"type": "integer"}},
            "touched_table_indices": {"type": "array", "items": {"type": "integer"}},
            "touched_table_cells": {"type": "array", "items": {"oneOf": [{"type": "string"}, {"type": "array", "items": {"type": "integer"}, "minItems": 3, "maxItems": 3}, {"type": "object", "properties": {"table_index": {"type": "integer"}, "row": {"type": "integer"}, "col": {"type": "integer"}}, "required": ["table_index", "row", "col"], "additionalProperties": False}]}},
            "allow_paragraph_count_change": {"type": "boolean", "default": False},
            "allow_table_dimension_change": {"type": "boolean", "default": False},
            "allowed_added_content_control_tags": {"type": "array", "items": strp},
        }
        specs = [
            ("docx_health_check", "Read-only. Conservative stability report: duplicate tags, complex content controls, fields, tracked changes, comments, table previews and recommended safety policy.", {"docx_path": strp, "max_items": {"type": "integer", "default": 50}}, ["docx_path"]),
            ("docx_inspect", "Read-only. Inspect counts, anchors and optional visible text.", {"docx_path": strp, "include_text": {"type": "boolean", "default": False}}, ["docx_path"]),
            ("docx_package_manifest", "Read-only. DOCX ZIP manifest with part sizes and optional SHA-256 hashes.", {"docx_path": strp, "include_hashes": {"type": "boolean", "default": True}}, ["docx_path"]),
            ("docx_list_parts", "Read-only. List all package parts with sizes and hashes.", {"docx_path": strp}, ["docx_path"]),
            ("docx_part_hashes", "Read-only. Compact DOCX part-hash map for before/after invariance checks.", {"docx_path": strp}, ["docx_path"]),
            ("docx_structural_fingerprint", "Read-only. Strong structure fingerprint: package parts, content-control hashes, paragraph hashes, table hashes, outline and health report.", {"docx_path": strp}, ["docx_path"]),
            ("docx_map", "Read-only. Codex-friendly map: health, headings, anchors, paragraph chunks, table summaries and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 500}, "include_text": {"type": "boolean", "default": False}}, ["docx_path"]),
            ("docx_get_outline", "Read-only. Heading outline with section paragraph ranges for large-document chunking.", {"docx_path": strp}, ["docx_path"]),
            ("docx_outline", "Read-only alias for docx_get_outline; kept for Codex prompt ergonomics.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_headings", "Read-only. Heading-only view with previews and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_read_heading_section", "Read-only. Read one heading section by heading anchor_id or exact heading text.", {"docx_path": strp, "heading_anchor_id": strp, "heading_text": strp, "max_chars": {"type": "integer", "default": 20000}}, ["docx_path"]),
            ("docx_list_anchors", "Read-only. List stable editing anchors: content controls, headings, bookmarks and paraId paragraphs.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_read_anchor", "Read-only. Read text behind an anchor_id from docx_list_anchors.", {"docx_path": strp, "anchor_id": strp, "max_chars": {"type": "integer", "default": 20000}}, ["docx_path", "anchor_id"]),
            ("docx_search_text", "Read-only. Search visible paragraph text; returns paraId, paragraph index, scope and hashes.", {"docx_path": strp, "query": strp, "regex": {"type": "boolean", "default": False}, "case_sensitive": {"type": "boolean", "default": False}, "max_results": {"type": "integer", "default": 50}, "context_chars": {"type": "integer", "default": 120}}, ["docx_path", "query"]),
            ("docx_list_paragraphs", "Read-only. Paragraph inventory with style, paraId, heading level, scope, preview, hash and complexity.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}, "include_empty": {"type": "boolean", "default": False}}, ["docx_path"]),
            ("docx_read_paragraph", "Read-only. Read one paragraph by paragraph_index or paraId and return hash/complexity.", {"docx_path": strp, "paragraph_index": intp, "paraId": strp}, ["docx_path"]),
            ("docx_list_content_controls", "Read-only. List content controls with tag, alias, id, preview, text hash and complexity.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_read_content_control", "Read-only. Read content-control text by w:tag. Preferred primitive for safe edits.", {"docx_path": strp, "tag": strp}, ["docx_path", "tag"]),
            ("docx_extract_plain_text", "Read-only. Extract visible body text for coarse review; never use to regenerate DOCX.", {"docx_path": strp}, ["docx_path"]),
            ("docx_export_plain_text", "Write-sidecar only. Export visible body text to .txt; does not modify DOCX.", {"docx_path": strp, "out_path": strp}, ["docx_path"]),
            ("docx_list_tables", "Read-only. List table dimensions, scope tags, complexity and preview rows.", {"docx_path": strp, "max_cell_chars": {"type": "integer", "default": 120}}, ["docx_path"]),
            ("docx_read_table", "Read-only. Read a table matrix by table_index and optional content-control scope_tag.", {"docx_path": strp, "table_index": intp, "scope_tag": strp, "max_chars_per_cell": {"type": "integer", "default": 2000}}, ["docx_path", "table_index"]),
            ("docx_read_table_cell", "Read-only. Read a single table cell and return hash for write preconditions.", {"docx_path": strp, "table_index": intp, "row": intp, "col": intp, "scope_tag": strp, "max_chars": {"type": "integer", "default": 20000}}, ["docx_path", "table_index", "row", "col"]),
            ("docx_export_table_csv", "Write-sidecar only. Export table to CSV; does not modify DOCX. Uses default sidecar naming when out_path is omitted.", {"docx_path": strp, "table_index": intp, "out_path": strp, "scope_tag": strp}, ["docx_path", "table_index"]),
            ("docx_table_to_csv", "Write-sidecar only. Export table to CSV using default naming when out_path is omitted.", {"docx_path": strp, "table_index": intp, "out_path": strp, "scope_tag": strp}, ["docx_path", "table_index"]),
            ("docx_list_styles", "Read-only. List Word styles from styles.xml for style-aware generation.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_numbering", "Read-only. Inspect numbering.xml numId/abstractNum mappings.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_sections", "Read-only. Inspect section properties, margins, page size and header/footer refs.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_fields", "Read-only. List field instructions such as TOC, REF, PAGEREF and SEQ.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_list_images", "Read-only. List drawing/pict objects with relationship IDs and alt metadata.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_hyperlinks", "Read-only. List hyperlinks with relationship targets, preview and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_list_comments", "Read-only. Extract comments.xml previews, author/date/id and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 500}}, ["docx_path"]),
            ("docx_extract_comments", "Read-only. Extract full comment metadata and text.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_bookmarks", "Read-only. List bookmarks and paths for cross-reference preservation.", {"docx_path": strp}, ["docx_path"]),
            ("docx_list_headers_footers", "Read-only. Inspect header/footer parts and visible text without modifying them.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 500}}, ["docx_path"]),
            ("docx_list_notes", "Read-only. Inspect footnotes/endnotes when present.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 500}}, ["docx_path"]),
            ("docx_list_revisions", "Read-only. List tracked-change nodes with previews and hashes.", {"docx_path": strp, "max_preview": {"type": "integer", "default": 240}}, ["docx_path"]),
            ("docx_assess_patchset", "Read-only. Resolve PatchSet targets and report risks, touched objects and precondition gaps.", {"docx_path": strp, "patchset": patchset}, ["docx_path", "patchset"]),
            ("docx_plan_patchset", "Read-only. Alias-grade PatchSet planning assessment.", {"docx_path": strp, "patchset": patchset}, ["docx_path", "patchset"]),
            ("docx_preflight_patchset", "Write-temp. Dry-run PatchSet to a temporary copy, validate, and remove output unless keep_output=true.", {"docx_path": strp, "patchset": patchset, "keep_output": {"type": "boolean", "default": False}}, ["docx_path", "patchset"]),
            ("docx_dry_run_patchset", "Write-temp. Same as preflight; useful for explicit Codex dry-run stage.", {"docx_path": strp, "patchset": patchset, "keep_output": {"type": "boolean", "default": False}}, ["docx_path", "patchset"]),
            ("docx_write_index", "Write-sidecar only. Write .wordai index JSON; does not modify DOCX.", {"docx_path": strp, "out_path": strp}, ["docx_path"]),
            ("docx_backup", "Write-sidecar only. Create timestamped backup; does not modify original DOCX.", {"docx_path": strp, "backup_dir": strp}, ["docx_path"]),
            ("docx_restore_backup", "Write. Restore a backup to target_path.", {"backup_path": strp, "target_path": strp}, ["backup_path", "target_path"]),
            ("docx_rollback", "Write. Restore from backup and optionally back up replaced file first.", {"backup_path": strp, "restore_path": strp, "make_backup_of_current": {"type": "boolean", "default": True}}, ["backup_path", "restore_path"]),
            ("docx_apply_patchset", "Write. Apply constrained PatchSet to a new DOCX plus audit JSON; validation gates final commit.", {"docx_path": strp, "output_path": strp, "patchset": patchset}, ["docx_path", "patchset"]),
            ("docx_validate", "Read-only. Validate structural invariants. Supply touched_* for intentional edits.", validate_props, ["source_docx", "target_docx"]),
            ("docx_compare_structure", "Read-only. Same validation report phrased as structural comparison.", validate_props, ["source_docx", "target_docx"]),
            ("docx_text_diff", "Read-only. Unified visible-text diff for human review after validation.", {"source_docx": strp, "target_docx": strp, "context": {"type": "integer", "default": 2}}, ["source_docx", "target_docx"]),
        ]
        return [{"name": n, "description": d, "inputSchema": _schema(p, r)} for n, d, p, r in specs]

    # Package / health / map.
    def tool_docx_health_check(self, args: JSON) -> Any:
        return health_check(self._resolve_path(args["docx_path"]), int(args.get("max_items", 50)))

    def tool_docx_inspect(self, args: JSON) -> Any:
        return inspect_docx(self._resolve_path(args["docx_path"]), bool(args.get("include_text", False)))

    def tool_docx_package_manifest(self, args: JSON) -> Any:
        return package_manifest(self._resolve_path(args["docx_path"]), bool(args.get("include_hashes", True)))

    def tool_docx_list_parts(self, args: JSON) -> Any:
        return list_docx_parts(self._resolve_path(args["docx_path"]))

    def tool_docx_part_hashes(self, args: JSON) -> Any:
        return part_hashes(self._resolve_path(args["docx_path"]))

    def tool_docx_structural_fingerprint(self, args: JSON) -> Any:
        return structural_fingerprint(self._resolve_path(args["docx_path"]))

    def tool_docx_map(self, args: JSON) -> Any:
        return build_document_map(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 500)), bool(args.get("include_text", False)))

    # Navigation / retrieval.
    def tool_docx_get_outline(self, args: JSON) -> Any:
        return get_outline(self._resolve_path(args["docx_path"]))

    def tool_docx_list_headings(self, args: JSON) -> Any:
        return list_headings(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    def tool_docx_read_heading_section(self, args: JSON) -> Any:
        return read_heading_section(self._resolve_path(args["docx_path"]), args.get("heading_anchor_id"), args.get("heading_text"), int(args.get("max_chars", 20000)))

    def tool_docx_list_anchors(self, args: JSON) -> Any:
        anchors = list_anchors(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))
        return {"anchors": [a.to_dict() for a in anchors], "anchor_count": len(anchors)}

    def tool_docx_read_anchor(self, args: JSON) -> Any:
        return read_anchor(self._resolve_path(args["docx_path"]), args["anchor_id"], int(args.get("max_chars", 20000)))

    def tool_docx_search_text(self, args: JSON) -> Any:
        return search_text(self._resolve_path(args["docx_path"]), args["query"], bool(args.get("regex", False)), bool(args.get("case_sensitive", False)), int(args.get("max_results", 50)), int(args.get("context_chars", 120)))

    def tool_docx_list_paragraphs(self, args: JSON) -> Any:
        return list_paragraphs(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)), bool(args.get("include_empty", False)))

    def tool_docx_read_paragraph(self, args: JSON) -> Any:
        return read_paragraph(self._resolve_path(args["docx_path"]), args.get("paragraph_index"), args.get("paraId"))

    def tool_docx_list_content_controls(self, args: JSON) -> Any:
        return list_content_controls(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    def tool_docx_read_content_control(self, args: JSON) -> Any:
        return get_content_control_text(self._resolve_path(args["docx_path"]), args["tag"])

    def tool_docx_extract_plain_text(self, args: JSON) -> Any:
        return extract_plain_text(self._resolve_path(args["docx_path"]))

    def tool_docx_export_plain_text(self, args: JSON) -> Any:
        self._ensure_write()
        return export_plain_text(self._resolve_path(args["docx_path"]), self._maybe_resolve(args.get("out_path")))

    # Tables / styles / complex Word objects.
    def tool_docx_list_tables(self, args: JSON) -> Any:
        return list_tables(self._resolve_path(args["docx_path"]), int(args.get("max_cell_chars", 120)))

    def tool_docx_read_table(self, args: JSON) -> Any:
        data = read_table(self._resolve_path(args["docx_path"]), int(args["table_index"]), args.get("scope_tag"), int(args.get("max_chars_per_cell", 2000)))
        if isinstance(data, dict) and "rows" in data:
            data.setdefault("row_count", len(data["rows"]))
            data.setdefault("column_counts", [len(r) for r in data["rows"]])
        return data

    def tool_docx_read_table_cell(self, args: JSON) -> Any:
        return read_table_cell(self._resolve_path(args["docx_path"]), int(args["table_index"]), int(args["row"]), int(args["col"]), args.get("scope_tag"), int(args.get("max_chars", 20000)))

    def tool_docx_export_table_csv(self, args: JSON) -> Any:
        self._ensure_write()
        return table_to_csv(self._resolve_path(args["docx_path"]), int(args["table_index"]), self._maybe_resolve(args.get("out_path")), args.get("scope_tag"))

    def tool_docx_table_to_csv(self, args: JSON) -> Any:
        self._ensure_write()
        return table_to_csv(self._resolve_path(args["docx_path"]), int(args["table_index"]), self._maybe_resolve(args.get("out_path")), args.get("scope_tag"))

    def tool_docx_list_styles(self, args: JSON) -> Any:
        return list_styles(self._resolve_path(args["docx_path"]))

    def tool_docx_list_numbering(self, args: JSON) -> Any:
        return list_numbering(self._resolve_path(args["docx_path"]))

    def tool_docx_list_sections(self, args: JSON) -> Any:
        return list_sections(self._resolve_path(args["docx_path"]))

    def tool_docx_list_fields(self, args: JSON) -> Any:
        return list_fields(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    def tool_docx_list_images(self, args: JSON) -> Any:
        return list_images(self._resolve_path(args["docx_path"]))

    def tool_docx_list_hyperlinks(self, args: JSON) -> Any:
        return list_hyperlinks(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    def tool_docx_list_comments(self, args: JSON) -> Any:
        return list_comments(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 500)))

    def tool_docx_extract_comments(self, args: JSON) -> Any:
        return extract_comments(self._resolve_path(args["docx_path"]))

    def tool_docx_list_bookmarks(self, args: JSON) -> Any:
        return list_bookmarks(self._resolve_path(args["docx_path"]))

    def tool_docx_list_headers_footers(self, args: JSON) -> Any:
        return list_headers_footers(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 500)))

    def tool_docx_list_notes(self, args: JSON) -> Any:
        return list_notes(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 500)))

    def tool_docx_list_revisions(self, args: JSON) -> Any:
        return list_revisions(self._resolve_path(args["docx_path"]), int(args.get("max_preview", 240)))

    # Patch / write / validation lifecycle.
    def tool_docx_assess_patchset(self, args: JSON) -> Any:
        return assess_patchset(self._resolve_path(args["docx_path"]), args["patchset"])

    def tool_docx_plan_patchset(self, args: JSON) -> Any:
        return plan_patchset(self._resolve_path(args["docx_path"]), args["patchset"])

    def tool_docx_preflight_patchset(self, args: JSON) -> Any:
        self._ensure_write()
        return dry_run_patchset(self._resolve_path(args["docx_path"]), args["patchset"], bool(args.get("keep_output", False)))

    def tool_docx_dry_run_patchset(self, args: JSON) -> Any:
        self._ensure_write()
        return dry_run_patchset(self._resolve_path(args["docx_path"]), args["patchset"], bool(args.get("keep_output", False)))

    def tool_docx_write_index(self, args: JSON) -> Any:
        self._ensure_write()
        return {"index_path": write_sidecar_index(self._resolve_path(args["docx_path"]), self._maybe_resolve(args.get("out_path")))}

    def tool_docx_backup(self, args: JSON) -> Any:
        self._ensure_write()
        return {"backup_path": backup_docx(self._resolve_path(args["docx_path"]), self._maybe_resolve(args.get("backup_dir")))}

    def tool_docx_restore_backup(self, args: JSON) -> Any:
        self._ensure_write()
        return restore_backup(self._resolve_path(args["backup_path"]), self._resolve_path(args["target_path"]))

    def tool_docx_rollback(self, args: JSON) -> Any:
        self._ensure_write()
        return rollback_docx(self._resolve_path(args["backup_path"]), self._resolve_path(args["restore_path"]), bool(args.get("make_backup_of_current", True)))

    def tool_docx_apply_patchset(self, args: JSON) -> Any:
        self._ensure_write()
        return apply_patchset(self._resolve_path(args["docx_path"]), args["patchset"], self._maybe_resolve(args.get("output_path")))

    def _validation_kwargs(self, args: JSON, target_docx: str | None = None) -> dict[str, Any]:
        """Build validation kwargs and, when possible, auto-load touched scopes
        from the adjacent audit JSON produced by docx_apply_patchset. This keeps
        Codex validation ergonomic while preserving strict structure checks.
        Explicit arguments always override audit-derived values.
        """
        touched_tags = args.get("touched_content_control_tags") or []
        touched_para_ids = args.get("touched_para_ids") or []
        touched_paragraph_indices = args.get("touched_paragraph_indices") or []
        touched_table_indices = args.get("touched_table_indices") or []
        touched_table_cells = args.get("touched_table_cells") or []
        allowed_added_tags = args.get("allowed_added_content_control_tags") or []

        if target_docx and not (touched_tags or touched_para_ids or touched_paragraph_indices or touched_table_indices or touched_table_cells or allowed_added_tags) and args.get("use_audit", True):
            audit_path = Path(target_docx).with_suffix(".audit.json")
            if audit_path.exists():
                try:
                    audit = json.loads(audit_path.read_text(encoding="utf-8"))
                    touched = ((audit.get("safety_assessment") or {}).get("touched") or {})
                    touched_tags = touched.get("content_control_tags") or []
                    touched_para_ids = touched.get("paraIds") or []
                    touched_paragraph_indices = touched.get("paragraph_indices") or []
                    touched_table_indices = touched.get("table_indices") or []
                    touched_table_cells = touched.get("table_cells") or []
                    for op in audit.get("applied") or []:
                        if op.get("paragraph_index") is not None:
                            touched_paragraph_indices.append(op["paragraph_index"])
                        if op.get("anchor_paragraph_index") is not None:
                            touched_paragraph_indices.append(op["anchor_paragraph_index"])
                        if op.get("target_paragraph_index") is not None:
                            touched_paragraph_indices.append(op["target_paragraph_index"])
                        if op.get("table_index") is not None and op.get("row") is not None and op.get("col") is not None:
                            touched_table_cells.append(f"{op['table_index']}:{op['row']}:{op['col']}")
                        if op.get("op") == "wrap_paragraph_with_content_control" and op.get("tag"):
                            allowed_added_tags.append(op["tag"])
                except Exception:
                    pass

        return {
            "touched_content_control_tags": touched_tags,
            "touched_para_ids": touched_para_ids,
            "touched_paragraph_indices": sorted({int(x) for x in touched_paragraph_indices}),
            "touched_table_indices": touched_table_indices,
            "touched_table_cells": sorted({str(x) for x in touched_table_cells}),
            "allow_paragraph_count_change": bool(args.get("allow_paragraph_count_change", False)),
            "allow_table_dimension_change": bool(args.get("allow_table_dimension_change", False)),
            "allowed_added_content_control_tags": allowed_added_tags,
        }

    def tool_docx_validate(self, args: JSON) -> Any:
        source = self._resolve_path(args["source_docx"])
        target = self._resolve_path(args["target_docx"])
        return validate_structure(source, target, bool(args.get("strict", True)), **self._validation_kwargs(args, target)).to_dict()

    def tool_docx_compare_structure(self, args: JSON) -> Any:
        source = self._resolve_path(args["source_docx"])
        target = self._resolve_path(args["target_docx"])
        return validate_structure(
            source,
            target,
            bool(args.get("strict", True)),
            **self._validation_kwargs(args, target),
        ).to_dict()

    def tool_docx_text_diff(self, args: JSON) -> Any:
        return diff_text(self._resolve_path(args["source_docx"]), self._resolve_path(args["target_docx"]), int(args.get("context", 2)))

    def handle(self, request: JSON) -> JSON | None:
        method = request.get("method")
        req_id = request.get("id")
        try:
            if method == "initialize":
                result = {"protocolVersion": request.get("params", {}).get("protocolVersion", "2025-11-25"), "serverInfo": {"name": "word-ai-mcp", "version": "0.5.0"}, "capabilities": {"tools": {"listChanged": False}, "resources": {}, "prompts": {}}}
                return {"jsonrpc": "2.0", "id": req_id, "result": result}
            if method == "notifications/initialized":
                return None
            if method == "tools/list":
                return {"jsonrpc": "2.0", "id": req_id, "result": {"tools": self.list_tools()}}
            if method == "tools/call":
                params = request.get("params") or {}
                name = params.get("name")
                arguments = params.get("arguments") or {}
                if name not in self.tools:
                    raise ValueError(f"Unknown tool: {name}")
                result = self.tools[name](arguments)
                return {"jsonrpc": "2.0", "id": req_id, "result": {"content": _text_payload(result), "isError": False}}
            if method == "resources/list":
                return {"jsonrpc": "2.0", "id": req_id, "result": {"resources": []}}
            if method == "prompts/list":
                return {"jsonrpc": "2.0", "id": req_id, "result": {"prompts": []}}
            raise ValueError(f"Unsupported method: {method}")
        except Exception as exc:
            return {"jsonrpc": "2.0", "id": req_id, "error": {"code": -32000, "message": str(exc), "data": {"traceback": traceback.format_exc(limit=8)}}}


def run_stdio(root: str | None, allow_write: bool) -> None:
    server = WordAiMcpServer(root=root, allow_write=allow_write)
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
            resp = server.handle(req)
        except Exception as exc:
            resp = {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": str(exc)}}
        if resp is not None:
            sys.stdout.write(json.dumps(resp, ensure_ascii=False) + "\n")
            sys.stdout.flush()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Word AI MCP stdio server")
    parser.add_argument("--root", default=os.getcwd(), help="Allowed workspace root. All DOCX paths must be inside this directory.")
    parser.add_argument("--read-only", action="store_true", help="Disable write tools")
    args = parser.parse_args(argv)
    run_stdio(args.root, allow_write=not args.read_only)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
