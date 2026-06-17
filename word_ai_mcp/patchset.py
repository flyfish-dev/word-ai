from __future__ import annotations

import copy
import re
from typing import Any


JSON = dict[str, Any]

OP_ALIASES = {
    "replace_content_control": "replace_content_control_text",
    "set_content_control": "replace_content_control_text",
    "set_content_control_text": "replace_content_control_text",
    "update_content_control": "replace_content_control_text",
    "update_content_control_text": "replace_content_control_text",
    "append_content_control": "append_content_control_text",
    "prepend_content_control": "prepend_content_control_text",
    "replace_in_content_control": "replace_text_in_content_control",
    "replace_content_control_inner_text": "replace_text_in_content_control",
    "replace_table_cell": "replace_table_cell_text",
    "set_table_cell": "replace_table_cell_text",
    "set_table_cell_text": "replace_table_cell_text",
    "append_row": "append_table_row",
    "replace_paragraph": "replace_paragraph_text",
    "set_paragraph_text": "replace_paragraph_text",
    "insert_after_paragraph": "insert_paragraph_after",
    "insert_before_paragraph": "insert_paragraph_before",
    "wrap_with_content_control": "wrap_paragraph_with_content_control",
    "wrap_content_control": "wrap_paragraph_with_content_control",
    "comment": "add_comment",
    "insert_comment": "add_comment",
}

OP_NAME_ALIASES = ("operation", "operation_type", "operationType", "type", "action", "name")
OPS_ALIASES = ("ops", "edits", "patches")
TEXT_ALIASES = ("new_text", "newText", "replacement_text", "replacementText", "value", "content")
EXPECTED_SHA_ALIASES = ("expected_sha256", "expectedSha256", "old_sha256", "oldSha256", "expected_text_sha256", "expectedTextSha256", "text_sha256", "textSha256")
EXPECTED_TEXT_ALIASES = ("expected_text", "expectedText", "current_text", "currentText")
CONTENT_CONTROL_TEXT_OPS = {
    "replace_content_control_text",
    "append_content_control_text",
    "prepend_content_control_text",
    "replace_text_in_content_control",
}
TEXT_TARGET_OPS = {
    "replace_content_control_text",
    "append_content_control_text",
    "prepend_content_control_text",
    "replace_paragraph_text",
    "insert_paragraph_after",
    "insert_paragraph_before",
    "replace_table_cell_text",
    "add_comment",
}


def normalize_patchset(patchset: JSON) -> JSON:
    """Return a canonical PatchSet copy for agent-generated payloads.

    This is intentionally only an input-compatibility layer. It maps common field
    aliases to the official PatchSet schema but does not introduce any new write
    operation or bypass preconditions, assessment, dry-run, validation, or audit.
    """
    if not isinstance(patchset, dict):
        raise ValueError("patchset must be an object")

    normalized: JSON = copy.deepcopy(patchset)
    consumed: set[str] = set()
    _move_first(normalized, "schema_version", ("schemaVersion", "version"), consumed)
    _move_first(normalized, "source_sha256", ("sourceSha256", "source_hash", "sourceHash", "docx_sha256", "docxSha256", "sha256"), consumed)
    _move_first(normalized, "abort_on_validation_error", ("abortOnValidationError",), consumed)
    _move_first(normalized, "keep_invalid_output", ("keepInvalidOutput",), consumed)
    _delete_consumed(normalized, consumed)

    guard = normalized.get("guard")
    if guard is None and (_has_alias(normalized, "require_preconditions") or _has_alias(normalized, "allow_overwrite")):
        guard = {}
        normalized["guard"] = guard
    if isinstance(guard, dict):
        guard_consumed: set[str] = set()
        _move_first(guard, "require_preconditions", ("requirePreconditions", "preconditions_required", "preconditionsRequired"), guard_consumed)
        _move_first(guard, "allow_overwrite", ("allowOverwrite", "overwrite_allowed", "overwriteAllowed"), guard_consumed)
        _delete_consumed(guard, guard_consumed)
        top_guard_consumed: set[str] = set()
        if guard.get("require_preconditions") is None:
            _move_first_from_to(normalized, guard, "require_preconditions", ("requirePreconditions", "preconditions_required", "preconditionsRequired"), top_guard_consumed)
        if guard.get("allow_overwrite") is None:
            _move_first_from_to(normalized, guard, "allow_overwrite", ("allowOverwrite", "overwrite_allowed", "overwriteAllowed"), top_guard_consumed)
        _delete_consumed(normalized, top_guard_consumed)

    operations = normalized.get("operations")
    if operations is None:
        key = _find_key(normalized, OPS_ALIASES)
        if key:
            operations = normalized[key]
            normalized["operations"] = operations
            if key != "operations":
                del normalized[key]

    if operations is None:
        return normalized
    if not isinstance(operations, list):
        raise ValueError("patchset.operations must be a list")

    normalized["operations"] = [_normalize_operation(op, idx) for idx, op in enumerate(operations)]
    return normalized


def _normalize_operation(operation: Any, index: int) -> JSON:
    if not isinstance(operation, dict):
        raise ValueError(f"patchset.operations[{index}] must be an object")
    op: JSON = copy.deepcopy(operation)
    consumed: set[str] = set()

    raw_op = op.get("op")
    if raw_op is None:
        op_key = _find_key(op, OP_NAME_ALIASES)
        if op_key:
            raw_op = op[op_key]
            consumed.add(op_key)
    canonical_op = _canonical_op(raw_op)
    if not canonical_op:
        keys = ", ".join(sorted(str(key) for key in op.keys())) or "<none>"
        raise ValueError(
            f"patchset.operations[{index}].op is required; accepted aliases are "
            f"{', '.join(OP_NAME_ALIASES)}. Got keys: {keys}"
        )
    op["op"] = canonical_op

    if canonical_op in CONTENT_CONTROL_TEXT_OPS:
        _move_first(op, "tag", ("content_control_tag", "contentControlTag", "target_tag", "targetTag", "anchor_tag", "anchorTag"), consumed)

    if canonical_op in TEXT_TARGET_OPS:
        _move_first(op, "text", TEXT_ALIASES, consumed)
    if canonical_op == "replace_text_in_content_control":
        _move_first(op, "find", ("find_text", "findText", "old_text", "oldText", "search_text", "searchText", "search"), consumed)
        _move_first(op, "replace", ("new_text", "newText", "replacement_text", "replacementText", "replace_with", "replaceWith", "value", "content"), consumed)

    if canonical_op != "replace_text_in_content_control":
        _move_first(op, "expected_old_text", (*EXPECTED_TEXT_ALIASES, "old_text", "oldText"), consumed)
    else:
        _move_first(op, "expected_old_text", EXPECTED_TEXT_ALIASES, consumed)
    _move_first(op, "expected_old_sha256", EXPECTED_SHA_ALIASES, consumed)

    _move_first(op, "paraId", ("para_id", "paraID", "paragraph_id", "paragraphId"), consumed)
    _move_first(op, "paragraph_index", ("paragraphIndex", "paragraph_no", "paragraphNo"), consumed)
    _move_first(op, "content_control_tag", ("contentControlTag",), consumed)
    _move_first(op, "target_tag", ("targetTag",), consumed)
    _move_first(op, "scope_tag", ("scopeTag", "table_scope_tag", "tableScopeTag"), consumed)
    _move_first(op, "table_index", ("tableIndex", "table_no", "tableNo"), consumed)
    _move_first(op, "row", ("row_index", "rowIndex"), consumed)
    _move_first(op, "col", ("column", "column_index", "columnIndex", "col_index", "colIndex"), consumed)
    _move_first(op, "template_row", ("templateRow", "template_row_index", "templateRowIndex"), consumed)
    _move_first(op, "preserve_style", ("preserveStyle",), consumed)
    _move_first(op, "allow_complex_content", ("allowComplexContent",), consumed)
    _move_first(op, "allow_paragraph_count_change", ("allowParagraphCountChange",), consumed)
    _move_first(op, "require_match", ("requireMatch",), consumed)
    _move_first(op, "inherit_style", ("inheritStyle",), consumed)
    _move_first(op, "inherit_heading_style", ("inheritHeadingStyle",), consumed)

    _delete_consumed(op, consumed)
    return {"op": op["op"], **{key: value for key, value in op.items() if key != "op"}}


def _canonical_op(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"(?<=[a-z0-9])([A-Z])", r"_\1", text)
    text = re.sub(r"(?<=[A-Z])([A-Z][a-z])", r"_\1", text)
    text = re.sub(r"[\s\-]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_").lower()
    return OP_ALIASES.get(text, text)


def _token(value: str) -> str:
    text = re.sub(r"(?<=[a-z0-9])([A-Z])", r"_\1", str(value))
    text = re.sub(r"(?<=[A-Z])([A-Z][a-z])", r"_\1", text)
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _find_key(obj: dict[str, Any], aliases: tuple[str, ...]) -> str | None:
    lookup = {_token(str(key)): key for key in obj.keys()}
    for alias in aliases:
        key = lookup.get(_token(alias))
        if key is not None:
            return str(key)
    return None


def _has_alias(obj: dict[str, Any], alias: str) -> bool:
    return _find_key(obj, (alias,)) is not None


def _move_first(obj: dict[str, Any], target: str, aliases: tuple[str, ...], consumed: set[str]) -> None:
    if obj.get(target) is not None:
        return
    key = _find_key(obj, aliases)
    if key is None or key == target or obj.get(key) is None:
        return
    obj[target] = obj[key]
    consumed.add(key)


def _move_first_from_to(source: dict[str, Any], target_obj: dict[str, Any], target: str, aliases: tuple[str, ...], consumed: set[str]) -> None:
    key = _find_key(source, (target, *aliases))
    if key is None or source.get(key) is None:
        return
    target_obj[target] = source[key]
    consumed.add(key)


def _delete_consumed(obj: dict[str, Any], consumed: set[str]) -> None:
    for key in consumed:
        if key in obj:
            del obj[key]
