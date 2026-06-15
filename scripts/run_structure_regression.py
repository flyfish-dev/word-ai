from __future__ import annotations

import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from word_ai_mcp.ooxml import (
    apply_patchset,
    list_comments,
    list_content_controls,
    list_tables,
    read_paragraph,
    read_table_cell,
)


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    src = root / "examples" / "sample_contract.docx"
    ops_tested: list[str] = []

    with TemporaryDirectory() as td0:
        td = Path(td0)

        old = read_table_cell(src, 1, 3, 3)
        patch = {
            "schema_version": "2.0",
            "strict": True,
            "source_sha256": sha256_file(src),
            "guard": {"require_preconditions": True},
            "reason": "regression: replace one table cell only",
            "operations": [
                {
                    "op": "replace_table_cell_text",
                    "table_index": 1,
                    "row": 3,
                    "col": 3,
                    "expected_old_sha256": old["text_sha256"],
                    "text": "仅修改指定锚点范围，并对每次写入生成可验证审计记录。",
                }
            ],
        }
        out = td / "cell.docx"
        audit = apply_patchset(src, patch, out)
        assert audit["validation"]["ok"], audit["validation"]
        assert read_table_cell(out, 1, 3, 3)["text"].startswith("仅修改指定")
        ops_tested.append("replace_table_cell_text")

        oldp = read_paragraph(src, paragraph_index=6)
        patch = {
            "schema_version": "2.0",
            "strict": True,
            "source_sha256": sha256_file(src),
            "guard": {"require_preconditions": True},
            "reason": "regression: replace one paragraph with precondition",
            "operations": [
                {
                    "op": "replace_paragraph_text",
                    "paragraph_index": 6,
                    "expected_old_sha256": oldp["text_sha256"],
                    "text": "本项目面向多部门协同场景，要求文档编辑动作可定位、可审计、可回滚。",
                }
            ],
        }
        out = td / "paragraph.docx"
        audit = apply_patchset(src, patch, out)
        assert audit["validation"]["ok"], audit["validation"]
        assert read_paragraph(out, paragraph_index=6)["text"].startswith("本项目面向多部门")
        ops_tested.append("replace_paragraph_text")

        patch = {
            "schema_version": "2.0",
            "strict": True,
            "source_sha256": sha256_file(src),
            "guard": {"require_preconditions": True},
            "reason": "regression: append one table row with permission",
            "operations": [
                {
                    "op": "append_table_row",
                    "table_index": 1,
                    "template_row": 4,
                    "expected_old_sha256": list_tables(src)["tables"][0]["text_sha256"],
                    "values": ["FR-004", "结构校验", "验证非目标内容、表格、图片、字段和批注未发生未授权变化。"],
                }
            ],
        }
        out = td / "row.docx"
        audit = apply_patchset(src, patch, out)
        assert audit["validation"]["ok"], audit["validation"]
        assert list_tables(out)["tables"][0]["row_count"] == list_tables(src)["tables"][0]["row_count"] + 1
        ops_tested.append("append_table_row")

        patch = {
            "schema_version": "2.0",
            "strict": True,
            "source_sha256": sha256_file(src),
            "guard": {"require_preconditions": True},
            "reason": "regression: add one comment with controlled relationship changes",
            "operations": [
                {
                    "op": "add_comment",
                    "paragraph_index": 6,
                    "expected_old_sha256": read_paragraph(src, paragraph_index=6)["text_sha256"],
                    "text": "请确认项目背景是否需要补充实施范围。",
                    "author": "Word AI",
                    "initials": "AI",
                }
            ],
        }
        out = td / "comment.docx"
        audit = apply_patchset(src, patch, out)
        assert audit["validation"]["ok"], audit["validation"]
        assert list_comments(out)["comment_count"] == 1
        ops_tested.append("add_comment")

        oldp = read_paragraph(src, paragraph_index=6)
        patch = {
            "schema_version": "2.0",
            "strict": True,
            "source_sha256": sha256_file(src),
            "guard": {"require_preconditions": True},
            "reason": "regression: insert paragraph without modifying anchor paragraph",
            "operations": [
                {
                    "op": "insert_paragraph_after",
                    "paragraph_index": 6,
                    "expected_old_sha256": oldp["text_sha256"],
                    "text": "新增说明：该项目背景段落的补充内容以独立段落插入，不修改原段落文本。",
                }
            ],
        }
        out = td / "insert.docx"
        audit = apply_patchset(src, patch, out)
        assert audit["validation"]["ok"], audit["validation"]
        ops_tested.append("insert_paragraph_after")

        patch = {
            "schema_version": "2.0",
            "strict": True,
            "source_sha256": sha256_file(src),
            "guard": {"require_preconditions": True},
            "reason": "regression: wrap paragraph with new content control tag",
            "operations": [
                {
                    "op": "wrap_paragraph_with_content_control",
                    "paragraph_index": 6,
                    "tag": "WORD-AI:REGRESSION:background",
                    "alias": "项目背景正文锚点",
                    "lock": True,
                    "expected_old_sha256": read_paragraph(src, paragraph_index=6)["text_sha256"],
                }
            ],
        }
        out = td / "wrap.docx"
        audit = apply_patchset(src, patch, out)
        assert audit["validation"]["ok"], audit["validation"]
        tags = [cc["tag"] for cc in list_content_controls(out)["content_controls"]]
        assert "WORD-AI:REGRESSION:background" in tags
        ops_tested.append("wrap_paragraph_with_content_control")

    print(json.dumps({"ok": True, "ops_tested": ops_tested}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
