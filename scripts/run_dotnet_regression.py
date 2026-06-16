from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory

from word_ai_mcp.ooxml import (
    list_comments,
    list_content_controls,
    list_tables,
    read_paragraph,
    read_table_cell,
)
from scripts.run_outline_regression import make_fixture


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "examples" / "sample_contract.docx"
PROJECT = ROOT / "dotnet" / "WordAi.OpenXml" / "WordAi.OpenXml.csproj"
DOTNET_ROOT = Path("/opt/homebrew/opt/dotnet@8/libexec")


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def dotnet_cmd() -> list[str]:
    dotnet = DOTNET_ROOT / "dotnet"
    if dotnet.exists():
        return [str(dotnet)]
    found = shutil.which("dotnet")
    if not found:
        raise RuntimeError("dotnet not found")
    return [found]


def run_dotnet(*args: str, input_path: Path | None = None) -> dict:
    env = os.environ.copy()
    if DOTNET_ROOT.exists():
        env["DOTNET_ROOT"] = str(DOTNET_ROOT)
    cmd = dotnet_cmd() + ["run", "--project", str(PROJECT), "-c", "Release", "--", *args]
    result = subprocess.run(cmd, cwd=ROOT, env=env, text=True, capture_output=True, timeout=60)
    if result.returncode != 0:
        raise RuntimeError(result.stderr or result.stdout)
    if not result.stdout.strip():
        return {}
    return json.loads(result.stdout)


def apply_patch(src: Path, patch: dict, out: Path) -> dict:
    patch_path = out.with_suffix(".patch.json")
    patch_path.write_text(json.dumps(patch, ensure_ascii=False, indent=2), encoding="utf-8")
    return run_dotnet("apply", str(src), str(patch_path), str(out))


def guarded_patch(reason: str, operations: list[dict]) -> dict:
    return {
        "schema_version": "2.0",
        "strict": True,
        "source_sha256": sha256_file(SRC),
        "guard": {"require_preconditions": True},
        "reason": reason,
        "operations": operations,
    }


def main() -> int:
    ops_tested: list[str] = []
    run_dotnet("inspect", str(SRC))

    with TemporaryDirectory() as td0:
        td = Path(td0)

        outline_fixture = td / "outline-regression.docx"
        make_fixture(outline_fixture)
        outline_profile = run_dotnet("inspect", str(outline_fixture))
        outline_anchors = [a for a in outline_profile["anchors"] if a["kind"] == "heading"]
        assert outline_profile["heading_count"] == 6, outline_profile
        assert [a["text_preview"] for a in outline_anchors] == ["功能介绍", "目录管理系统", "中文一级标题", "中文二级标题", "直接大纲级别", "附录一"], outline_anchors
        ops_tested.append("inspect_chinese_outline_without_toc")

        cc = list_content_controls(SRC)["content_controls"][0]
        out = td / "replace-control.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: replace content control",
                [
                    {
                        "op": "replace_content_control_text",
                        "tag": cc["tag"],
                        "expected_old_sha256": cc["text_sha256"],
                        "text": "本文档描述系统建设目标、范围、角色和总体约束，并由 .NET OpenXML 引擎安全替换。",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        ops_tested.append("replace_content_control_text")

        out = td / "replace-phrase.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: replace phrase in content control",
                [
                    {
                        "op": "replace_text_in_content_control",
                        "tag": cc["tag"],
                        "expected_old_sha256": cc["text_sha256"],
                        "find": "系统建设目标",
                        "replace": "系统建设目标与治理边界",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        ops_tested.append("replace_text_in_content_control")

        out = td / "append-control.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: append content control text",
                [
                    {
                        "op": "append_content_control_text",
                        "tag": cc["tag"],
                        "expected_old_sha256": cc["text_sha256"],
                        "text": ".NET 引擎追加的独立段落。",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        ops_tested.append("append_content_control_text")

        out = td / "prepend-control.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: prepend content control text",
                [
                    {
                        "op": "prepend_content_control_text",
                        "tag": cc["tag"],
                        "expected_old_sha256": cc["text_sha256"],
                        "text": ".NET 引擎前置的独立段落。",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        ops_tested.append("prepend_content_control_text")

        old = read_table_cell(SRC, 1, 3, 3)
        out = td / "cell.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: replace table cell",
                [
                    {
                        "op": "replace_table_cell_text",
                        "table_index": 1,
                        "row": 3,
                        "col": 3,
                        "expected_old_sha256": old["text_sha256"],
                        "text": "仅修改指定锚点范围，并由 .NET OpenXML 引擎验证。",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        assert read_table_cell(out, 1, 3, 3)["text"].startswith("仅修改指定")
        ops_tested.append("replace_table_cell_text")

        oldp = read_paragraph(SRC, paragraph_index=6)
        out = td / "paragraph.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: replace paragraph",
                [
                    {
                        "op": "replace_paragraph_text",
                        "paragraph_index": 6,
                        "expected_old_sha256": oldp["text_sha256"],
                        "text": "本项目面向多部门协同场景，要求 .NET 引擎同样可定位、可审计、可回滚。",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        ops_tested.append("replace_paragraph_text")

        out = td / "row.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: append table row",
                [
                    {
                        "op": "append_table_row",
                        "table_index": 1,
                        "template_row": 4,
                        "expected_old_sha256": list_tables(SRC)["tables"][0]["text_sha256"],
                        "values": ["FR-004", ".NET 结构校验", "验证非目标内容和表格单元格未发生未授权变化。"],
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        assert list_tables(out)["tables"][0]["row_count"] == list_tables(SRC)["tables"][0]["row_count"] + 1
        ops_tested.append("append_table_row")

        out = td / "comment.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: add comment",
                [
                    {
                        "op": "add_comment",
                        "paragraph_index": 6,
                        "expected_old_sha256": oldp["text_sha256"],
                        "text": "请确认项目背景是否需要补充实施范围。",
                        "author": "Word AI",
                        "initials": "AI",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        assert list_comments(out)["comment_count"] == 1
        ops_tested.append("add_comment")

        out = td / "insert.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: insert paragraph",
                [
                    {
                        "op": "insert_paragraph_after",
                        "paragraph_index": 6,
                        "expected_old_sha256": oldp["text_sha256"],
                        "text": "新增说明：该段落由 .NET 引擎插入，并保留原段落。",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        ops_tested.append("insert_paragraph_after")

        out = td / "insert-before.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: insert paragraph before",
                [
                    {
                        "op": "insert_paragraph_before",
                        "paragraph_index": 6,
                        "expected_old_sha256": oldp["text_sha256"],
                        "text": "前置说明：该段落由 .NET 引擎插入，并保留原段落。",
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        ops_tested.append("insert_paragraph_before")

        out = td / "wrap.docx"
        audit = apply_patch(
            SRC,
            guarded_patch(
                "dotnet regression: wrap paragraph",
                [
                    {
                        "op": "wrap_paragraph_with_content_control",
                        "paragraph_index": 6,
                        "tag": "WORD-AI:DOTNET:background",
                        "alias": ".NET 项目背景锚点",
                        "expected_old_sha256": oldp["text_sha256"],
                    }
                ],
            ),
            out,
        )
        assert audit["validation"]["ok"], audit["validation"]
        tags = [cc["tag"] for cc in list_content_controls(out)["content_controls"]]
        assert "WORD-AI:DOTNET:background" in tags
        ops_tested.append("wrap_paragraph_with_content_control")

    print(json.dumps({"ok": True, "ops_tested": ops_tested}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
