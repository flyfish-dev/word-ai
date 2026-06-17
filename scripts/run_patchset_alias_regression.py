from __future__ import annotations

import json
from pathlib import Path
from tempfile import TemporaryDirectory

from word_ai_mcp.server import WordAiMcpServer


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "examples" / "sample_contract.docx"
TAG = "WORD-AI:SRS:1.0:overview"


def run_cycle(server: WordAiMcpServer, patchset: dict, out: Path, expected_engine: str, engine: str | None = None) -> None:
    base_args = {"docx_path": str(SAMPLE), "patchset": patchset}
    if engine:
        base_args["engine"] = engine

    assess = server.tool_docx_assess_patchset(base_args)
    assert assess["ok"], assess
    assert assess["engine"] == expected_engine, (assess, expected_engine)
    assert assess["touched"]["content_control_tags"] == [TAG], assess

    dry_run = server.tool_docx_dry_run_patchset({**base_args, "keep_output": False})
    assert dry_run["validation"]["ok"], dry_run["validation"]

    audit = server.tool_docx_apply_patchset({**base_args, "output_path": str(out)})
    assert audit["validation"]["ok"], audit["validation"]
    assert out.exists(), out
    assert out.with_suffix(".audit.json").exists(), out.with_suffix(".audit.json")


def main() -> int:
    with TemporaryDirectory() as temp_dir:
        temp = Path(temp_dir)
        server = WordAiMcpServer(root=str(ROOT), allowed_roots=[str(temp)])
        engine = server.offline_engine_status()["selected"]
        inspect = server.tool_docx_inspect({"docx_path": str(SAMPLE)})
        current = server.tool_docx_read_content_control({"docx_path": str(SAMPLE), "tag": TAG})
        patchset = {
            "schemaVersion": "2.0",
            "strict": True,
            "sourceSha256": inspect["sha256"],
            "reason": "patchset alias regression",
            "guard": {
                "requirePreconditions": True,
                "allowOverwrite": False,
            },
            "operations": [
                {
                    "operation": "replaceContentControlText",
                    "target_tag": TAG,
                    "text_sha256": current["text_sha256"],
                    "new_text": "[[CC:overview]] PatchSet alias regression replacement.",
                    "preserveStyle": True,
                    "allowComplexContent": False,
                }
            ],
        }

        run_cycle(server, patchset, temp / "alias-regression-auto.docx", engine)
        run_cycle(server, patchset, temp / "alias-regression-python.docx", "python", "python")

        invalid = {
            "schema_version": "2.0",
            "operations": [{"target_tag": TAG, "text": "missing op"}],
        }
        try:
            server.tool_docx_assess_patchset({"docx_path": str(SAMPLE), "patchset": invalid})
        except ValueError as exc:
            assert "patchset.operations[0].op is required" in str(exc), exc
        else:
            raise AssertionError("missing operation name should fail with a clear PatchSet error")

    print(json.dumps({"ok": True, "selected_engine": engine, "covered": ["operation", "target_tag", "new_text", "text_sha256"]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
