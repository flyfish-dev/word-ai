from __future__ import annotations

import json
from pathlib import Path
from tempfile import TemporaryDirectory

from word_ai_mcp.server import WordAiMcpServer


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "examples" / "sample_contract.docx"
PATCH = ROOT / "examples" / "patches" / "replace_srs_sections.json"


def main() -> int:
    patchset = json.loads(PATCH.read_text(encoding="utf-8"))
    with TemporaryDirectory() as td:
        server = WordAiMcpServer(root=str(ROOT), allowed_roots=[td])
        status = server.offline_engine_status()
        assert status["selected"] in {"dotnet", "python"}, status

        forced_python = server.tool_docx_assess_patchset(
            {"docx_path": str(SAMPLE), "patchset": patchset, "engine": "python"}
        )
        assert forced_python["engine"] == "python", forced_python
        assert forced_python["ok"], forced_python

        auto_assess = server.tool_docx_assess_patchset({"docx_path": str(SAMPLE), "patchset": patchset})
        assert auto_assess["engine"] == status["selected"], (auto_assess, status)
        assert auto_assess["ok"], auto_assess

        out = Path(td) / "engine-route.docx"
        audit = server.tool_docx_apply_patchset(
            {"docx_path": str(SAMPLE), "patchset": patchset, "output_path": str(out)}
        )
        assert audit["engine"] == status["selected"], audit
        assert audit["validation"]["ok"], audit["validation"]
        assert out.exists(), out

        validation = server.tool_docx_validate({"source_docx": str(SAMPLE), "target_docx": str(out)})
        assert validation["engine"] == status["selected"], validation
        assert validation["ok"], validation
        protected = validation.get("metrics", {}).get("protected_object_checks", {})
        if status["selected"] == "dotnet":
            assert protected.get("touched_content_control_tags"), protected

    print(json.dumps({"ok": True, "selected_engine": status["selected"], "dotnet": status["dotnet"]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
