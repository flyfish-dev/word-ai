from __future__ import annotations

import json
from pathlib import Path

from word_ai_mcp.ooxml import (
    apply_patchset,
    assess_patchset,
    diff_text,
    health_check,
    inspect_docx,
    list_fields,
    list_images,
    list_paragraphs,
    list_tables,
    validate_structure,
    write_sidecar_index,
)

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "examples" / "sample_contract.docx"
PATCH = ROOT / "examples" / "patches" / "replace_srs_sections.json"
OUT = ROOT / "examples" / "sample_contract.edited.docx"


def main() -> None:
    print("Inspecting sample...")
    info = inspect_docx(SAMPLE)
    assert info["content_control_count"] == 3, info
    print(json.dumps({k: info[k] for k in ["paragraph_count", "table_count", "content_control_count", "field_count"]}, ensure_ascii=False, indent=2))

    print("Health check...")
    health = health_check(SAMPLE)
    assert health["ok_for_ai_editing"], health

    print("Listing rich read surfaces...")
    assert list_paragraphs(SAMPLE)["count"] > 0
    assert list_tables(SAMPLE)["table_count"] == 2
    assert list_fields(SAMPLE)["field_count"] == 0
    assert list_images(SAMPLE)["image_count"] == 0

    idx = write_sidecar_index(SAMPLE)
    print(f"Index written: {idx}")

    print("Assessing patchset...")
    patchset = json.loads(PATCH.read_text(encoding="utf-8"))
    assessment = assess_patchset(SAMPLE, patchset)
    print(json.dumps(assessment, ensure_ascii=False, indent=2))
    assert assessment["ok"], assessment

    print("Applying patchset...")
    for stale in [OUT, OUT.with_suffix(".audit.json"), OUT.with_suffix(".invalid.audit.json"), OUT.with_suffix(".invalid.docx")]:
        stale.unlink(missing_ok=True)
    for candidate in OUT.parent.glob(OUT.name + ".candidate.*"):
        candidate.unlink(missing_ok=True)
    audit = apply_patchset(SAMPLE, patchset, OUT)
    print(json.dumps(audit["validation"], ensure_ascii=False, indent=2))
    assert audit["validation"]["ok"], audit["validation"]

    touched_tags = [x["tag"] for x in audit["applied"] if x.get("tag")]
    touched_para_ids = [x.get("paraId") or x.get("anchor_paraId") for x in audit["applied"] if x.get("paraId") or x.get("anchor_paraId")]
    touched_tables = [x["table_index"] for x in audit["applied"] if x.get("table_index")]
    report = validate_structure(
        SAMPLE,
        OUT,
        touched_content_control_tags=touched_tags,
        touched_para_ids=touched_para_ids,
        touched_table_indices=touched_tables,
    )
    assert report.ok, report.to_dict()

    print("Diff preview:")
    print(diff_text(SAMPLE, OUT))
    print(f"OK: {OUT}")


if __name__ == "__main__":
    main()
