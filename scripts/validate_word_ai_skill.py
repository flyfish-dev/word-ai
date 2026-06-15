from __future__ import annotations

import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def fail(message: str) -> None:
    raise SystemExit(f"skill validation failed: {message}")


def read(path: Path) -> str:
    if not path.exists():
        fail(f"missing {path}")
    return path.read_text(encoding="utf-8")


def validate_skill(skill_dir: Path) -> None:
    skill = read(skill_dir / "SKILL.md")
    if not skill.startswith("---\n"):
        fail(f"{skill_dir}/SKILL.md must start with YAML frontmatter")
    frontmatter_end = skill.find("\n---\n", 4)
    if frontmatter_end == -1:
        fail(f"{skill_dir}/SKILL.md must close YAML frontmatter")
    frontmatter = skill[4:frontmatter_end]
    if "name: word-ai" not in frontmatter:
        fail(f"{skill_dir}/SKILL.md must declare name: word-ai")
    if "description:" not in frontmatter or len(frontmatter) < 80:
        fail(f"{skill_dir}/SKILL.md must include a useful description")

    required_terms = [
        "docx_health_check",
        "docx_assess_patchset",
        "docx_dry_run_patchset",
        "docx_apply_patchset",
        "word_session_list",
        "word_session_preview_patchset",
        "word_session_apply_patchset",
        "OfficeCLI",
        "officecli view <file> html",
        "officecli view <file> screenshot",
        "officecli view <file> issues",
        "officecli query <file> <selector> --json",
        "officecli validate <file>",
        "officecli_view_html",
        "officecli_view_screenshot",
        "officecli_view_issues",
        "officecli_query",
        "officecli_validate",
        "officecli set",
        "officecli add",
        "officecli remove",
        "raw-set",
        "PatchSet",
        "expected_old_sha256",
    ]
    missing = [term for term in required_terms if term not in skill]
    if missing:
        fail(f"{skill_dir}/SKILL.md missing required terms: {', '.join(missing)}")

    agent = read(skill_dir / "agents" / "openai.yaml")
    for term in ["display_name", "Word AI", "default_prompt"]:
        if term not in agent:
            fail(f"{skill_dir}/agents/openai.yaml missing {term}")


def main() -> int:
    validate_skill(ROOT / "skills" / "word-ai")
    validate_skill(ROOT / "codex-skill")
    print("word-ai skill validation passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
