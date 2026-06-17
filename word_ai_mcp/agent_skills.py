from __future__ import annotations

import argparse
import json
import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from .resources import resource_path


SKILL_NAME = "word-ai"


@dataclass(frozen=True)
class AgentTarget:
    key: str
    label: str
    path: Path
    install_by_default: bool
    install_when_detected: bool = False

    @property
    def agent_home(self) -> Path:
        return self.path.parents[1]

    @property
    def detected(self) -> bool:
        return self.agent_home.exists()


@dataclass(frozen=True)
class InstallResult:
    key: str
    label: str
    path: Path
    action: str

    def as_dict(self) -> dict[str, str]:
        return {
            "key": self.key,
            "label": self.label,
            "path": str(self.path),
            "action": self.action,
        }


def repo_root(value: str | None = None) -> Path:
    if value:
        return Path(value).expanduser().resolve()
    return Path(__file__).resolve().parent.parent


def source_skill_dir(root: Path) -> Path:
    candidates = [
        root / "skills" / SKILL_NAME,
        resource_path("skills", SKILL_NAME),
        resource_path("codex-skill"),
    ]
    for candidate in candidates:
        if (candidate / "SKILL.md").exists():
            return candidate
    searched = ", ".join(str(path) for path in candidates)
    raise FileNotFoundError(f"missing bundled {SKILL_NAME} skill. Searched: {searched}")


def _env_home(name: str, fallback: Path) -> Path:
    value = os.environ.get(name)
    return Path(value).expanduser() if value else fallback


def known_targets(root: Path) -> list[AgentTarget]:
    home = Path.home()
    codex_home = _env_home("CODEX_HOME", home / ".codex")
    claude_home = _env_home("CLAUDE_HOME", home / ".claude")
    return [
        AgentTarget(
            key="codex",
            label="OpenAI Codex user skills",
            path=home / ".agents" / "skills" / SKILL_NAME,
            install_by_default=True,
        ),
        AgentTarget(
            key="codex-legacy",
            label="Codex app compatibility skills",
            path=codex_home / "skills" / SKILL_NAME,
            install_by_default=True,
        ),
        AgentTarget(
            key="claude",
            label="Claude Code personal skills",
            path=claude_home / "skills" / SKILL_NAME,
            install_by_default=True,
        ),
        AgentTarget(
            key="codex-project",
            label="Codex repository skills",
            path=root / ".agents" / "skills" / SKILL_NAME,
            install_by_default=False,
        ),
        AgentTarget(
            key="claude-project",
            label="Claude Code project skills",
            path=root / ".claude" / "skills" / SKILL_NAME,
            install_by_default=False,
        ),
        AgentTarget(
            key="cursor",
            label="Cursor skills",
            path=home / ".cursor" / "skills" / SKILL_NAME,
            install_by_default=False,
            install_when_detected=True,
        ),
        AgentTarget(
            key="windsurf",
            label="Windsurf skills",
            path=home / ".windsurf" / "skills" / SKILL_NAME,
            install_by_default=False,
            install_when_detected=True,
        ),
        AgentTarget(
            key="copilot",
            label="GitHub Copilot skills",
            path=home / ".copilot" / "skills" / SKILL_NAME,
            install_by_default=False,
            install_when_detected=True,
        ),
        AgentTarget(
            key="openclaw",
            label="OpenClaw skills",
            path=home / ".openclaw" / "skills" / SKILL_NAME,
            install_by_default=False,
            install_when_detected=True,
        ),
    ]


def parse_agent_selector(value: str) -> set[str]:
    out: set[str] = set()
    for item in value.split(","):
        item = item.strip().lower()
        if item:
            out.add(item)
    return out or {"auto"}


def selected_targets(root: Path, selector: str, include_project: bool = False) -> list[AgentTarget]:
    requested = parse_agent_selector(selector)
    targets = known_targets(root)
    by_key = {target.key: target for target in targets}
    if "all" in requested:
        chosen = targets
    elif "auto" in requested:
        chosen = [
            target
            for target in targets
            if target.install_by_default or (target.install_when_detected and target.detected)
        ]
    else:
        unknown = sorted(key for key in requested if key not in by_key)
        if unknown:
            known = ", ".join(sorted(by_key))
            raise ValueError(f"unknown agent target(s): {', '.join(unknown)}. Known targets: {known}, auto, all")
        chosen = [by_key[key] for key in requested]
    if include_project:
        for project_key in ("codex-project", "claude-project"):
            target = by_key[project_key]
            if target not in chosen:
                chosen.append(target)
    return chosen


def validate_source_skill(source: Path) -> None:
    skill = source / "SKILL.md"
    if not skill.exists():
        raise FileNotFoundError(f"missing skill file: {skill}")
    text = skill.read_text(encoding="utf-8")
    if not text.startswith("---\n") or "\n---\n" not in text[4:]:
        raise ValueError(f"{skill} must contain YAML frontmatter")
    frontmatter = text[4 : text.find("\n---\n", 4)]
    if "description:" not in frontmatter:
        raise ValueError(f"{skill} must include a description in frontmatter")


def copy_skill(source: Path, target: Path, dry_run: bool = False) -> str:
    existed = (target / "SKILL.md").exists()
    if dry_run:
        return "would-update" if existed else "would-install"
    target.mkdir(parents=True, exist_ok=True)
    shutil.copytree(source, target, dirs_exist_ok=True)
    return "updated" if existed else "installed"


def install_agent_skills(root: Path, selector: str = "auto", include_project: bool = False, dry_run: bool = False) -> list[InstallResult]:
    source = source_skill_dir(root)
    validate_source_skill(source)
    results: list[InstallResult] = []
    for target in selected_targets(root, selector, include_project=include_project):
        action = copy_skill(source, target.path, dry_run=dry_run)
        results.append(InstallResult(target.key, target.label, target.path, action))
    return results


def print_results(results: Iterable[InstallResult]) -> None:
    for result in results:
        print(f"{result.action:13} {result.label}: {result.path / 'SKILL.md'}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Install the Word AI skill into local agent skill directories.")
    parser.add_argument("--root", default=None, help="Repository root. Defaults to this package's repository.")
    parser.add_argument(
        "--agents",
        default="auto",
        help="Comma-separated targets: auto, all, codex, codex-legacy, claude, cursor, windsurf, copilot, openclaw.",
    )
    parser.add_argument("--project", action="store_true", help="Also install repository-scoped .agents and .claude skills.")
    parser.add_argument("--dry-run", action="store_true", help="Print target paths without writing files.")
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON.")
    args = parser.parse_args(argv)

    root = repo_root(args.root)
    results = install_agent_skills(root, selector=args.agents, include_project=args.project, dry_run=args.dry_run)
    if args.json:
        print(json.dumps([result.as_dict() for result in results], indent=2))
    else:
        print_results(results)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
