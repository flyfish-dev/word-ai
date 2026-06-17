from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

from .openxml_engine import dotnet_status


WRITE_APPROVAL_TOOLS = [
    "docx_export_plain_text",
    "docx_export_table_csv",
    "docx_table_to_csv",
    "docx_preflight_patchset",
    "docx_dry_run_patchset",
    "docx_write_index",
    "docx_backup",
    "docx_restore_backup",
    "docx_rollback",
    "docx_apply_patchset",
    "word_session_apply_patchset",
    "word_session_wrap_selection",
    "word_session_rollback",
    "officecli_view_screenshot",
]


def repo_root(value: str | None = None) -> Path:
    if value:
        return Path(value).expanduser().resolve()
    return Path(__file__).resolve().parent.parent


def preferred_python(root: Path) -> Path:
    if os.name == "nt":
        candidate = root / ".venv" / "Scripts" / "python.exe"
    else:
        candidate = root / ".venv" / "bin" / "python"
    return candidate if candidate.exists() else Path(sys.executable).resolve()


def toml_string(value: str) -> str:
    return json.dumps(value)


def default_allowed_roots(root: Path) -> list[Path]:
    home = Path.home()
    candidates = [
        home / "Downloads",
        home / "Documents",
        home / "Desktop",
    ]
    out: list[Path] = []
    for candidate in candidates:
        path = candidate.expanduser().resolve()
        if path.exists() and path != root and path not in out:
            out.append(path)
    return out


def find_officecli() -> Path | None:
    configured = os.environ.get("WORD_AI_OFFICECLI")
    candidates = []
    if configured:
        candidates.append(Path(configured).expanduser())
    found = shutil.which("officecli")
    if found:
        candidates.append(Path(found))
    candidates.append(Path.home() / ".local" / "bin" / "officecli")
    for candidate in candidates:
        path = candidate.expanduser()
        if path.exists() and os.access(path, os.X_OK):
            return path.resolve()
    return None


def build_codex_config(root: Path, server_name: str = "word_ai", allowed_roots: list[str] | None = None, include_common_user_roots: bool = True) -> str:
    python_path = preferred_python(root)
    extra_roots = [Path(p).expanduser().resolve() for p in (allowed_roots or [])]
    if include_common_user_roots:
        extra_roots = [*default_allowed_roots(root), *extra_roots]
    resolved_extra_roots: list[Path] = []
    for path in extra_roots:
        if path != root and path not in resolved_extra_roots:
            resolved_extra_roots.append(path)
    server_args = ["-m", "word_ai_mcp.server", "--root", str(root)]
    for allowed in resolved_extra_roots:
        server_args.extend(["--allow-root", str(allowed)])
    lines = [
        f"[mcp_servers.{server_name}]",
        f"command = {toml_string(str(python_path))}",
        f"args = {json.dumps(server_args)}",
        "enabled = true",
        "startup_timeout_sec = 30",
        "",
        f"[mcp_servers.{server_name}.env]",
        f"PYTHONPATH = {toml_string(str(root))}",
    ]
    officecli = find_officecli()
    if officecli:
        lines.append(f"WORD_AI_OFFICECLI = {toml_string(str(officecli))}")
    lines.append("")
    for tool in WRITE_APPROVAL_TOOLS:
        lines.extend(
            [
                f"[mcp_servers.{server_name}.tools.{tool}]",
                'approval_mode = "approve"',
                "",
            ]
        )
    return "\n".join(lines).rstrip() + "\n"


def command_version(command: list[str], cwd: Path | None = None) -> str | None:
    try:
        proc = subprocess.run(command, cwd=str(cwd) if cwd else None, text=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=12)
    except (OSError, subprocess.SubprocessError):
        return None
    if proc.returncode != 0:
        return None
    return proc.stdout.strip().splitlines()[0] if proc.stdout.strip() else "ok"


def doctor(root: Path) -> int:
    checks: list[tuple[str, bool, str]] = []
    checks.append(("python", sys.version_info >= (3, 10), sys.version.split()[0]))
    for module_name in ["lxml", "docx"]:
        try:
            __import__(module_name)
            checks.append((module_name, True, "installed"))
        except Exception as exc:
            checks.append((module_name, False, str(exc)))

    node = command_version(["node", "--version"])
    npm = command_version(["npm", "--version"])
    dotnet = command_version(["dotnet", "--version"])
    openxml = dotnet_status()
    officecli_path = find_officecli()
    officecli = command_version([str(officecli_path), "--version"]) if officecli_path else None
    checks.extend(
        [
            ("node", node is not None, node or "not found"),
            ("npm", npm is not None, npm or "not found"),
            ("dotnet", dotnet is not None, dotnet or "not found"),
            (
                "openxml backend",
                bool(openxml.get("available")),
                f"{openxml.get('mode')} ({openxml.get('runtime_id')})" if openxml.get("available") else str(openxml.get("reason")),
            ),
            ("office-addin package", (root / "office-addin" / "package.json").exists(), "office-addin/package.json"),
            (
                "office-addin node_modules",
                (root / "office-addin" / "node_modules").exists(),
                "office-addin/node_modules" if (root / "office-addin" / "node_modules").exists() else "run scripts/install.sh if missing",
            ),
            ("sample docx", (root / "examples" / "sample_contract.docx").exists(), "examples/sample_contract.docx"),
        ]
    )
    optional_checks = [("officecli", officecli is not None, officecli or "not installed; optional auxiliary backend")]

    width = max(len(name) for name, _, _ in checks + optional_checks)
    ok = True
    for name, passed, detail in checks:
        ok = ok and passed
        status = "OK" if passed else "MISSING"
        print(f"{status:7} {name:<{width}} {detail}")
    for name, passed, detail in optional_checks:
        status = "OK" if passed else "OPTION"
        print(f"{status:7} {name:<{width}} {detail}")

    print()
    print(f"Repo root: {root}")
    print(f"Python for Codex: {preferred_python(root)}")
    print("Default allowed roots:")
    for allowed in default_allowed_roots(root):
        print(f"  - {allowed}")
    print(f"Codex config snippet: {root / '.wordai' / 'codex-config.toml'}")
    return 0 if ok else 1


def write_codex_config(root: Path, server_name: str, output: str | None, allowed_roots: list[str] | None, include_common_user_roots: bool) -> int:
    text = build_codex_config(root, server_name, allowed_roots=allowed_roots, include_common_user_roots=include_common_user_roots)
    if output:
        out = Path(output).expanduser()
        if not out.is_absolute():
            out = root / out
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text, encoding="utf-8")
        print(f"Wrote {out}")
    else:
        print(text, end="")
    return 0


def install_skills(root: Path, agents: str, include_project: bool, dry_run: bool, json_output: bool) -> int:
    from .agent_skills import install_agent_skills, print_results

    results = install_agent_skills(root, selector=agents, include_project=include_project, dry_run=dry_run)
    if json_output:
        print(json.dumps([result.as_dict() for result in results], indent=2))
    else:
        print_results(results)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Word AI quickstart helpers.")
    parser.add_argument("--root", default=None, help="Repository root. Defaults to this package's repository.")
    sub = parser.add_subparsers(dest="command", required=True)

    doctor_parser = sub.add_parser("doctor", help="Check local runtime readiness.")
    doctor_parser.set_defaults(func=lambda args: doctor(repo_root(args.root)))

    config_parser = sub.add_parser("codex-config", help="Print or write a Codex MCP config snippet.")
    config_parser.add_argument("--server-name", default="word_ai")
    config_parser.add_argument("--output", default=None)
    config_parser.add_argument("--allow-root", action="append", default=[], help="Additional allowed directory. Repeatable.")
    config_parser.add_argument("--no-common-user-roots", action="store_true", help="Do not include ~/Downloads, ~/Documents, and ~/Desktop automatically.")
    config_parser.set_defaults(
        func=lambda args: write_codex_config(
            repo_root(args.root),
            args.server_name,
            args.output,
            args.allow_root,
            not args.no_common_user_roots,
        )
    )

    skills_parser = sub.add_parser("install-skills", help="Install Word AI agent skills for Codex, Claude Code, and compatible clients.")
    skills_parser.add_argument(
        "--agents",
        default="auto",
        help="Comma-separated targets: auto, all, codex, codex-legacy, claude, cursor, windsurf, copilot, openclaw.",
    )
    skills_parser.add_argument("--project", action="store_true", help="Also install repository-scoped .agents and .claude skills.")
    skills_parser.add_argument("--dry-run", action="store_true", help="Print target paths without writing files.")
    skills_parser.add_argument("--json", action="store_true", help="Print machine-readable JSON.")
    skills_parser.set_defaults(
        func=lambda args: install_skills(
            repo_root(args.root),
            args.agents,
            args.project,
            args.dry_run,
            args.json,
        )
    )

    args = parser.parse_args(argv)
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
