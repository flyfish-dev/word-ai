#!/usr/bin/env python3
"""Build a deterministic Word AI MCPB bundle."""

from __future__ import annotations

import argparse
import hashlib
import json
import stat
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
FIXED_TIME = (2026, 1, 1, 0, 0, 0)


def read_version() -> str:
    for line in (ROOT / "pyproject.toml").read_text().splitlines():
        if line.startswith("version = "):
            return line.split("=", 1)[1].strip().strip('"')
    raise RuntimeError("Could not read version from pyproject.toml")


def manifest(version: str) -> dict:
    return {
        "manifest_version": "0.3",
        "name": "word-ai",
        "display_name": "Word AI",
        "version": version,
        "description": "Structure-preserving Word DOCX editing MCP server for AI agents.",
        "long_description": (
            "Word AI edits Microsoft Word DOCX files through constrained PatchSet operations, "
            "hash preconditions, dry-run, validation, audit JSON, rollback, and diff workflows."
        ),
        "author": {
            "name": "flyfish-dev",
            "url": "https://github.com/flyfish-dev",
        },
        "repository": {
            "type": "git",
            "url": "https://github.com/flyfish-dev/word-ai.git",
        },
        "homepage": "https://github.com/flyfish-dev/word-ai",
        "documentation": "https://github.com/flyfish-dev/word-ai#readme",
        "support": "https://github.com/flyfish-dev/word-ai/issues",
        "license": "AGPL-3.0-or-later",
        "keywords": [
            "mcp",
            "docx",
            "word",
            "office-js",
            "openxml",
            "codex",
            "ai-agents",
        ],
        "server": {
            "type": "python",
            "entry_point": "mcpb_bootstrap.py",
            "mcp_config": {
                "command": "python",
                "args": [
                    "${__dirname}/mcpb_bootstrap.py",
                    "--root",
                    "${user_config.workspace_root}",
                    "--allow-root",
                    "${user_config.allowed_root_1}",
                    "--allow-root",
                    "${user_config.allowed_root_2}",
                ],
                "env": {
                    "PYTHONPATH": "${__dirname}",
                },
            },
        },
        "tools_generated": True,
        "tools": [
            {
                "name": "docx_health_check",
                "description": "Inspect DOCX structure, anchors, fields, comments, revisions, and edit risk.",
            },
            {
                "name": "docx_dry_run_patchset",
                "description": "Dry-run a constrained PatchSet before writing an output DOCX.",
            },
            {
                "name": "docx_apply_patchset",
                "description": "Apply a validated PatchSet to a new DOCX with audit JSON.",
            },
            {
                "name": "word_session_apply_patchset",
                "description": "Apply supported content-control edits to an open Word session through Office.js.",
            },
        ],
        "user_config": {
            "workspace_root": {
                "type": "directory",
                "title": "Primary Word AI workspace",
                "description": "Main folder for relative DOCX paths and Word AI sidecar files.",
                "required": True,
                "default": "${DOCUMENTS}",
            },
            "allowed_root_1": {
                "type": "directory",
                "title": "Additional document folder 1",
                "description": "A second folder Word AI may read and write, such as Downloads.",
                "required": False,
                "default": "${DOWNLOADS}",
            },
            "allowed_root_2": {
                "type": "directory",
                "title": "Additional document folder 2",
                "description": "A third folder Word AI may read and write, such as Desktop.",
                "required": False,
                "default": "${DESKTOP}",
            },
        },
        "compatibility": {
            "platforms": ["darwin", "win32", "linux"],
            "runtimes": {
                "python": ">=3.10",
            },
        },
    }


def iter_files() -> list[Path]:
    roots = [
        ROOT / "word_ai_mcp",
        ROOT / "schemas",
        ROOT / "docs",
    ]
    for optional_root in [ROOT / "native", ROOT / "dist" / "native"]:
        if optional_root.exists():
            roots.append(optional_root)
    files = [
        ROOT / "README.md",
        ROOT / "README.zh-CN.md",
        ROOT / "LICENSE",
        ROOT / "mcpb_bootstrap.py",
        ROOT / "pyproject.toml",
        ROOT / "requirements.txt",
    ]
    for root in roots:
        files.extend(path for path in root.rglob("*") if path.is_file())
    return sorted(
        path
        for path in files
        if "__pycache__" not in path.parts and path.suffix not in {".pyc", ".pyo"}
    )


def write_entry(zf: zipfile.ZipFile, arcname: str, data: bytes, mode: int = 0o644) -> None:
    info = zipfile.ZipInfo(arcname, FIXED_TIME)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = (stat.S_IFREG | mode) << 16
    zf.writestr(info, data)


def build(out_path: Path, version: str) -> str:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, "w") as zf:
        write_entry(
            zf,
            "manifest.json",
            (json.dumps(manifest(version), indent=2) + "\n").encode(),
        )
        for path in iter_files():
            mode = stat.S_IMODE(path.stat().st_mode)
            write_entry(zf, path.relative_to(ROOT).as_posix(), path.read_bytes(), mode=mode)
    digest = hashlib.sha256(out_path.read_bytes()).hexdigest()
    return digest


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", default=read_version())
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    out = args.out or ROOT / "dist" / f"word-ai-{args.version}.mcpb"
    digest = build(out, args.version)
    print(json.dumps({"path": str(out), "sha256": digest}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
