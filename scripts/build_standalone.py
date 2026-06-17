#!/usr/bin/env python3
"""Build a one-file Word AI executable with PyInstaller."""

from __future__ import annotations

import argparse
import json
import os
import stat
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.publish_native_backends import executable_name, host_rid, publish_rid  # noqa: E402


def read_version() -> str:
    for line in (ROOT / "pyproject.toml").read_text(encoding="utf-8").splitlines():
        if line.startswith("version = "):
            return line.split("=", 1)[1].strip().strip('"')
    raise RuntimeError("Could not read version from pyproject.toml")


def pyinstaller_command(python: str) -> list[str]:
    proc = subprocess.run(
        [python, "-m", "PyInstaller", "--version"],
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if proc.returncode != 0:
        raise SystemExit(
            "PyInstaller is required to build the standalone executable. "
            "Install it with: python -m pip install pyinstaller"
        )
    return [python, "-m", "PyInstaller"]


def add_data_arg(source: Path, dest: str) -> str:
    return f"{source}{os.pathsep}{dest}"


def add_binary_arg(source: Path, dest: str) -> str:
    return f"{source}{os.pathsep}{dest}"


def build(args: argparse.Namespace) -> dict[str, str | int]:
    rid = args.rid or host_rid()
    native_dir = ROOT / "dist" / "native" / rid
    native_exe = native_dir / executable_name(rid)
    if args.build_native or not native_exe.exists():
        publish_rid(rid, clean=args.clean_native)
    if not native_exe.exists():
        raise SystemExit(f"Missing native backend for {rid}: {native_exe}")

    python = args.python or sys.executable
    pyinstaller = pyinstaller_command(python)
    dist_dir = args.dist_dir or ROOT / "dist" / "standalone" / rid
    work_dir = args.work_dir or ROOT / "build" / "pyinstaller" / rid
    spec_dir = args.spec_dir or ROOT / "build" / "pyinstaller" / "spec"
    name = args.name or ("word-ai.exe" if os.name == "nt" else "word-ai")
    if name.endswith(".exe"):
        pyinstaller_name = name[:-4]
    else:
        pyinstaller_name = name

    cmd = [
        *pyinstaller,
        "--clean",
        "--noconfirm",
        "--onefile",
        "--name",
        pyinstaller_name,
        "--distpath",
        str(dist_dir),
        "--workpath",
        str(work_dir),
        "--specpath",
        str(spec_dir),
        "--add-data",
        add_data_arg(ROOT / "skills", "skills"),
        "--add-data",
        add_data_arg(ROOT / "schemas", "schemas"),
        "--add-data",
        add_data_arg(ROOT / "README.md", "."),
        "--add-data",
        add_data_arg(ROOT / "README.zh-CN.md", "."),
        "--add-binary",
        add_binary_arg(native_exe, f"dist/native/{rid}"),
        "--hidden-import",
        "lxml.etree",
        "--hidden-import",
        "lxml._elementpath",
        "--hidden-import",
        "docx",
        str(ROOT / "word_ai_mcp" / "standalone.py"),
    ]
    subprocess.run(cmd, cwd=ROOT, check=True)

    exe = dist_dir / name
    produced = dist_dir / pyinstaller_name
    if not exe.exists() and produced.exists():
        produced.rename(exe)
    if not exe.exists():
        raise SystemExit(f"Expected standalone executable was not produced: {exe}")
    if os.name != "nt":
        exe.chmod(exe.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)

    return {
        "ok": True,
        "version": args.version,
        "rid": rid,
        "path": str(exe),
        "size": exe.stat().st_size,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", default=read_version())
    parser.add_argument("--rid", default=None, help="Runtime identifier. Defaults to the host RID.")
    parser.add_argument("--python", default=sys.executable, help="Python interpreter with PyInstaller installed.")
    parser.add_argument("--name", default=None, help="Output executable name. Defaults to word-ai.")
    parser.add_argument("--dist-dir", type=Path, default=None)
    parser.add_argument("--work-dir", type=Path, default=None)
    parser.add_argument("--spec-dir", type=Path, default=None)
    parser.add_argument("--build-native", action="store_true", help="Build the current RID .NET native backend before packaging.")
    parser.add_argument("--clean-native", action="store_true", help="Clean the native RID output before building it.")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = build(args)
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"{result['rid']}: {result['path']} ({result['size']} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
