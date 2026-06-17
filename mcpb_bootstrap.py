#!/usr/bin/env python3
"""Bootstrap Word AI from an MCPB bundle."""

from __future__ import annotations

import os
import subprocess
import sys
import venv
from pathlib import Path


DEFAULT_VERSION = "0.8.6"


def bundle_version(bundle_dir: Path) -> str:
    pyproject = bundle_dir / "pyproject.toml"
    if pyproject.exists():
        for line in pyproject.read_text(encoding="utf-8").splitlines():
            if line.startswith("version = "):
                return line.split("=", 1)[1].strip().strip('"')
    return DEFAULT_VERSION


def ensure_venv(bundle_dir: Path) -> Path:
    version = bundle_version(bundle_dir)
    venv_dir = bundle_dir / ".word-ai-mcpb-venv"
    python = venv_dir / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
    marker = venv_dir / ".word-ai-version"
    if python.exists() and marker.exists() and marker.read_text().strip() == version:
        return python

    venv.EnvBuilder(with_pip=True, clear=True).create(venv_dir)
    subprocess.check_call(
        [str(python), "-m", "pip", "install", "--upgrade", "pip"],
        stdout=sys.stderr,
        stderr=sys.stderr,
    )
    subprocess.check_call(
        [str(python), "-m", "pip", "install", str(bundle_dir)],
        stdout=sys.stderr,
        stderr=sys.stderr,
    )
    marker.write_text(version + "\n")
    return python


def main() -> int:
    if sys.version_info < (3, 10):
        raise SystemExit("Word AI MCPB requires Python 3.10 or newer.")
    bundle_dir = Path(__file__).resolve().parent
    python = ensure_venv(bundle_dir)
    env = os.environ.copy()
    env["PYTHONPATH"] = str(bundle_dir)
    os.execve(
        str(python),
        [str(python), "-m", "word_ai_mcp.server", *sys.argv[1:]],
        env,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
