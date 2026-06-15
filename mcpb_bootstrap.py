#!/usr/bin/env python3
"""Bootstrap Word AI from an MCPB bundle."""

from __future__ import annotations

import os
import subprocess
import sys
import venv
from pathlib import Path


VERSION = "0.8.1"


def ensure_venv(bundle_dir: Path) -> Path:
    venv_dir = bundle_dir / ".word-ai-mcpb-venv"
    python = venv_dir / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
    marker = venv_dir / ".word-ai-version"
    if python.exists() and marker.exists() and marker.read_text().strip() == VERSION:
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
    marker.write_text(VERSION + "\n")
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
