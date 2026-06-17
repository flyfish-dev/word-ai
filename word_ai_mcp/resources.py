from __future__ import annotations

import sys
from pathlib import Path


def runtime_root() -> Path:
    """Return the root that contains bundled runtime assets.

    In source and wheel installs this is the repository/package root. In a
    PyInstaller onefile build it is the temporary extraction directory that
    contains files added with --add-data/--add-binary.
    """
    frozen_root = getattr(sys, "_MEIPASS", None)
    if frozen_root:
        return Path(str(frozen_root)).resolve()
    return Path(__file__).resolve().parents[1]


def resource_path(*parts: str) -> Path:
    return runtime_root().joinpath(*parts)


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def executable_path() -> Path:
    return Path(sys.executable).resolve()
