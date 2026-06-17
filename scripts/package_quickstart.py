#!/usr/bin/env python3
"""Package a Word AI standalone executable for fast agent integration."""

from __future__ import annotations

import argparse
import hashlib
import json
import stat
import tarfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
FIXED_TIME = (2026, 1, 1, 0, 0, 0)
FIXED_MTIME = 1767225600


def read_version() -> str:
    for line in (ROOT / "pyproject.toml").read_text(encoding="utf-8").splitlines():
        if line.startswith("version = "):
            return line.split("=", 1)[1].strip().strip('"')
    raise RuntimeError("Could not read version from pyproject.toml")


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def quickstart_readme(version: str, rid: str, exe_name: str) -> str:
    return f"""# Word AI Quickstart Bundle

Version: {version}
Runtime ID: {rid}

This package contains a single-file `word-ai` executable. It bundles the Python
MCP facade, Python dependencies, the current-platform .NET Open XML backend,
schemas, and the official `word-ai` agent skill template.

## Start the MCP server

```bash
./{exe_name} mcp --root "$PWD" --allow-root "$HOME/Downloads" --allow-root "$HOME/Documents"
```

## Install the agent skill

```bash
./{exe_name} install-skill
```

## Generate a Codex MCP config snippet

```bash
./{exe_name} codex-config --output .wordai/codex-config.toml
```

## Check runtime readiness

```bash
./{exe_name} doctor
```

No local Python, pip install, .NET SDK, or source checkout is required for the
offline DOCX editing path. Full Office.js taskpane development still uses the
source install path.
"""


def manifest(version: str, rid: str, exe_name: str, exe_sha256: str) -> dict:
    return {
        "name": "word-ai-quickstart",
        "version": version,
        "rid": rid,
        "executable": exe_name,
        "executable_sha256": exe_sha256,
        "commands": {
            "mcp": f"./{exe_name} mcp --root <workspace>",
            "install_skill": f"./{exe_name} install-skill",
            "codex_config": f"./{exe_name} codex-config --output .wordai/codex-config.toml",
            "doctor": f"./{exe_name} doctor",
        },
    }


def tar_filter(info: tarfile.TarInfo) -> tarfile.TarInfo:
    info.mtime = FIXED_MTIME
    info.uid = 0
    info.gid = 0
    info.uname = ""
    info.gname = ""
    return info


def add_tar_file(tf: tarfile.TarFile, source: Path, arcname: str, mode: int | None = None) -> None:
    info = tf.gettarinfo(str(source), arcname)
    info = tar_filter(info)
    if mode is not None:
        info.mode = mode
    with source.open("rb") as fh:
        tf.addfile(info, fh)


def add_zip_file(zf: zipfile.ZipFile, source: Path, arcname: str, mode: int | None = None) -> None:
    info = zipfile.ZipInfo(arcname, FIXED_TIME)
    info.compress_type = zipfile.ZIP_DEFLATED
    file_mode = mode if mode is not None else stat.S_IMODE(source.stat().st_mode)
    info.external_attr = (stat.S_IFREG | file_mode) << 16
    zf.writestr(info, source.read_bytes())


def write_text_temp(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def package(args: argparse.Namespace) -> dict[str, str | int]:
    version = args.version
    rid = args.rid
    exe = args.executable or ROOT / "dist" / "standalone" / rid / ("word-ai.exe" if rid.startswith("win-") else "word-ai")
    exe = exe.resolve()
    if not exe.exists():
        raise SystemExit(f"Standalone executable not found: {exe}. Run scripts/build_standalone.py first.")
    exe_name = "word-ai.exe" if rid.startswith("win-") else "word-ai"
    out_dir = args.out_dir
    out_dir.mkdir(parents=True, exist_ok=True)
    suffix = "zip" if rid.startswith("win-") else "tar.gz"
    out = out_dir / f"word-ai-quickstart-{version}-{rid}.{suffix}"

    tmp = ROOT / "build" / "quickstart" / rid
    readme_path = tmp / "README.quickstart.md"
    manifest_path = tmp / "word-ai-quickstart.json"
    write_text_temp(readme_path, quickstart_readme(version, rid, exe_name))
    write_text_temp(manifest_path, json.dumps(manifest(version, rid, exe_name, sha256(exe)), indent=2) + "\n")

    prefix = f"word-ai-quickstart-{version}-{rid}"
    if out.exists():
        out.unlink()
    if suffix == "zip":
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            add_zip_file(zf, exe, f"{prefix}/{exe_name}", mode=0o755)
            add_zip_file(zf, readme_path, f"{prefix}/README.md")
            add_zip_file(zf, manifest_path, f"{prefix}/word-ai-quickstart.json")
            add_zip_file(zf, ROOT / "LICENSE", f"{prefix}/LICENSE")
    else:
        with tarfile.open(out, "w:gz") as tf:
            add_tar_file(tf, exe, f"{prefix}/{exe_name}", mode=0o755)
            add_tar_file(tf, readme_path, f"{prefix}/README.md")
            add_tar_file(tf, manifest_path, f"{prefix}/word-ai-quickstart.json")
            add_tar_file(tf, ROOT / "LICENSE", f"{prefix}/LICENSE")

    return {"ok": True, "path": str(out), "sha256": sha256(out), "size": out.stat().st_size}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--version", default=read_version())
    parser.add_argument("--rid", required=True)
    parser.add_argument("--executable", type=Path, default=None)
    parser.add_argument("--out-dir", type=Path, default=ROOT / "dist" / "quickstart")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    result = package(args)
    if args.json:
        print(json.dumps(result, indent=2))
    else:
        print(f"{result['sha256']}  {result['path']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
