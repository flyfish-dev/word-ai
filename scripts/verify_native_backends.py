#!/usr/bin/env python3
"""Verify packaged WordAi.OpenXml native backends and current-platform execution."""

from __future__ import annotations

import argparse
import json
import os
import platform
import stat
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.publish_native_backends import DEFAULT_RIDS, executable_name  # noqa: E402

SAMPLE = ROOT / "examples" / "sample_contract.docx"


def current_rid() -> str | None:
    machine = platform.machine().lower()
    arch = "arm64" if machine in {"arm64", "aarch64"} else "x64" if machine in {"x86_64", "amd64"} else None
    if not arch:
        return None
    if sys.platform == "darwin":
        return f"osx-{arch}"
    if sys.platform.startswith("linux"):
        libc = (platform.libc_ver()[0] or "").lower()
        prefix = "linux-musl" if "musl" in libc else "linux"
        return f"{prefix}-{arch}"
    if sys.platform.startswith("win"):
        return f"win-{arch}"
    return None


def verify_file(root: Path, rid: str) -> dict[str, str | int | bool]:
    exe = root / "dist" / "native" / rid / executable_name(rid)
    if not exe.exists():
        raise SystemExit(f"Missing native backend for {rid}: {exe}")
    size = exe.stat().st_size
    if size < 1024 * 1024:
        raise SystemExit(f"Native backend for {rid} is unexpectedly small: {size} bytes")
    executable = bool(exe.stat().st_mode & stat.S_IXUSR)
    if not rid.startswith("win-") and not executable:
        raise SystemExit(f"Native backend for {rid} is not executable: {exe}")
    return {"rid": rid, "path": str(exe), "size": size, "executable": executable}


def run_current(root: Path, rid: str) -> dict:
    exe = root / "dist" / "native" / rid / executable_name(rid)
    proc = subprocess.run(
        [str(exe), "inspect", str(SAMPLE)],
        cwd=root,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=60,
    )
    if proc.returncode != 0:
        raise SystemExit(proc.stderr.strip() or proc.stdout.strip() or f"{exe} exited with {proc.returncode}")
    payload = json.loads(proc.stdout)
    if not payload.get("sha256") or payload.get("content_control_count", 0) < 1:
        raise SystemExit(f"Native backend inspect smoke failed for {rid}: {payload}")
    return {"rid": rid, "inspect_sha256": payload["sha256"], "content_control_count": payload["content_control_count"]}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--require-all", action="store_true", help="Require every supported release RID.")
    parser.add_argument("--require-rids", nargs="*", default=[], help="Specific RIDs to require.")
    parser.add_argument("--skip-current-run", action="store_true", help="Do not execute the current-platform native backend.")
    args = parser.parse_args()

    root = args.root.resolve()
    rids = list(DEFAULT_RIDS if args.require_all else args.require_rids)
    if not rids:
        detected = current_rid()
        if not detected:
            raise SystemExit("Could not detect current RID.")
        rids = [detected]

    files = [verify_file(root, rid) for rid in rids]
    run = None
    detected = current_rid()
    if detected and detected in rids and not args.skip_current_run:
        run = run_current(root, detected)

    print(json.dumps({"ok": True, "files": files, "current_run": run}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
