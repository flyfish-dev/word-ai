#!/usr/bin/env python3
"""Publish self-contained WordAi.OpenXml native backends for release RIDs."""

from __future__ import annotations

import argparse
import json
import os
import platform
import shutil
import stat
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PROJECT = ROOT / "dotnet" / "WordAi.OpenXml" / "WordAi.OpenXml.csproj"
DEFAULT_RIDS = (
    "osx-arm64",
    "osx-x64",
    "linux-x64",
    "linux-arm64",
    "linux-musl-x64",
    "linux-musl-arm64",
    "win-x64",
    "win-arm64",
)


def host_rid() -> str:
    machine = platform.machine().lower()
    arch = "arm64" if machine in {"arm64", "aarch64"} else "x64" if machine in {"x86_64", "amd64"} else ""
    if not arch:
        raise SystemExit(f"Unsupported host architecture: {platform.machine()}")
    if sys.platform == "darwin":
        return f"osx-{arch}"
    if sys.platform.startswith("linux"):
        return f"linux-{arch}"
    if sys.platform.startswith("win"):
        return f"win-{arch}"
    raise SystemExit(f"Unsupported host platform: {sys.platform}")


def dotnet_command() -> str:
    configured = os.environ.get("WORD_AI_DOTNET")
    if configured:
        return configured
    homebrew_dotnet = Path("/opt/homebrew/opt/dotnet@8/libexec/dotnet")
    if homebrew_dotnet.exists():
        return str(homebrew_dotnet)
    found = shutil.which("dotnet")
    if not found:
        raise SystemExit("dotnet was not found. Install .NET SDK 8 or set WORD_AI_DOTNET.")
    return found


def executable_name(rid: str) -> str:
    return "WordAi.OpenXml.exe" if rid.startswith("win-") else "WordAi.OpenXml"


def publish_rid(rid: str, *, clean: bool = False) -> dict[str, str | int]:
    out = ROOT / "dist" / "native" / rid
    if clean and out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True, exist_ok=True)
    cmd = [
        dotnet_command(),
        "publish",
        str(PROJECT),
        "-c",
        "Release",
        "-r",
        rid,
        "--self-contained",
        "true",
        "-p:UseAppHost=true",
        "-p:PublishSingleFile=true",
        "-p:PublishTrimmed=false",
        "-p:EnableCompressionInSingleFile=true",
        "-p:DebugType=none",
        "-p:DebugSymbols=false",
        "-o",
        str(out),
    ]
    env = os.environ.copy()
    homebrew_root = Path("/opt/homebrew/opt/dotnet@8/libexec")
    if homebrew_root.exists() and not env.get("DOTNET_ROOT"):
        env["DOTNET_ROOT"] = str(homebrew_root)
    subprocess.run(cmd, cwd=ROOT, env=env, check=True)

    exe = out / executable_name(rid)
    if not exe.exists():
        raise SystemExit(f"Expected native backend was not produced: {exe}")
    if not rid.startswith("win-"):
        exe.chmod(exe.stat().st_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
    for pdb in out.glob("*.pdb"):
        pdb.unlink()
    return {"rid": rid, "path": str(exe), "size": exe.stat().st_size}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("rids", nargs="*", help="Runtime identifiers to publish. Defaults to the host RID.")
    parser.add_argument("--all", action="store_true", help="Publish all supported release RIDs.")
    parser.add_argument("--clean", action="store_true", help="Remove each RID output directory before publishing.")
    parser.add_argument("--json", action="store_true", help="Print a JSON summary.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    rids = list(DEFAULT_RIDS if args.all else args.rids or [host_rid()])
    seen: set[str] = set()
    results = []
    for rid in rids:
        if rid in seen:
            continue
        if rid not in DEFAULT_RIDS:
            raise SystemExit(f"Unsupported release RID: {rid}. Supported: {', '.join(DEFAULT_RIDS)}")
        seen.add(rid)
        print(f"Publishing WordAi.OpenXml native backend for {rid}...", file=sys.stderr)
        results.append(publish_rid(rid, clean=args.clean))
    if args.json:
        print(json.dumps({"ok": True, "rids": results}, indent=2))
    else:
        for result in results:
            print(f"{result['rid']}: {result['path']} ({result['size']} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
