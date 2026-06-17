#!/usr/bin/env python3
"""Create per-RID GitHub Release assets for WordAi.OpenXml native backends."""

from __future__ import annotations

import argparse
import gzip
import hashlib
import stat
import tarfile
import zipfile
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.publish_native_backends import DEFAULT_RIDS  # noqa: E402

FIXED_TIME = (2026, 1, 1, 0, 0, 0)
FIXED_MTIME = 1767225600


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def iter_rid_files(rid_dir: Path) -> list[Path]:
    return sorted(path for path in rid_dir.rglob("*") if path.is_file())


def tar_filter(info: tarfile.TarInfo) -> tarfile.TarInfo:
    info.mtime = FIXED_MTIME
    info.uid = 0
    info.gid = 0
    info.uname = ""
    info.gname = ""
    return info


def make_tar_gz(rid_dir: Path, out: Path) -> None:
    with out.open("wb") as raw:
        with gzip.GzipFile(filename="", mode="wb", fileobj=raw, mtime=FIXED_MTIME) as gz:
            with tarfile.open(fileobj=gz, mode="w") as tf:
                for path in iter_rid_files(rid_dir):
                    tf.add(path, arcname=f"{rid_dir.name}/{path.relative_to(rid_dir).as_posix()}", filter=tar_filter)


def make_zip(rid_dir: Path, out: Path) -> None:
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in iter_rid_files(rid_dir):
            arcname = f"{rid_dir.name}/{path.relative_to(rid_dir).as_posix()}"
            info = zipfile.ZipInfo(arcname, FIXED_TIME)
            mode = stat.S_IMODE(path.stat().st_mode)
            info.external_attr = (stat.S_IFREG | mode) << 16
            zf.writestr(info, path.read_bytes())


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", required=True)
    parser.add_argument("--out-dir", type=Path, default=ROOT / "dist" / "release-assets")
    parser.add_argument("--rids", nargs="*", default=list(DEFAULT_RIDS))
    args = parser.parse_args()

    args.out_dir.mkdir(parents=True, exist_ok=True)
    checksums: list[str] = []
    for rid in args.rids:
        rid_dir = ROOT / "dist" / "native" / rid
        if not rid_dir.exists():
            raise SystemExit(f"Missing native RID directory: {rid_dir}")
        suffix = "zip" if rid.startswith("win-") else "tar.gz"
        out = args.out_dir / f"word-ai-openxml-{args.version}-{rid}.{suffix}"
        if out.exists():
            out.unlink()
        if suffix == "zip":
            make_zip(rid_dir, out)
        else:
            make_tar_gz(rid_dir, out)
        checksums.append(f"{sha256(out)}  {out.name}")

    checksum_path = args.out_dir / f"word-ai-openxml-{args.version}-checksums.sha256"
    checksum_path.write_text("\n".join(checksums) + "\n", encoding="utf-8")
    for line in checksums:
        print(line)
    print(f"{sha256(checksum_path)}  {checksum_path.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
