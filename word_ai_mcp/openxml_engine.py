from __future__ import annotations

import json
import os
import platform
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .patchset import normalize_patchset
from .resources import runtime_root


JSON = dict[str, Any]
ENGINE_CHOICES = {"auto", "dotnet", "python"}
SUPPORTED_NATIVE_RIDS = {
    "osx-arm64",
    "osx-x64",
    "linux-x64",
    "linux-arm64",
    "linux-musl-x64",
    "linux-musl-arm64",
    "win-x64",
    "win-arm64",
}


class EngineUnavailable(RuntimeError):
    pass


def package_root() -> Path:
    return runtime_root()


def normalize_engine(value: str | None = None) -> str:
    engine = (value or os.environ.get("WORD_AI_ENGINE") or "auto").strip().lower()
    if engine not in ENGINE_CHOICES:
        raise ValueError("WORD_AI_ENGINE must be one of: auto, dotnet, python")
    return engine


def default_output_path(source_docx: str | Path) -> str:
    source = Path(source_docx)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return str(source.with_name(f"{source.stem}.wordai.{stamp}.docx"))


def _runtime_id() -> str | None:
    machine = platform.machine().lower()
    arch = "arm64" if machine in {"arm64", "aarch64"} else "x64" if machine in {"x86_64", "amd64"} else None
    if arch is None:
        return None
    if os.name == "nt":
        return f"win-{arch}"
    if sys_platform().startswith("darwin"):
        return f"osx-{arch}"
    if sys_platform().startswith("linux"):
        return f"{_linux_rid_prefix()}-{arch}"
    return None


def _linux_rid_prefix() -> str:
    libc_name = (platform.libc_ver()[0] or "").lower()
    if "musl" in libc_name:
        return "linux-musl"
    return "linux"


def _runtime_id_candidates() -> list[str]:
    configured = os.environ.get("WORD_AI_DOTNET_RID")
    configured_text = configured.strip() if configured else ""
    configured_rid = configured_text if configured_text in SUPPORTED_NATIVE_RIDS else None
    candidates: list[str] = [configured_rid] if configured_rid else []
    rid = _runtime_id()
    if rid:
        candidates.append(rid)
        if rid.startswith("linux-musl-"):
            candidates.append(rid.replace("linux-musl-", "linux-", 1))
        elif rid.startswith("linux-"):
            candidates.append(rid.replace("linux-", "linux-musl-", 1))
    deduped: list[str] = []
    for candidate in candidates:
        if candidate not in deduped:
            deduped.append(candidate)
    return deduped


def sys_platform() -> str:
    # Isolated for tests and for keeping platform normalization in one place.
    import sys

    return sys.platform


def _dotnet_command() -> str | None:
    configured = os.environ.get("WORD_AI_DOTNET")
    if configured:
        return configured
    homebrew_dotnet = Path("/opt/homebrew/opt/dotnet@8/libexec/dotnet")
    if homebrew_dotnet.exists():
        return str(homebrew_dotnet)
    return shutil.which("dotnet")


def _native_name(rid: str | None = None) -> str:
    if (rid or "").startswith("win-") or os.name == "nt":
        return "WordAi.OpenXml.exe"
    return "WordAi.OpenXml"


def _candidate_native_executables(root: Path) -> list[Path]:
    configured = os.environ.get("WORD_AI_DOTNET_EXE")
    candidates: list[Path] = [Path(configured).expanduser()] if configured else []
    configured_native_root = os.environ.get("WORD_AI_DOTNET_NATIVE_DIR")
    native_roots = []
    if configured_native_root:
        native_roots.append(Path(configured_native_root).expanduser())
    native_roots.extend([root / "native", root / "dist" / "native"])
    for rid in _runtime_id_candidates():
        name = _native_name(rid)
        for native_root in native_roots:
            candidates.append(native_root / rid / name)
        candidates.extend(
            [
                root / "dotnet" / "WordAi.OpenXml" / "bin" / "Release" / "net8.0" / rid / "publish" / name,
            ]
        )
    return candidates


def _candidate_dlls(root: Path) -> list[Path]:
    configured = os.environ.get("WORD_AI_DOTNET_DLL")
    candidates: list[Path] = [Path(configured).expanduser()] if configured else []
    candidates.append(root / "dotnet" / "WordAi.OpenXml" / "bin" / "Release" / "net8.0" / "WordAi.OpenXml.dll")
    return candidates


def dotnet_status(root: str | Path | None = None) -> JSON:
    root_path = Path(root).resolve() if root else package_root()
    for exe in _candidate_native_executables(root_path):
        if exe.exists() and os.access(exe, os.X_OK):
            return {
                "available": True,
                "engine": "dotnet",
                "mode": "native",
                "command": [str(exe)],
                "root": str(root_path),
                "runtime_id": _runtime_id(),
                "runtime_id_candidates": _runtime_id_candidates(),
                "native_path": str(exe),
            }

    dotnet = _dotnet_command()
    if dotnet:
        for dll in _candidate_dlls(root_path):
            if dll.exists():
                return {
                    "available": True,
                    "engine": "dotnet",
                    "mode": "dll",
                    "command": [dotnet, str(dll)],
                    "root": str(root_path),
                    "runtime_id": _runtime_id(),
                    "runtime_id_candidates": _runtime_id_candidates(),
                }
        project = root_path / "dotnet" / "WordAi.OpenXml" / "WordAi.OpenXml.csproj"
        if project.exists():
            return {
                "available": True,
                "engine": "dotnet",
                "mode": "project",
                "command": [dotnet, "run", "--project", str(project), "-c", "Release", "--"],
                "root": str(root_path),
                "runtime_id": _runtime_id(),
                "runtime_id_candidates": _runtime_id_candidates(),
            }

    return {
        "available": False,
        "engine": "dotnet",
        "mode": None,
        "root": str(root_path),
        "runtime_id": _runtime_id(),
        "runtime_id_candidates": _runtime_id_candidates(),
        "reason": "No WordAi.OpenXml native binary, DLL, or .NET SDK/project backend was found.",
    }


def select_engine(requested: str | None = None, *, root: str | Path | None = None) -> tuple[str, JSON | None]:
    engine = normalize_engine(requested)
    if engine == "python":
        return "python", None
    status = dotnet_status(root)
    if status.get("available"):
        return "dotnet", status
    if engine == "dotnet":
        raise EngineUnavailable(str(status.get("reason") or "The .NET Open XML backend is unavailable."))
    return "python", {"fallback_from": "dotnet", "fallback_reason": status.get("reason"), "dotnet_status": status}


def _env_for_dotnet() -> dict[str, str]:
    env = os.environ.copy()
    homebrew_root = Path("/opt/homebrew/opt/dotnet@8/libexec")
    if homebrew_root.exists() and not env.get("DOTNET_ROOT"):
        env["DOTNET_ROOT"] = str(homebrew_root)
    return env


def _run_dotnet(args: list[str], *, root: str | Path | None = None, timeout_seconds: float | None = None) -> JSON:
    status = dotnet_status(root)
    if not status.get("available"):
        raise EngineUnavailable(str(status.get("reason") or "The .NET Open XML backend is unavailable."))
    cmd = [*status["command"], *args]
    timeout = timeout_seconds or float(os.environ.get("WORD_AI_DOTNET_TIMEOUT_SECONDS", "180"))
    proc = subprocess.run(
        cmd,
        cwd=str(Path(status["root"])),
        env=_env_for_dotnet(),
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=timeout,
    )
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or f".NET backend exited with status {proc.returncode}")
    try:
        payload = json.loads(proc.stdout) if proc.stdout.strip() else {}
    except json.JSONDecodeError as exc:
        raise RuntimeError(f".NET backend returned invalid JSON: {exc}: {proc.stdout[:1000]}") from exc
    payload.setdefault("engine", "dotnet")
    payload.setdefault("engine_detail", {k: v for k, v in status.items() if k not in {"command"}})
    return payload


def _write_json_temp(data: Any, directory: Path, suffix: str) -> Path:
    fd, name = tempfile.mkstemp(prefix="word-ai-", suffix=suffix, dir=directory)
    path = Path(name)
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    return path


def dotnet_assess_patchset(docx_path: str | Path, patchset: JSON, *, root: str | Path | None = None) -> JSON:
    patchset = normalize_patchset(patchset)
    with tempfile.TemporaryDirectory(prefix="word-ai-openxml-") as tmp:
        patch_path = _write_json_temp(patchset, Path(tmp), ".patchset.json")
        return _run_dotnet(["assess", str(docx_path), str(patch_path)], root=root)


def dotnet_dry_run_patchset(docx_path: str | Path, patchset: JSON, keep_output: bool = False, *, root: str | Path | None = None) -> JSON:
    patchset = normalize_patchset(patchset)
    with tempfile.TemporaryDirectory(prefix="word-ai-openxml-") as tmp:
        patch_path = _write_json_temp(patchset, Path(tmp), ".patchset.json")
        return _run_dotnet(["dry-run", str(docx_path), str(patch_path), str(bool(keep_output)).lower()], root=root)


def dotnet_apply_patchset(docx_path: str | Path, patchset: JSON, output_path: str | Path | None = None, *, root: str | Path | None = None) -> JSON:
    patchset = normalize_patchset(patchset)
    target = str(output_path or default_output_path(docx_path))
    with tempfile.TemporaryDirectory(prefix="word-ai-openxml-") as tmp:
        patch_path = _write_json_temp(patchset, Path(tmp), ".patchset.json")
        return _run_dotnet(["apply", str(docx_path), str(patch_path), target], root=root)


def dotnet_validate(source_docx: str | Path, target_docx: str | Path, strict: bool = True, options: JSON | None = None, *, root: str | Path | None = None) -> JSON:
    payload = {"strict": strict, **(options or {})}
    with tempfile.TemporaryDirectory(prefix="word-ai-openxml-") as tmp:
        options_path = _write_json_temp(payload, Path(tmp), ".validation.json")
        return _run_dotnet(["validate", str(source_docx), str(target_docx), str(options_path)], root=root)


def mark_python_result(result: Any, detail: JSON | None = None) -> Any:
    if isinstance(result, dict):
        result.setdefault("engine", "python")
        if detail:
            result.setdefault("engine_detail", detail)
    return result
