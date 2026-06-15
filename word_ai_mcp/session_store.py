from __future__ import annotations

import json
import os
import secrets
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

JSON = dict[str, Any]

SESSION_TTL_SECONDS = 45


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _safe_id(value: str) -> str:
    safe = "".join(ch for ch in value if ch.isalnum() or ch in "-_")[:96]
    if not safe:
        raise ValueError("Invalid Word session or command id")
    return safe


def _store_dir(root: str | Path) -> Path:
    path = Path(root).resolve() / ".wordai" / "sessions"
    path.mkdir(parents=True, exist_ok=True)
    (path / "commands").mkdir(parents=True, exist_ok=True)
    return path


def _session_path(root: str | Path, session_id: str) -> Path:
    return _store_dir(root) / f"{_safe_id(session_id)}.json"


def _command_path(root: str | Path, command_id: str) -> Path:
    return _store_dir(root) / "commands" / f"{_safe_id(command_id)}.json"


def _write_json(path: Path, payload: JSON) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + f".{os.getpid()}.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def _read_json(path: Path) -> JSON:
    return json.loads(path.read_text(encoding="utf-8"))


def _age_seconds(iso: str | None) -> float | None:
    if not iso:
        return None
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return max(0.0, time.time() - dt.timestamp())
    except ValueError:
        return None


def _decorate_session(session: JSON) -> JSON:
    age = _age_seconds(session.get("updated_at"))
    return {
        **session,
        "age_seconds": age,
        "active": age is not None and age <= SESSION_TTL_SECONDS,
    }


def register_session(root: str | Path, payload: JSON) -> JSON:
    requested = str(payload.get("session_id") or "").strip()
    session_id = _safe_id(requested) if requested else secrets.token_urlsafe(12)
    now = utc_now()
    existing: JSON = {}
    path = _session_path(root, session_id)
    if path.exists():
        existing = _read_json(path)
    session = {
        **existing,
        "session_id": session_id,
        "created_at": existing.get("created_at") or now,
        "updated_at": now,
        "status": "active",
        "client": payload.get("client") or existing.get("client") or {},
        "document": payload.get("document") or existing.get("document") or {},
        "capabilities": payload.get("capabilities") or existing.get("capabilities") or {},
        "content_controls": payload.get("content_controls") or existing.get("content_controls") or [],
        "selection": payload.get("selection") or existing.get("selection"),
        "last_error": None,
    }
    _write_json(path, session)
    return _decorate_session(session)


def update_session(root: str | Path, session_id: str, payload: JSON) -> JSON:
    path = _session_path(root, session_id)
    if not path.exists():
        raise FileNotFoundError(f"Word session not found: {session_id}")
    session = _read_json(path)
    now = utc_now()
    for key in ["client", "document", "capabilities", "content_controls", "selection", "status", "last_error"]:
        if key in payload:
            session[key] = payload[key]
    session["updated_at"] = now
    _write_json(path, session)
    return _decorate_session(session)


def list_sessions(root: str | Path, include_inactive: bool = False) -> list[JSON]:
    sessions: list[JSON] = []
    for path in sorted(_store_dir(root).glob("*.json")):
        try:
            session = _decorate_session(_read_json(path))
        except Exception:
            continue
        if include_inactive or session.get("active"):
            sessions.append(session)
    return sessions


def get_session(root: str | Path, session_id: str) -> JSON:
    path = _session_path(root, session_id)
    if not path.exists():
        raise FileNotFoundError(f"Word session not found: {session_id}")
    return _decorate_session(_read_json(path))


def enqueue_command(root: str | Path, session_id: str, command_type: str, payload: JSON | None = None) -> JSON:
    session = get_session(root, session_id)
    if not session.get("active"):
        raise TimeoutError(f"Word session is inactive: {session_id}")
    now = utc_now()
    command_id = secrets.token_urlsafe(16)
    command = {
        "command_id": command_id,
        "session_id": session_id,
        "type": command_type,
        "payload": payload or {},
        "status": "queued",
        "created_at": now,
        "updated_at": now,
        "dispatched_at": None,
        "completed_at": None,
        "result": None,
        "error": None,
    }
    _write_json(_command_path(root, command_id), command)
    return command


def get_command(root: str | Path, command_id: str) -> JSON:
    path = _command_path(root, command_id)
    if not path.exists():
        raise FileNotFoundError(f"Word session command not found: {command_id}")
    return _read_json(path)


def claim_pending_commands(root: str | Path, session_id: str, limit: int = 5) -> list[JSON]:
    claimed: list[JSON] = []
    now = utc_now()
    for path in sorted((_store_dir(root) / "commands").glob("*.json")):
        if len(claimed) >= limit:
            break
        try:
            command = _read_json(path)
        except Exception:
            continue
        if command.get("session_id") != session_id or command.get("status") != "queued":
            continue
        command["status"] = "dispatched"
        command["dispatched_at"] = now
        command["updated_at"] = now
        _write_json(path, command)
        claimed.append(command)
    return claimed


def complete_command(root: str | Path, session_id: str, command_id: str, *, result: Any = None, error: Any = None) -> JSON:
    command = get_command(root, command_id)
    if command.get("session_id") != session_id:
        raise PermissionError("Command does not belong to this Word session")
    now = utc_now()
    command["status"] = "failed" if error else "succeeded"
    command["completed_at"] = now
    command["updated_at"] = now
    command["result"] = result
    command["error"] = error
    _write_json(_command_path(root, command_id), command)
    return command


def wait_for_command(root: str | Path, command_id: str, timeout_seconds: float = 20.0, poll_interval: float = 0.25) -> JSON:
    deadline = time.time() + max(0.0, timeout_seconds)
    while True:
        command = get_command(root, command_id)
        if command.get("status") in {"succeeded", "failed"}:
            return command
        if time.time() >= deadline:
            return {
                **command,
                "status": "timeout",
                "error": f"Timed out waiting for Word session command {command_id}",
            }
        time.sleep(poll_interval)


def session_summary(root: str | Path, session_id: str) -> JSON:
    session = get_session(root, session_id)
    controls = session.get("content_controls") or []
    return {
        "session_id": session["session_id"],
        "active": session["active"],
        "updated_at": session["updated_at"],
        "age_seconds": session["age_seconds"],
        "document": session.get("document") or {},
        "content_control_count": len(controls),
        "content_controls": controls,
        "capabilities": session.get("capabilities") or {},
    }
