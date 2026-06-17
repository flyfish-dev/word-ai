from __future__ import annotations

import argparse
import json
import os
import secrets
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from .server import WordAiMcpServer
from .session_store import claim_pending_commands, complete_command, list_sessions, register_session, update_session

WRITE_TOOLS = {
    "docx_export_plain_text",
    "docx_export_table_csv",
    "docx_table_to_csv",
    "docx_preflight_patchset",
    "docx_dry_run_patchset",
    "docx_write_index",
    "docx_backup",
    "docx_restore_backup",
    "docx_rollback",
    "docx_apply_patchset",
    "word_session_apply_patchset",
    "word_session_wrap_selection",
    "word_session_rollback",
    "officecli_view_screenshot",
}

CONTENT_CONTROL_OPS = {
    "replace_content_control_text",
    "append_content_control_text",
    "prepend_content_control_text",
    "replace_text_in_content_control",
}


def _load_or_create_token(root: Path) -> str:
    token_dir = root / ".wordai"
    token_path = token_dir / "bridge.token"
    if token_path.exists():
        return token_path.read_text(encoding="utf-8").strip()
    token_dir.mkdir(parents=True, exist_ok=True)
    token = secrets.token_urlsafe(32)
    token_path.write_text(token + "\n", encoding="utf-8")
    try:
        token_path.chmod(0o600)
    except OSError:
        pass
    return token


def _json_text(payload: Any) -> str:
    return json.dumps(payload, ensure_ascii=False, indent=2)


def _tool_content(result: Any) -> Any:
    content = ((result or {}).get("result") or {}).get("content") or []
    if not content:
        return result
    text = content[0].get("text", "")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


def _default_output_path(source_docx: str) -> str:
    source = Path(source_docx)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    out_dir = source.parent / ".wordai" / "outputs"
    return str(out_dir / f"{source.stem}.{stamp}.docx")


def make_handler(root: str, allow_write: bool, token: str, allowed_roots: list[str] | None = None):
    server = WordAiMcpServer(root=root, allow_write=allow_write, allowed_roots=allowed_roots)
    allowed_origins = {
        "http://localhost:3000",
        "https://localhost:3000",
        "http://127.0.0.1:3000",
        "https://127.0.0.1:3000",
        "http://localhost:3100",
        "https://localhost:3100",
        "http://127.0.0.1:3100",
        "https://127.0.0.1:3100",
        "http://localhost:5173",
        "https://localhost:5173",
        "null",
    }

    class Handler(BaseHTTPRequestHandler):
        server_version = "WordAiMcpHTTP/0.8.1"

        def log_message(self, fmt: str, *args: Any) -> None:
            print("%s - - [%s] %s" % (self.client_address[0], self.log_date_time_string(), fmt % args))

        def _origin_allowed(self) -> bool:
            origin = self.headers.get("Origin")
            if not origin:
                return True
            if origin in allowed_origins:
                return True
            return origin.startswith("http://localhost:") or origin.startswith("https://localhost:")

        def _send_json(self, status: int, payload: Any) -> None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self._send_cors_headers()
            self.end_headers()
            self.wfile.write(data)

        def _send_cors_headers(self) -> None:
            origin = self.headers.get("Origin")
            if origin and self._origin_allowed():
                self.send_header("Access-Control-Allow-Origin", origin)
            elif not origin:
                self.send_header("Access-Control-Allow-Origin", "https://localhost:3100")
            self.send_header("Vary", "Origin")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Word-AI-Token, Authorization")
            self.send_header("Access-Control-Allow-Private-Network", "true")

        def _read_json_body(self) -> dict[str, Any]:
            length = int(self.headers.get("content-length", "0"))
            if length <= 0:
                return {}
            return json.loads(self.rfile.read(length).decode("utf-8"))

        def _token_ok(self) -> bool:
            provided = self.headers.get("X-Word-AI-Token") or ""
            auth = self.headers.get("Authorization") or ""
            if auth.lower().startswith("bearer "):
                provided = auth[7:].strip()
            return secrets.compare_digest(provided, token)

        def _require_token(self) -> bool:
            if self._token_ok():
                return True
            self._send_json(401, {"ok": False, "error": "missing or invalid X-Word-AI-Token"})
            return False

        def _call_tool(self, name: str, arguments: dict[str, Any]) -> Any:
            req = {"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": name, "arguments": arguments}}
            resp = server.handle(req)
            if resp and resp.get("error"):
                raise ValueError(resp["error"].get("message", "tool call failed"))
            return _tool_content(resp)

        def _read_docx_bundle(self, docx_path: str) -> dict[str, Any]:
            return {
                "inspect": self._call_tool("docx_inspect", {"docx_path": docx_path, "include_text": False}),
                "health": self._call_tool("docx_health_check", {"docx_path": docx_path, "max_items": 50}),
                "content_controls": self._call_tool("docx_list_content_controls", {"docx_path": docx_path, "max_preview": 500}),
                "tables": self._call_tool("docx_list_tables", {"docx_path": docx_path, "max_cell_chars": 120}),
                "fields": self._call_tool("docx_list_fields", {"docx_path": docx_path, "max_preview": 240}),
                "comments": self._call_tool("docx_list_comments", {"docx_path": docx_path, "max_preview": 240}),
                "revisions": self._call_tool("docx_list_revisions", {"docx_path": docx_path, "max_preview": 240}),
            }

        def _build_content_control_patchset(self, body: dict[str, Any]) -> dict[str, Any]:
            docx_path = str(body["docx_path"])
            tag = str(body["tag"]).strip()
            op = str(body.get("operation") or "replace_content_control_text")
            if op not in CONTENT_CONTROL_OPS:
                raise ValueError(f"Unsupported Office bridge operation: {op}")
            inspect = self._call_tool("docx_inspect", {"docx_path": docx_path, "include_text": False})
            current = self._call_tool("docx_read_content_control", {"docx_path": docx_path, "tag": tag})
            operation: dict[str, Any] = {
                "op": op,
                "tag": tag,
                "expected_old_sha256": current["text_sha256"],
                "preserve_style": True,
                "allow_complex_content": False,
            }
            if op == "replace_text_in_content_control":
                find = str(body.get("find") or "")
                if not find:
                    raise ValueError("replace_text_in_content_control requires find")
                operation["find"] = find
                operation["replace"] = str(body.get("replace") or "")
            else:
                operation["text"] = str(body.get("text") or "")
            patchset = {
                "schema_version": "2.0",
                "strict": True,
                "source_sha256": inspect["sha256"],
                "reason": str(body.get("reason") or f"Office.js closed-loop update for {tag}"),
                "guard": {"require_preconditions": True, "allow_overwrite": False},
                "operations": [operation],
            }
            return {"patchset": patchset, "source": {"inspect": inspect, "content_control": current}}

        def _assess_patchset(self, body: dict[str, Any]) -> dict[str, Any]:
            docx_path = str(body["docx_path"])
            patchset = body["patchset"]
            return {
                "health": self._call_tool("docx_health_check", {"docx_path": docx_path, "max_items": 50}),
                "assess": self._call_tool("docx_assess_patchset", {"docx_path": docx_path, "patchset": patchset}),
            }

        def _preview_patchset(self, body: dict[str, Any]) -> dict[str, Any]:
            docx_path = str(body["docx_path"])
            patchset = body["patchset"]
            assessed = self._assess_patchset(body)
            assess = assessed["assess"]
            if assess.get("ok") is False:
                return {**assessed, "dry_run": None, "ok": False}
            dry_run = self._call_tool(
                "docx_dry_run_patchset",
                {"docx_path": docx_path, "patchset": patchset, "keep_output": bool(body.get("keep_output", False))},
            )
            return {**assessed, "dry_run": dry_run, "ok": bool((dry_run.get("validation") or {}).get("ok", False))}

        def _apply_patchset(self, body: dict[str, Any]) -> dict[str, Any]:
            docx_path = str(body["docx_path"])
            patchset = body["patchset"]
            output_path = str(body.get("output_path") or _default_output_path(server._resolve_path(docx_path)))
            preview = self._preview_patchset({"docx_path": docx_path, "patchset": patchset, "keep_output": False})
            if not preview.get("ok"):
                return {"ok": False, "stage": "dry_run", **preview}
            backup = self._call_tool("docx_backup", {"docx_path": docx_path})
            apply_result = self._call_tool(
                "docx_apply_patchset",
                {"docx_path": docx_path, "output_path": output_path, "patchset": patchset},
            )
            validation = self._call_tool("docx_validate", {"source_docx": docx_path, "target_docx": output_path, "strict": True})
            diff = self._call_tool("docx_text_diff", {"source_docx": docx_path, "target_docx": output_path, "context": 2})
            return {
                "ok": bool(validation.get("ok") and (apply_result.get("validation") or {}).get("ok", True)),
                "stage": "complete",
                "preview": preview,
                "backup": backup,
                "apply": apply_result,
                "validation": validation,
                "diff": diff,
                "output_path": output_path,
                "audit_path": str(Path(output_path).with_suffix(".audit.json")),
            }

        def do_OPTIONS(self) -> None:  # noqa: N802
            if not self._origin_allowed():
                self._send_json(403, {"ok": False, "error": "origin not allowed"})
                return
            self.send_response(204)
            self._send_cors_headers()
            self.end_headers()

        def do_GET(self) -> None:  # noqa: N802
            path = urlparse(self.path).path
            if not self._origin_allowed():
                self._send_json(403, {"ok": False, "error": "origin not allowed"})
                return
            if path in ("/health", "/"):
                self._send_json(
                    200,
                    {
                        "ok": True,
                        "name": "word-ai-mcp",
                        "transport": "basic-jsonrpc-http",
                        "office_bridge": True,
                        "auth_required": True,
                        "allow_write": allow_write,
                        "allowed_roots": [str(p) for p in server.allowed_roots],
                    },
                )
            elif path == "/office/capabilities":
                self._send_json(
                    200,
                    {
                        "ok": True,
                        "name": "word-ai-office-bridge",
                        "allow_write": allow_write,
                        "auth_required": True,
                        "endpoints": [
                            "/office/read",
                            "/office/build-patchset",
                            "/office/assess-patchset",
                            "/office/preview-patchset",
                            "/office/apply-patchset",
                            "/office/session/register",
                            "/office/session/heartbeat",
                            "/office/session/poll",
                            "/office/session/result",
                            "/office/session/list",
                        ],
                    },
                )
            else:
                self._send_json(404, {"ok": False, "error": "not found"})

        def do_POST(self) -> None:  # noqa: N802
            path = urlparse(self.path).path
            if not self._origin_allowed():
                self._send_json(403, {"ok": False, "error": "origin not allowed"})
                return
            try:
                body = self._read_json_body()
                if path in ("/mcp", "/"):
                    params = body.get("params") or {}
                    name = params.get("name")
                    if body.get("method") == "tools/call" and name in WRITE_TOOLS and not self._require_token():
                        return
                    resp = server.handle(body)
                    self._send_json(200, resp or {})
                    return
                if path.startswith("/office/") and not self._require_token():
                    return
                if path == "/office/read":
                    self._send_json(200, {"ok": True, **self._read_docx_bundle(str(body["docx_path"]))})
                elif path == "/office/build-patchset":
                    self._send_json(200, {"ok": True, **self._build_content_control_patchset(body)})
                elif path == "/office/assess-patchset":
                    self._send_json(200, {"ok": True, **self._assess_patchset(body)})
                elif path == "/office/preview-patchset":
                    self._send_json(200, self._preview_patchset(body))
                elif path == "/office/apply-patchset":
                    self._send_json(200, self._apply_patchset(body))
                elif path == "/office/session/register":
                    self._send_json(200, {"ok": True, "session": register_session(root, body)})
                elif path == "/office/session/heartbeat":
                    self._send_json(200, {"ok": True, "session": update_session(root, str(body["session_id"]), body)})
                elif path == "/office/session/poll":
                    commands = claim_pending_commands(root, str(body["session_id"]), int(body.get("limit", 5)))
                    self._send_json(200, {"ok": True, "commands": commands, "count": len(commands)})
                elif path == "/office/session/result":
                    command = complete_command(
                        root,
                        str(body["session_id"]),
                        str(body["command_id"]),
                        result=body.get("result"),
                        error=body.get("error"),
                    )
                    snapshot = body.get("snapshot")
                    if isinstance(snapshot, dict):
                        update_session(root, str(body["session_id"]), snapshot)
                    self._send_json(200, {"ok": True, "command": command})
                elif path == "/office/session/list":
                    sessions = list_sessions(root, bool(body.get("include_inactive", False)))
                    self._send_json(200, {"ok": True, "sessions": sessions, "count": len(sessions)})
                else:
                    self._send_json(404, {"ok": False, "error": "not found"})
            except KeyError as exc:
                self._send_json(400, {"ok": False, "error": f"missing required field: {exc}"})
            except Exception as exc:
                self._send_json(500, {"ok": False, "error": str(exc)})

    return Handler


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Local JSON-RPC HTTP and Office.js bridge. Use a full MCP Streamable HTTP implementation for production remote MCP."
    )
    parser.add_argument("--root", default=os.getcwd())
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--read-only", action="store_true")
    parser.add_argument("--token", default=os.environ.get("WORD_AI_MCP_TOKEN"))
    parser.add_argument("--allow-root", action="append", default=[], help="Additional directory allowed for DOCX input/output. Repeatable. WORD_AI_ALLOWED_ROOTS also works.")
    args = parser.parse_args(argv)

    root = Path(args.root).resolve()
    token = args.token or _load_or_create_token(root)
    httpd = ThreadingHTTPServer((args.host, args.port), make_handler(str(root), allow_write=not args.read_only, token=token, allowed_roots=args.allow_root))
    print(f"word-ai-mcp HTTP server listening at http://{args.host}:{args.port}/mcp")
    print(f"office bridge listening at http://{args.host}:{args.port}/office")
    print(f"office bridge token: {token}")
    httpd.serve_forever()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
