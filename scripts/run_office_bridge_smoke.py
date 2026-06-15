from __future__ import annotations

import json
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TOKEN = "office-bridge-smoke-token"


def free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def request(base: str, path: str, payload: dict | None = None, token: str | None = TOKEN, method: str = "POST") -> tuple[int, dict]:
    data = None if payload is None else json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(base + path, data=data, method=method)
    req.add_header("Content-Type", "application/json")
    if token:
        req.add_header("X-Word-AI-Token", token)
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            return resp.status, body
    except urllib.error.HTTPError as exc:
        body = json.loads(exc.read().decode("utf-8"))
        return exc.code, body


def wait_for_server(base: str) -> None:
    deadline = time.time() + 15
    while time.time() < deadline:
        try:
            status, body = request(base, "/health", None, token=None, method="GET")
            if status == 200 and body.get("ok"):
                return
        except Exception:
            pass
        time.sleep(0.2)
    raise RuntimeError("office bridge server did not start")


def main() -> int:
    port = free_port()
    base = f"http://127.0.0.1:{port}"
    proc = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "word_ai_mcp.server_http",
            "--root",
            str(ROOT),
            "--host",
            "127.0.0.1",
            "--port",
            str(port),
            "--token",
            TOKEN,
        ],
        cwd=str(ROOT),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
    )
    try:
        wait_for_server(base)
        docx_path = "examples/sample_contract.docx"
        status, _ = request(base, "/office/read", {"docx_path": docx_path}, token=None)
        assert status == 401, "office read must require token"

        status, read_payload = request(base, "/office/read", {"docx_path": docx_path})
        assert status == 200 and read_payload["ok"]
        assert read_payload["content_controls"]["count"] >= 1

        status, build_payload = request(
            base,
            "/office/build-patchset",
            {
                "docx_path": docx_path,
                "tag": "WORD-AI:SRS:1.0:overview",
                "operation": "replace_content_control_text",
                "text": "[[CC:overview]] Office bridge smoke replacement.",
            },
        )
        assert status == 200 and build_payload["ok"]
        patchset = build_payload["patchset"]
        assert patchset["source_sha256"]
        assert patchset["operations"][0]["expected_old_sha256"]

        status, assess_payload = request(base, "/office/assess-patchset", {"docx_path": docx_path, "patchset": patchset})
        assert status == 200 and assess_payload["assess"]["ok"]

        status, preview_payload = request(base, "/office/preview-patchset", {"docx_path": docx_path, "patchset": patchset})
        assert status == 200 and preview_payload["ok"]
        assert preview_payload["dry_run"]["validation"]["ok"]

        out_dir = ROOT / "examples" / ".wordai" / "bridge-smoke"
        out_dir.mkdir(parents=True, exist_ok=True)
        output_path = out_dir / "sample_contract.bridge.docx"
        audit_path = output_path.with_suffix(".audit.json")
        for path in (output_path, audit_path):
            if path.exists():
                path.unlink()

        status, apply_payload = request(
            base,
            "/office/apply-patchset",
            {"docx_path": docx_path, "patchset": patchset, "output_path": str(output_path)},
        )
        assert status == 200 and apply_payload["ok"], apply_payload
        assert output_path.exists()
        assert audit_path.exists()
        assert apply_payload["validation"]["ok"]
        assert "word/document.xml" in apply_payload["validation"]["metrics"]["changed_parts"]

        status, mcp_write = request(
            base,
            "/mcp",
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "tools/call",
                "params": {"name": "docx_backup", "arguments": {"docx_path": docx_path}},
            },
            token=None,
        )
        assert status == 401, mcp_write

        print("office bridge smoke passed")
        return 0
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()


if __name__ == "__main__":
    raise SystemExit(main())
