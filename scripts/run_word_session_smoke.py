from __future__ import annotations

import shutil
from pathlib import Path

from word_ai_mcp.server import WordAiMcpServer
from word_ai_mcp.session_store import claim_pending_commands, complete_command, register_session

ROOT = Path(__file__).resolve().parent.parent
SESSION_DIR = ROOT / ".wordai" / "sessions"


def main() -> int:
    shutil.rmtree(SESSION_DIR, ignore_errors=True)
    session = register_session(
        ROOT,
        {
            "client": {"name": "smoke-taskpane"},
            "document": {"host": "Word", "url": "mock://open-document"},
            "capabilities": {"live_session": True},
            "content_controls": [
                {
                    "id": 1,
                    "tag": "WORD-AI:SRS:1.0:overview",
                    "title": "Overview",
                    "text": "Original text",
                    "text_sha256": "hash-old",
                }
            ],
        },
    )
    sid = session["session_id"]
    server = WordAiMcpServer(root=str(ROOT), allow_write=True)

    listed = server.tool_word_session_list({})
    assert listed["count"] == 1, listed
    snapshot = server.tool_word_session_snapshot({"session_id": sid})
    assert snapshot["content_control_count"] == 1, snapshot

    refresh = server.tool_word_session_refresh({"session_id": sid, "wait": False})
    pending = claim_pending_commands(ROOT, sid)
    assert pending and pending[0]["command_id"] == refresh["command_id"]
    complete_command(ROOT, sid, refresh["command_id"], result={"ok": True, "count": 1})
    status = server.tool_word_session_command_status({"command_id": refresh["command_id"]})
    assert status["status"] == "succeeded", status

    patchset = {
        "schema_version": "2.0",
        "strict": True,
        "guard": {"require_preconditions": True, "allow_overwrite": False},
        "operations": [
            {
                "op": "replace_content_control_text",
                "tag": "WORD-AI:SRS:1.0:overview",
                "expected_old_sha256": "hash-old",
                "text": "New text",
            }
        ],
    }
    apply = server.tool_word_session_apply_patchset({"session_id": sid, "patchset": patchset, "wait": False})
    pending = claim_pending_commands(ROOT, sid)
    assert pending and pending[0]["type"] == "apply_patchset"
    rollback_patchset = {
        "schema_version": "2.0",
        "strict": True,
        "guard": {"require_preconditions": True, "allow_overwrite": False},
        "operations": [
            {
                "op": "replace_content_control_text",
                "tag": "WORD-AI:SRS:1.0:overview",
                "expected_old_sha256": "hash-new",
                "text": "Original text",
            }
        ],
    }
    complete_command(
        ROOT,
        sid,
        apply["command_id"],
        result={"ok": True, "audit": {"rollback_patchset": rollback_patchset}},
    )
    rollback = server.tool_word_session_rollback({"session_id": sid, "command_id": apply["command_id"], "wait": False})
    pending = claim_pending_commands(ROOT, sid)
    assert pending and pending[0]["command_id"] == rollback["command_id"]
    assert pending[0]["payload"]["patchset"] == rollback_patchset

    print("word session smoke passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
