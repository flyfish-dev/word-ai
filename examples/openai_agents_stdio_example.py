"""OpenAI Agents SDK example for the local stdio MCP server.

Install the Agents SDK in your own environment, then run this file from the
repository root. This example intentionally requires approval for write tools.
"""
from __future__ import annotations

import asyncio
from pathlib import Path

from agents import Agent, Runner
from agents.mcp import MCPServerStdio, create_static_tool_filter

ROOT = Path(__file__).resolve().parent.parent
SAFE_READ_TOOLS = [
    "docx_inspect",
    "docx_list_anchors",
    "docx_read_content_control",
    "docx_write_index",
    "docx_text_diff",
    "docx_validate",
    "word_session_list",
    "word_session_snapshot",
    "word_session_read_content_control",
    "word_session_preview_patchset",
]
WRITE_TOOLS = ["docx_backup", "docx_apply_patchset", "word_session_apply_patchset", "word_session_wrap_selection", "word_session_rollback"]


async def main() -> None:
    async with MCPServerStdio(
        name="word_ai",
        params={
            "command": "python",
            "args": ["-m", "word_ai_mcp.server", "--root", str(ROOT)],
        },
        tool_filter=create_static_tool_filter(allowed_tool_names=SAFE_READ_TOOLS + WRITE_TOOLS),
        require_approval={"always": {"tool_names": WRITE_TOOLS}, "never": {"tool_names": SAFE_READ_TOOLS}},
    ) as server:
        agent = Agent(
            name="Word delivery editor",
            instructions=(
                "You edit DOCX delivery documents through MCP only. "
                "Inspect anchors first, prefer content controls, produce PatchSet JSON, "
                "then use validate and diff. Never regenerate the full DOCX."
            ),
            mcp_servers=[server],
        )
        result = await Runner.run(
            agent,
            "Inspect examples/sample_contract.docx and list the AI-editable content control anchors.",
        )
        print(result.final_output)


if __name__ == "__main__":
    asyncio.run(main())
