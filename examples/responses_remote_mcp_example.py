"""Responses API remote MCP example.

This repository's built-in HTTP adapter is a basic JSON-RPC helper for local development.
For production remote MCP, deploy a proper Streamable HTTP MCP server and put its URL below.
"""
from __future__ import annotations

from openai import OpenAI

client = OpenAI()

resp = client.responses.create(
    model="gpt-5.5",
    tools=[
        {
            "type": "mcp",
            "server_label": "word_ai",
            "server_url": "https://your-word-ai-mcp.example.com/mcp",
            "require_approval": {
                "always": {"tool_names": ["docx_apply_patchset", "docx_backup"]},
                "never": {"tool_names": ["docx_inspect", "docx_list_anchors", "docx_read_content_control", "docx_validate", "docx_text_diff"]},
            },
            "allowed_tools": [
                "docx_inspect",
                "docx_list_anchors",
                "docx_read_content_control",
                "docx_validate",
                "docx_text_diff",
                "docx_backup",
                "docx_apply_patchset",
            ],
        }
    ],
    input="Inspect the sample Word document and summarize editable anchors.",
)

print(resp.output_text)
