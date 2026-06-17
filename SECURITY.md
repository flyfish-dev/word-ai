# Security Policy

| Language | Preview |
| --- | --- |
| [English](SECURITY.md) | Supported use, vulnerability reporting, and the local-first Word AI security model. |
| [中文](SECURITY.zh-CN.md) | Word AI 支持场景、漏洞报告方式和本地优先安全边界。 |

## Supported Use

Word AI is a local-first developer tool. The default stdio MCP server is intended to run on the same machine as the agent. The local Office bridge is intended for localhost taskpane workflows only.

## Reporting Vulnerabilities

Please report security issues privately to the maintainers of `flyfish-dev/word-ai`. Do not disclose exploitable details in public issues before a fix is available.

## Security Model

- All DOCX paths are scoped to the configured root directory or explicitly allowed roots.
- File-level writes are constrained to PatchSet operations.
- Live Word session writes are constrained to Office.js content-control PatchSet operations.
- Default writes create a new DOCX instead of overwriting the source.
- Office bridge write endpoints require a local token.
- Live session snapshots, command results, audits, and rollback PatchSets are stored locally under `.wordai/sessions`.
- The bridge CORS policy is limited to localhost development origins.
- The project does not require macros or arbitrary shell execution.

See [docs/SECURITY.en.md](docs/SECURITY.en.md) for the detailed design.
