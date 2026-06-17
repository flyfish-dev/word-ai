# Security Design

| Language | Preview |
| --- | --- |
| [English](SECURITY.en.md) | Path safety, write boundaries, Office bridge token model, data safety, and production hardening notes. |
| [中文](SECURITY.md) | 路径安全、工具写入边界、Office bridge token、数据安全和生产增强建议。 |

## Path Safety

The MCP server starts with `--root` as the primary workspace root. Relative paths resolve under that directory. To access original documents outside the project, pass `--allow-root <dir>` repeatedly or set `WORD_AI_ALLOWED_ROOTS`. After path resolution, every path must stay inside `--root` or an explicitly allowed root; unauthorized `../` traversal and absolute path escapes are rejected.

The installer-generated Codex config adds common user document folders when they exist: `~/Downloads`, `~/Documents`, and `~/Desktop`. This improves usability for real DOCX files while avoiding arbitrary system-wide file access.

## Tool Safety

- Read tools may be called freely by agents.
- Write tools should require approval in Codex or agent configuration.
- PatchSet uses an operation allowlist.
- Live Word session writes also go through PatchSet: `word_session_apply_patchset` only runs content-control text operations and requires `expected_old_sha256`.
- Word AI does not expose arbitrary Python, shell, PowerShell, or macro execution.

## Office Bridge Security

`word_ai_mcp.server_http` provides both local JSON-RPC and the Office.js bridge. The default security policy is:

- All `/office/*` POST requests require `X-Word-AI-Token` or `Authorization: Bearer <token>`.
- Write tools under `/mcp` require a token; read tools may be used for local development probing.
- When `--token` is omitted, the service creates a random token at `.wordai/bridge.token` and prints it in the startup log.
- Office.js live sessions store open-document snapshots, command queues, results, audits, and rollback PatchSets under `.wordai/sessions`.
- CORS is limited to localhost and 127.0.0.1 development origins. The Office taskpane normally accesses the local bridge through same-origin `/bridge/*` proxy routes.
- `--read-only` continues to block bridge and MCP write paths.

Do not expose the basic HTTP bridge directly to the public internet. For remote MCP deployments, use the formal Streamable HTTP MCP transport, authentication, TLS, audit storage, and network access controls.

## Data Safety

- Word AI does not upload documents by default.
- Each file write creates a new DOCX instead of overwriting the source.
- Audit records include source/output paths, operations, validation reports, and timestamps.
- Live session writes affect the currently open Word process rather than creating a new DOCX file. They return audit and rollback PatchSet data and temporarily store command results under `.wordai/sessions/commands`.

## Prompt Injection Defense

Document content is untrusted input. Agents must not execute instructions embedded in document text. MCP tools also do not treat hidden document text as tool behavior instructions.

## Production Hardening

- Sign and pin MCP tool descriptions.
- Require approval for write tools.
- Use an allowlist for callable tools.
- Log every data exchange sent to the model and MCP server.
- Redact sensitive documents or expose only the target chunk.
