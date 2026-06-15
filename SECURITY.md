# Security Policy / 安全策略

## Supported Use

Word AI is a local-first developer tool. The default stdio MCP server is intended to run on the same machine as the agent. The local Office bridge is intended for localhost taskpane workflows only.

## Reporting Vulnerabilities

Please report security issues privately to the maintainers of `flyfish-dev/word-ai`. Do not disclose exploitable details in public issues before a fix is available.

## Security Model

- All DOCX paths are scoped to the configured root directory.
- File-level writes are constrained to PatchSet operations.
- Default writes create a new DOCX instead of overwriting the source.
- Office bridge write endpoints require a local token.
- The bridge CORS policy is limited to localhost development origins.
- The project does not require macros or arbitrary shell execution.

See [docs/SECURITY.md](docs/SECURITY.md) for the detailed design.

## 中文说明

Word AI 是本地优先的开发工具。默认 MCP Server 通过 stdio 与本机 Agent 通信；Office bridge 仅用于 localhost taskpane 工作流。

安全问题请私下报告给 `flyfish-dev/word-ai` 维护者。修复发布前，请不要在公开 issue 中披露可利用细节。

核心安全边界：

- 所有 DOCX 路径限制在配置的 root 内。
- 文件级写入只能通过 PatchSet。
- 默认输出新 DOCX，不覆盖源文件。
- Office bridge 写入口要求本地 token。
- CORS 限制为 localhost 开发源。
- 不需要宏或任意 shell 执行能力。

