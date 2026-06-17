# 安全设计

| 文档 | 预览 |
| --- | --- |
| [中文](SECURITY.md) | 路径安全、工具写入边界、Office bridge token、数据安全和生产增强建议。 |
| [English](SECURITY.en.md) | Path safety, write boundaries, Office bridge token model, data safety, and production hardening notes. |

## 路径安全

MCP Server 启动时通过 `--root` 指定主工作区根目录；相对路径都基于该目录解析。需要访问项目外的原始文档时，可以重复传入 `--allow-root <dir>`，或设置 `WORD_AI_ALLOWED_ROOTS`。所有路径解析后必须位于 `--root` 或某个 `--allow-root` 目录内；任何未授权的 `../` 越界或绝对路径逃逸都会被拒绝。

安装脚本生成的 Codex 配置默认加入常用用户文档目录（存在时）：`~/Downloads`、`~/Documents`、`~/Desktop`。这提升了处理用户原始 DOCX 路径的可用性，同时仍避免访问系统任意位置。

## 工具安全

- 读工具可自动调用。
- 写工具建议在 Codex/Agent 配置里 require approval。
- PatchSet 使用白名单操作。
- 打开的 Word 会话写入也必须通过 PatchSet：`word_session_apply_patchset` 只执行内容控件文本类操作，并要求 `expected_old_sha256`。
- 不暴露任意 Python、Shell、PowerShell 或宏执行能力。

## Office Bridge 安全

`word_ai_mcp.server_http` 同时提供本地 JSON-RPC 和 Office.js bridge。默认安全策略：

- `/office/*` POST 全部要求 `X-Word-AI-Token` 或 `Authorization: Bearer <token>`。
- `/mcp` 中的写工具要求 token；只读工具可用于本地开发探测。
- 未传 `--token` 时，服务会在工作区 `.wordai/bridge.token` 生成随机 token，并在启动日志中打印。
- Office.js live session 使用 `.wordai/sessions` 保存打开文档快照、命令队列、执行结果、audit 和 rollback PatchSet。
- CORS 只允许 localhost/127.0.0.1 开发源；Office taskpane 默认通过同源 `/bridge/*` 代理访问本地 bridge。
- `--read-only` 会继续阻断 bridge 和 MCP 的写入链路。

生产部署不应把 basic HTTP bridge 直接暴露到公网；如需远程 MCP，应使用正式 Streamable HTTP MCP 传输、鉴权、TLS、审计和网络访问控制。

## 数据安全

- 默认不上传云端。
- 每次写操作输出新文件，不覆盖源文件。
- 审计记录保留 source/output 路径、变更操作、验证报告、时间。
- Live session 写入作用于当前 Word 进程中的打开文档，不生成新 DOCX 文件；它会返回 audit 和 rollback PatchSet，并把命令结果暂存在 `.wordai/sessions/commands`。

## Prompt Injection 防护

文档内容属于不可信输入。Agent 不应执行文档正文中的任何“指令”。MCP 工具也不把文档中的隐藏文本当作工具行为说明。

## 生产增强建议

- 对 MCP 工具描述做签名和版本固定。
- 对写工具强制审批。
- 使用 allowlist 限制可调用工具。
- 记录每次发送给模型和 MCP Server 的数据。
- 对敏感文档先做脱敏或仅暴露目标 chunk。
