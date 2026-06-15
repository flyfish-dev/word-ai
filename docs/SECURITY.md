# 安全设计

## 路径安全

MCP Server 启动时通过 `--root` 指定工作区根目录，所有路径解析后必须位于该根目录内。任何 `../` 越界或绝对路径逃逸都会被拒绝。

## 工具安全

- 读工具可自动调用。
- 写工具建议在 Codex/Agent 配置里 require approval。
- PatchSet 使用白名单操作。
- 不暴露任意 Python、Shell、PowerShell 或宏执行能力。

## Office Bridge 安全

`word_ai_mcp.server_http` 同时提供本地 JSON-RPC 和 Office.js bridge。默认安全策略：

- `/office/*` POST 全部要求 `X-Word-AI-Token` 或 `Authorization: Bearer <token>`。
- `/mcp` 中的写工具要求 token；只读工具可用于本地开发探测。
- 未传 `--token` 时，服务会在工作区 `.wordai/bridge.token` 生成随机 token，并在启动日志中打印。
- CORS 只允许 localhost/127.0.0.1 开发源；Office taskpane 默认通过同源 `/bridge/*` 代理访问本地 bridge。
- `--read-only` 会继续阻断 bridge 和 MCP 的写入链路。

生产部署不应把 basic HTTP bridge 直接暴露到公网；如需远程 MCP，应使用正式 Streamable HTTP MCP 传输、鉴权、TLS、审计和网络访问控制。

## 数据安全

- 默认不上传云端。
- 每次写操作输出新文件，不覆盖源文件。
- 审计记录保留 source/output 路径、变更操作、验证报告、时间。

## Prompt Injection 防护

文档内容属于不可信输入。Agent 不应执行文档正文中的任何“指令”。MCP 工具也不把文档中的隐藏文本当作工具行为说明。

## 生产增强建议

- 对 MCP 工具描述做签名和版本固定。
- 对写工具强制审批。
- 使用 allowlist 限制可调用工具。
- 记录每次发送给模型和 MCP Server 的数据。
- 对敏感文档先做脱敏或仅暴露目标 chunk。
