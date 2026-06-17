# 安全策略

| 文档 | 预览 |
| --- | --- |
| [中文](SECURITY.zh-CN.md) | Word AI 支持场景、漏洞报告方式和本地优先安全边界。 |
| [English](SECURITY.md) | Supported use, vulnerability reporting, and the local-first Word AI security model. |

## 支持场景

Word AI 是本地优先的开发工具。默认 stdio MCP Server 运行在 Agent 所在的同一台机器上；本地 Office bridge 仅面向 localhost taskpane 工作流。

## 漏洞报告

请将安全问题私下报告给 `flyfish-dev/word-ai` 维护者。修复发布前，请不要在公开 issue 中披露可利用细节。

## 安全模型

- 所有 DOCX 路径限制在配置的 root 或明确允许的 allow-root 内。
- 文件级写入只能通过 PatchSet 操作。
- 打开的 Word 会话写入只能通过 Office.js 内容控件 PatchSet 操作。
- 默认输出新 DOCX，不覆盖源文件。
- Office bridge 写入口要求本地 token。
- Live session 快照、命令结果、审计和 rollback PatchSet 保存在本地 `.wordai/sessions`。
- CORS 限制为 localhost 开发源。
- 项目不需要宏或任意 shell 执行能力。

详细设计见 [docs/SECURITY.md](docs/SECURITY.md)。
