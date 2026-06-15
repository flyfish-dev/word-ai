# v0.6 Live Word Session 闭环升级记录

## 核心变化

- MCP tool surface 从 49 个扩展到 58 个，新增 `word_session_*` 工具族。
- Office.js taskpane 可以把当前打开的 Word 文档注册为 live session。
- Codex 可以通过 MCP 读取 active Word session、读取内容控件、预览 PatchSet、应用 PatchSet、回滚上一次 live apply。
- 本地 bridge 新增 session register、heartbeat、poll、result、list 端点。
- 新增 `.wordai/sessions` 文件队列，用于在 Codex MCP server 与 Word taskpane 之间传递命令和结果。

## 安全与事务

- `word_session_apply_patchset` 只支持内容控件文本操作：
  `replace_content_control_text`、`append_content_control_text`、`prepend_content_control_text`、`replace_text_in_content_control`。
- 每个 live operation 必须带 `expected_old_sha256`。
- apply 内部会先运行 live preflight，然后再次校验 hash，再通过 Office.js 写入。
- apply 返回 audit、触达内容控件、validation 和 rollback PatchSet。
- `word_session_rollback` 使用上一次 apply 生成的 rollback PatchSet 再次通过 Office.js 执行。

## 离线文件路径保持不变

- `docx_*` 工具仍然负责离线 DOCX 文件事务。
- 文件级写入仍然执行 health、assess、dry-run、backup、apply、validate、diff。
- LibreOffice 或 PNG 渲染只属于离线交付后的视觉验证辅助，不是 live Word session 写入路径。

## 已验证

- Python compileall。
- 标准 DOCX smoke test。
- 结构回归测试。
- Word session command queue smoke。
- .NET Open XML SDK build 与 regression。
- Office add-in TypeScript build。

真实 `Word.run(...)` 仍需要在 Microsoft Word host 中加载 taskpane 后执行；浏览器调试无法模拟 Word host API。
