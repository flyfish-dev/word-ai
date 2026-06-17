# v0.8.4 Changelog - Agent PatchSet Compatibility

v0.8.4 is a patch release focused on real-world agent usability. It fixes a failure mode where an agent-generated PatchSet used common alias fields such as `operation`, `target_tag`, `new_text`, or `text_sha256`, causing the server to report `Unsupported op: None` before the normal safety lifecycle could run.

## Highlights

- Added a shared PatchSet normalization layer for agent-generated payloads.
- Accepts common operation-name aliases such as `operation`, `operation_type`, `type`, and `action`.
- Normalizes camelCase, kebab-case, and spaced operation names into canonical snake_case operations, for example `replaceContentControlText` -> `replace_content_control_text`.
- Maps common field aliases such as `target_tag`, `content_control_tag`, `new_text`, `replacement_text`, `text_sha256`, and `expected_sha256` to the canonical PatchSet schema.
- Preserves the existing safety model: no new mutation operations, no bypass of `assess`, `dry-run`, hash preconditions, validation, audit JSON, rollback, or diff.
- Applies the same normalization to Python OOXML fallback, .NET Open XML backend, and live Office.js `word_session_*` PatchSet commands.
- Improves missing-operation diagnostics to clearly report `patchset.operations[0].op is required`.

## Validation

- Added `scripts/run_patchset_alias_regression.py`.
- The new regression covers both the default `.NET` engine path and forced `python` fallback path.
- The release workflow now runs the alias regression in the MCP smoke gate.
- Local verification covered smoke tests, structure regression, outline/TOC regression, Word session smoke, Office bridge smoke, skill validation, and .NET regression.

## 中文

v0.8.4 是面向实际 Agent 使用问题的补丁版本。它修复了 Agent 生成的 PatchSet 使用 `operation`、`target_tag`、`new_text`、`text_sha256` 等常见别名字段时，服务端在安全流程前报出 `Unsupported op: None` 的问题。

## 重点

- 新增统一 PatchSet 归一化层，专门处理 Agent 生成的常见 payload 写法。
- 兼容 `operation`、`operation_type`、`type`、`action` 等操作名别名。
- 将 camelCase、kebab-case、空格分隔操作名归一化为正式 snake_case 操作，例如 `replaceContentControlText` -> `replace_content_control_text`。
- 将 `target_tag`、`content_control_tag`、`new_text`、`replacement_text`、`text_sha256`、`expected_sha256` 等常见字段别名映射为 canonical PatchSet schema。
- 保留原有安全模型：不新增写入操作，不绕过 `assess`、`dry-run`、hash 前置条件、验证、audit JSON、rollback 和 diff。
- Python OOXML fallback、.NET Open XML 后端、Office.js live `word_session_*` PatchSet 命令共用同一归一化入口。
- 缺少操作名时返回更清晰的 `patchset.operations[0].op is required`。

## 验证

- 新增 `scripts/run_patchset_alias_regression.py`。
- 新回归同时覆盖默认 `.NET` 引擎路径和强制 `python` fallback 路径。
- Release workflow 的 MCP smoke gate 已加入该回归。
- 本地验证覆盖 smoke tests、结构回归、outline/TOC 回归、Word session smoke、Office bridge smoke、skill validation 和 .NET regression。
