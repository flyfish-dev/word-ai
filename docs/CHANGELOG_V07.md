# v0.7 Changelog - Codex Skill And One-Command Startup

## English

v0.7 focuses on making Word AI easier to adopt in Codex while keeping the original PatchSet safety model.

### Added

- Formal `word-ai` Codex Skill in `skills/word-ai/SKILL.md`.
- Compatibility Skill entry in `codex-skill/SKILL.md`.
- Clear mode-selection rules:
  - use offline `docx_*` tools for file-path DOCX editing and batch workflows;
  - use live `word_session_*` tools for the currently open Microsoft Word document through Office.js;
  - never silently fall back from a requested live Word session to offline file editing.
- OfficeCLI auxiliary policy:
  - added allowlisted MCP wrappers: `officecli_view_html`, `officecli_view_screenshot`, `officecli_view_issues`, `officecli_query`, and `officecli_validate`;
  - allowed only for read-only or low-risk evidence: `view html`, `view screenshot`, `view issues`, `query --json`, and `validate`;
  - mutation commands such as `set`, `add`, `remove`, `raw-set`, `batch`, and `merge` are not part of the default Word AI flow.
- One-command local setup scripts:
  - `scripts/install.sh`
  - `scripts/start.sh`
  - `scripts/stop.sh`
  - Windows PowerShell equivalents.
- `word-ai` console helper with:
  - `.venv/bin/word-ai --root "$PWD" doctor`
  - `.venv/bin/word-ai --root "$PWD" codex-config`
- Skill validation script for CI.

### Preserved

- All writes remain governed by Word AI PatchSet, assess, dry-run, approval, audit, validation, diff, and rollback workflows.
- The Office.js path remains content-control focused for live Word sessions.
- OfficeCLI is optional evidence, not the authoritative writer.

## 中文

v0.7 重点降低 Codex 试用门槛，同时保留 Word AI 原有的 PatchSet 安全模型。

### 新增

- 正式 `word-ai` Codex Skill：`skills/word-ai/SKILL.md`。
- 兼容旧入口：`codex-skill/SKILL.md`。
- 明确选择规则：
  - 文件路径 DOCX 和批处理使用离线 `docx_*`；
  - 当前打开的 Word 文档使用 Office.js 闭环的 `word_session_*`；
  - 用户明确要求 live Word session 时，不允许静默退回离线文件编辑。
- OfficeCLI 辅助策略：
  - 新增白名单 MCP wrapper：`officecli_view_html`、`officecli_view_screenshot`、`officecli_view_issues`、`officecli_query`、`officecli_validate`；
  - 只允许只读/低风险证据能力：`view html`、`view screenshot`、`view issues`、`query --json`、`validate`；
  - 默认不开放 `set`、`add`、`remove`、`raw-set`、`batch`、`merge` 等写入/变更能力。
- 一键安装和启动脚本：
  - `scripts/install.sh`
  - `scripts/start.sh`
  - `scripts/stop.sh`
  - Windows PowerShell 等价脚本。
- `word-ai` 命令行助手：
  - `.venv/bin/word-ai --root "$PWD" doctor`
  - `.venv/bin/word-ai --root "$PWD" codex-config`
- CI 使用的 Skill 校验脚本。

### 保留

- 所有正式写入仍必须经过 PatchSet、assess、dry-run、approval、audit、validate、diff 和 rollback 机制。
- Office.js live session 仍以内容控件为主要安全写入边界。
- OfficeCLI 仅作为可选证据来源，不作为本项目默认写入后端。
