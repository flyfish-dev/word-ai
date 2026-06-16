# Contributing / 贡献指南

Thank you for contributing to Word AI.

谢谢你参与 Word AI。

## Development Setup

```bash
bash scripts/install.sh
```

The default installer sets up the Python MCP server, Office.js taskpane, .NET Open XML engine when available, Codex config snippet, and the `word-ai` Agent Skill for Codex, Claude Code, and compatible clients. Use `bash scripts/install.sh --skip-node`, `--skip-dotnet`, or `--no-agent-skills` only when intentionally working on a narrower area.

## Required Checks

Run these before opening a pull request:

```bash
.venv/bin/word-ai --root "$PWD" doctor
PYTHONPATH=. .venv/bin/python -m compileall word_ai_mcp scripts
PYTHONPATH=. .venv/bin/python scripts/run_smoke_test.py
PYTHONPATH=. .venv/bin/python scripts/run_structure_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_outline_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_word_session_smoke.py
PYTHONPATH=. .venv/bin/python scripts/validate_word_ai_skill.py
PYTHONPATH=. .venv/bin/python -m word_ai_mcp.quickstart --root "$PWD" codex-config --output /tmp/word-ai-codex.toml
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py

cd office-addin
npm run build
npm audit --omit=dev --registry=https://registry.npmjs.org --json
```

## Contribution Rules

- Do not rebuild whole DOCX files as an editing strategy.
- Do not introduce Markdown/HTML round-trip editing.
- Do not add broad write APIs that bypass PatchSet.
- Do not expose OfficeCLI mutation commands (`set`, `add`, `remove`, `raw-set`, `batch`, `merge`) as default Word AI flows.
- Prefer content-control tags as edit anchors.
- Add or update validation coverage for every write operation.
- Keep generated files, local tokens, build artifacts, and Office lock files out of commits.
- When updating installation docs, present MCP Registry/MCPB and Agent Skill installation as the primary path; keep npm as a secondary convenience channel for hosts that cannot consume MCP Registry packages yet.

## 中文说明

提交 PR 前请完成上述检查。涉及 DOCX 写入逻辑时，必须遵守结构稳定优先原则：

- 不重建整篇 DOCX。
- 不绕过 PatchSet。
- 不把 OfficeCLI 写入命令作为默认流程；除非包进 Word AI 的 PatchSet、dry-run、audit、rollback 和审批。
- 写入前必须 assess 和 dry-run。
- 写入后必须 validate 和 diff。
- 高风险操作必须带旧文本或旧 hash 前置条件。
