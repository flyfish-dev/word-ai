# Contributing

| Language | Preview |
| --- | --- |
| [English](CONTRIBUTING.md) | Development setup, required checks, and contribution rules for safe Word AI changes. |
| [中文](CONTRIBUTING.zh-CN.md) | Word AI 开发环境、必跑检查和结构稳定贡献规则。 |

Thank you for contributing to Word AI.

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
