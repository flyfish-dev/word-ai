# 贡献指南

| 文档 | 预览 |
| --- | --- |
| [中文](CONTRIBUTING.zh-CN.md) | Word AI 开发环境、必跑检查和结构稳定贡献规则。 |
| [English](CONTRIBUTING.md) | Development setup, required checks, and contribution rules for safe Word AI changes. |

谢谢你参与 Word AI。

## 开发环境

```bash
bash scripts/install.sh
```

默认安装脚本会配置 Python MCP Server、Office.js taskpane、可用时构建 .NET Open XML 引擎、生成 Codex 配置片段，并把正式 `word-ai` Agent Skill 安装到 Codex、Claude Code 和兼容客户端。只有在明确只处理某个子模块时，才使用 `bash scripts/install.sh --skip-node`、`--skip-dotnet` 或 `--no-agent-skills`。

## 必跑检查

提交 Pull Request 前请运行：

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

## 贡献规则

- 不把“重建整篇 DOCX”作为编辑策略。
- 不引入 Markdown/HTML 往返编辑链路。
- 不新增绕过 PatchSet 的宽泛写入 API。
- 不把 OfficeCLI 写入命令（`set`、`add`、`remove`、`raw-set`、`batch`、`merge`）作为 Word AI 默认流程开放。
- 优先使用内容控件 tag 作为编辑锚点。
- 每个写入操作都要新增或更新验证覆盖。
- 不提交生成文件、本地 token、构建产物和 Office lock 文件。
- 更新安装文档时，优先呈现 MCP Registry/MCPB 与 Agent Skill；npm 只作为暂不支持 MCP Registry 包的 host 的第二便利渠道。

## DOCX 写入要求

涉及 DOCX 写入逻辑时，必须遵守结构稳定优先原则：写入前必须 assess 和 dry-run，写入后必须 validate 和 diff，高风险操作必须带旧文本或旧 hash 前置条件。任何 OfficeCLI 变更能力都必须先包进 Word AI 的 PatchSet、dry-run、audit、rollback 和审批模型。
