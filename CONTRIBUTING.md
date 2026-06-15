# Contributing / 贡献指南

Thank you for contributing to Word AI.

谢谢你参与 Word AI。

## Development Setup

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cd office-addin
npm install
cd ..
```

## Required Checks

Run these before opening a pull request:

```bash
PYTHONPATH=. python -m compileall word_ai_mcp scripts
PYTHONPATH=. python scripts/run_smoke_test.py
PYTHONPATH=. python scripts/run_structure_regression.py
PYTHONPATH=. python scripts/run_dotnet_regression.py
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release

cd office-addin
npm run build
npm audit --omit=dev --registry=https://registry.npmjs.org --json
```

## Contribution Rules

- Do not rebuild whole DOCX files as an editing strategy.
- Do not introduce Markdown/HTML round-trip editing.
- Do not add broad write APIs that bypass PatchSet.
- Prefer content-control tags as edit anchors.
- Add or update validation coverage for every write operation.
- Keep generated files, local tokens, build artifacts, and Office lock files out of commits.

## 中文说明

提交 PR 前请完成上述检查。涉及 DOCX 写入逻辑时，必须遵守结构稳定优先原则：

- 不重建整篇 DOCX。
- 不绕过 PatchSet。
- 写入前必须 assess 和 dry-run。
- 写入后必须 validate 和 diff。
- 高风险操作必须带旧文本或旧 hash 前置条件。

