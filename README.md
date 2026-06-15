# Word AI

**Structure-preserving Word (DOCX) editing for AI agents.**

Word AI is an open-source MCP server and Office.js bridge for safe, auditable, incremental editing of Microsoft Word documents. It is designed for Codex, OpenAI Agents, and other MCP clients that need to edit `.docx` files without rebuilding the document or damaging styles, numbering, tables, images, fields, headers, footers, and relationships.

中文简介见下方：[中文说明](#中文说明)。

## Why Word AI

AI systems are good at generating text, but Word documents are structured packages. A naive DOCX-to-Markdown-to-DOCX workflow can break numbering, styles, tables, images, fields, references, and layout. Word AI keeps the original DOCX structure as the source of truth:

- Agents generate constrained `PatchSet` operations.
- The local engine applies targeted OOXML/Open XML edits.
- Every write goes through assessment, dry-run, backup, validation, audit, and diff.
- The Office.js add-in provides a Word taskpane for anchor governance and human-in-the-loop workflows.

## Key Features

- **49 MCP tools** for DOCX inspection, anchors, headings, paragraphs, tables, fields, images, comments, revisions, PatchSet planning, dry-run, apply, validation, rollback, and diff.
- **PatchSet-only writes**. No full document rebuilds, no Markdown/HTML round-trips, and no direct source overwrite by default.
- **Content-control first editing** using stable Word content control tags such as `WORD-AI:SRS:1.0:overview`.
- **Strong preconditions** with `source_sha256`, `expected_old_sha256`, and `expected_old_text`.
- **Structure validation** for package parts, content controls, tables, paragraphs, fields, comments, images, revisions, and protected body blocks.
- **Python MCP runtime** for local agent integration.
- **.NET 8 Open XML SDK engine** for production-grade typed Open XML processing.
- **Office.js taskpane** for Word-side anchors, PatchSet preview, dry-run, apply, and open-document content-control editing with hash checks.
- **Local HTTP bridge** secured by a local token and localhost-only CORS for Office add-in workflows.

## Architecture

```text
Codex / Agent / MCP Client
        |
        v
Word AI MCP Server
        |
        +--> Python OOXML engine
        +--> .NET Open XML SDK engine
        +--> Office bridge HTTP API
        |
        v
Original DOCX -> PatchSet -> Candidate DOCX -> Validation -> Output DOCX + Audit JSON + Diff
```

The Office.js taskpane is the Word session layer. It creates and lists content controls, connects to the local bridge, builds safe PatchSets, runs dry-runs, and applies approved edits.

## Quick Start

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

PYTHONPATH=. python scripts/run_smoke_test.py
PYTHONPATH=. python scripts/run_structure_regression.py
```

Build the .NET engine:

```bash
dotnet --version  # requires .NET SDK 8
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. python scripts/run_dotnet_regression.py
```

Build the Office add-in:

```bash
cd office-addin
npm install
npm run build
```

## Use With Codex MCP

Add this to your Codex MCP config:

```toml
[mcp_servers.word_ai]
command = "/absolute/path/to/word-ai/.venv/bin/python"
args = ["-m", "word_ai_mcp.server", "--root", "/absolute/path/to/word-ai"]
enabled = true
startup_timeout_sec = 30

[mcp_servers.word_ai.env]
PYTHONPATH = "/absolute/path/to/word-ai"
```

Recommended approval policy for write tools:

- `docx_dry_run_patchset`
- `docx_apply_patchset`
- `docx_backup`
- `docx_restore_backup`
- `docx_rollback`
- sidecar export tools

Example prompt:

```text
Use word_ai to inspect examples/sample_contract.docx, list content controls, read WORD-AI:SRS:1.0:overview, and prepare a PatchSet. Run assess and dry-run before applying.
```

## Office.js Bridge

Start the local bridge:

```bash
PYTHONPATH=. python -m word_ai_mcp.server_http \
  --root "$PWD" \
  --host 127.0.0.1 \
  --port 8765
```

Start the taskpane:

```bash
cd office-addin
npm run dev
```

Then sideload `office-addin/manifest.xml` in Word. The taskpane runs at `https://localhost:3000/taskpane.html` and proxies `/bridge/*` to the local bridge. The bridge prints a local token at startup. Use that token in the taskpane.

## Safe Editing Workflow

```text
docx_health_check
  -> docx_map / docx_list_anchors / docx_list_content_controls
  -> docx_read_content_control / docx_read_table_cell / docx_read_paragraph
  -> generate PatchSet with source_sha256 and expected_old_sha256
  -> docx_assess_patchset
  -> docx_dry_run_patchset
  -> docx_backup
  -> docx_apply_patchset
  -> docx_validate / docx_compare_structure
  -> docx_text_diff
```

## Documentation

- [Documentation Index](docs/README.md)
- [Getting Started](docs/GETTING_STARTED.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Architecture in English](docs/ARCHITECTURE.en.md)
- [Tool Contract](docs/TOOL_CONTRACT.md)
- [Security Design](docs/SECURITY.md)
- [QA Report](docs/QA_REPORT.md)
- [Validation Matrix](docs/VALIDATION_MATRIX.md)

## Repository Status

Word AI is currently a local-first developer tool. It is suitable for controlled DOCX editing experiments, agent integration, and internal workflow pilots. For production remote MCP deployments, use proper MCP Streamable HTTP transport, authentication, network controls, audit storage, and render/visual diff infrastructure.

## License

MIT License. See [LICENSE](LICENSE).

---

# 中文说明

**面向 AI Agent 的 Word（DOCX）结构稳定编辑系统。**

Word AI 是一个开源 MCP Server 与 Office.js Bridge，用于安全、可审计、可回滚地增量编辑 Microsoft Word 文档。它面向 Codex、OpenAI Agents 和其他 MCP 客户端，核心目标是避免“整篇重建”导致样式、编号、表格、图片、字段、页眉页脚和关系文件漂移。

## 为什么需要 Word AI

大模型擅长生成文本，但 Word 文档不是纯文本，而是复杂的 OOXML 包结构。直接把 DOCX 转成 Markdown/HTML 再转回 DOCX，容易破坏编号、样式、字段、交叉引用、图片关系和版式。Word AI 的设计原则是：

- 模型只生成受约束的 `PatchSet`。
- 本地执行层对 OOXML / Open XML 做定点修改。
- 每次写入必须经过评估、dry-run、备份、验证、审计和 diff。
- Office.js 加载项负责 Word 端锚点治理和人工审批体验。

## 核心能力

- **49 个 MCP tools**，覆盖 DOCX 检查、锚点、标题、段落、表格、字段、图片、批注、修订、PatchSet 规划、dry-run、正式写入、验证、回滚和 diff。
- **所有正式写入收口到 PatchSet**，默认不覆盖源文件。
- **优先使用内容控件 tag**，例如 `WORD-AI:SRS:1.0:overview`。
- **并发安全前置条件**：`source_sha256`、`expected_old_sha256`、`expected_old_text`。
- **结构验证**：保护 package parts、内容控件、表格、段落、字段、图片、批注、修订痕迹和正文块顺序。
- **Python MCP Server**，便于本地 Agent 集成。
- **.NET 8 Open XML SDK 引擎**，面向生产级 typed Open XML 处理。
- **Office.js taskpane**，支持创建/列出锚点、构建 PatchSet、预览、dry-run、apply，以及对当前打开的 Word 文档进行 hash 校验后的内容控件写入。

## 快速开始

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

PYTHONPATH=. python scripts/run_smoke_test.py
PYTHONPATH=. python scripts/run_structure_regression.py
```

构建 .NET 引擎：

```bash
dotnet --version  # 需要 .NET SDK 8
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. python scripts/run_dotnet_regression.py
```

构建 Office 加载项：

```bash
cd office-addin
npm install
npm run build
```

## 在 Codex 中使用

在 Codex MCP 配置中加入：

```toml
[mcp_servers.word_ai]
command = "/absolute/path/to/word-ai/.venv/bin/python"
args = ["-m", "word_ai_mcp.server", "--root", "/absolute/path/to/word-ai"]
enabled = true
startup_timeout_sec = 30

[mcp_servers.word_ai.env]
PYTHONPATH = "/absolute/path/to/word-ai"
```

推荐将写入类工具设置为需要审批，尤其是 `docx_apply_patchset`、`docx_restore_backup`、`docx_rollback`。

## 标准安全流程

```text
docx_health_check
  -> docx_map / docx_list_anchors / docx_list_content_controls
  -> docx_read_content_control / docx_read_table_cell / docx_read_paragraph
  -> 生成带 hash 前置条件的 PatchSet
  -> docx_assess_patchset
  -> docx_dry_run_patchset
  -> docx_backup
  -> docx_apply_patchset
  -> docx_validate / docx_compare_structure
  -> docx_text_diff
```

## 文档

- [文档导航](docs/README.md)
- [快速上手](docs/GETTING_STARTED.md)
- [架构说明](docs/ARCHITECTURE.md)
- [英文架构说明](docs/ARCHITECTURE.en.md)
- [工具契约](docs/TOOL_CONTRACT.md)
- [安全设计](docs/SECURITY.md)
- [QA 验证报告](docs/QA_REPORT.md)
- [验证矩阵](docs/VALIDATION_MATRIX.md)
