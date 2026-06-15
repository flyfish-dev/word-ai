# Word AI

<!-- mcp-name: io.github.flyfish-dev/word-ai -->

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

- **63 MCP tools** for DOCX inspection, anchors, headings, paragraphs, tables, fields, images, comments, revisions, PatchSet planning, dry-run, apply, validation, rollback, diff, live Word session editing, and optional read-only OfficeCLI evidence.
- **PatchSet-only writes**. No full document rebuilds, no Markdown/HTML round-trips, and no direct source overwrite by default.
- **Content-control first editing** using stable Word content control tags such as `WORD-AI:SRS:1.0:overview`.
- **Strong preconditions** with `source_sha256`, `expected_old_sha256`, and `expected_old_text`.
- **Structure validation** for package parts, content controls, tables, paragraphs, fields, comments, images, revisions, and protected body blocks.
- **Python MCP runtime** for local agent integration.
- **.NET 8 Open XML SDK engine** for production-grade typed Open XML processing.
- **Office.js taskpane** for Word-side anchors, PatchSet preview, dry-run, apply, and open-document content-control editing with hash checks.
- **Live Word session tools** (`word_session_*`) so Codex can read, preview, apply, and roll back edits in the currently open Word document through Office.js.
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
        +--> File-backed Word session command queue
        |
        v
Original DOCX -> PatchSet -> Candidate DOCX -> Validation -> Output DOCX + Audit JSON + Diff
```

The Office.js taskpane is the Word session layer. It creates and lists content controls, connects to the local bridge, registers the current Word document as a live session, polls commands queued by Codex, executes supported PatchSet operations through Office.js, and returns audit/rollback data.

## Quick Start

Use the most native distribution path your agent host supports:

1. **MCP Registry / MCPB first**: install the MCP server from the official MCP Registry using server name `io.github.flyfish-dev/word-ai`.
2. **Agent Skill next**: install the `word-ai` Skill so Codex, Claude Code, and compatible agents know when to choose offline `docx_*` versus live `word_session_*`.
3. **Local source install for full Word sessions**: use this when you need the Office.js taskpane, localhost bridge, .NET Open XML regression path, or development workflow.
4. **npm as a secondary channel**: use npm only when your MCP host cannot consume MCP Registry/MCPB yet, or when you want a no-clone stdio server command.

MCP Registry details:

- Server name: `io.github.flyfish-dev/word-ai`
- Registry metadata: [server.json](server.json)
- MCPB package: `https://github.com/flyfish-dev/word-ai/releases/download/v0.8.1/word-ai-0.8.1.mcpb`
- Registry latest API: `https://registry.modelcontextprotocol.io/v0.1/servers/io.github.flyfish-dev%2Fword-ai/versions/latest`

Install the Skill and full local runtime:

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

bash scripts/install.sh
bash scripts/start.sh
```

This installs the Python MCP server, builds the Office.js taskpane, builds the .NET Open XML engine when .NET SDK 8 is available, writes `.wordai/codex-config.toml`, and installs the `word-ai` skill into Codex, Claude Code, and detected compatible agent clients.

Install or refresh only the Agent Skill:

```bash
python3 scripts/install_agent_skills.py
```

For browser-only taskpane debugging:

```bash
bash scripts/start.sh --http
```

On Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\install.ps1
powershell -ExecutionPolicy Bypass -File scripts\start.ps1
```

For a readiness check:

```bash
.venv/bin/word-ai --root "$PWD" doctor
```

Developer checks:

```bash
PYTHONPATH=. .venv/bin/python scripts/run_smoke_test.py
PYTHONPATH=. .venv/bin/python scripts/run_structure_regression.py
```

Build the .NET engine:

```bash
dotnet --version  # requires .NET SDK 8
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py
```

Build the Office add-in:

```bash
cd office-addin
npm install
npm run build
```

## Agent Skill Auto-Install

Word AI ships a formal `word-ai` Agent Skill. This is the preferred way to teach agents the safe workflow, even when the MCP server is installed through the MCP Registry. The installer copies the Skill into the locations that current agent clients scan automatically:

- Codex official user skills: `~/.agents/skills/word-ai`
- Codex app compatibility skills: `~/.codex/skills/word-ai`
- Claude Code personal skills: `~/.claude/skills/word-ai`
- Existing compatible clients when detected: Cursor, Windsurf, GitHub Copilot, and OpenClaw skill folders

Install or refresh only the skills:

```bash
python3 scripts/install_agent_skills.py
```

Advanced targets:

```bash
python3 scripts/install_agent_skills.py --agents all
python3 scripts/install_agent_skills.py --project
python3 scripts/install_agent_skills.py --dry-run
```

After installation, start a new agent session or restart the client if the skill does not appear immediately. The skill can then be invoked directly as `word-ai` / `$word-ai`, or selected implicitly when a DOCX editing task mentions Word, Office.js, content controls, PatchSet, validation, rollback, or audit.

## Global MCP Distribution

Word AI is published for discovery through the official MCP Registry and MCPB distribution. Prefer this channel for MCP host installation because it carries standardized server metadata, versioning, transport details, and provenance:

- MCP server name: `io.github.flyfish-dev/word-ai`
- MCPB package: `https://github.com/flyfish-dev/word-ai/releases/download/v0.8.1/word-ai-0.8.1.mcpb`
- Registry metadata: [server.json](server.json)
- Publishing guide: [MCP Registry Publishing](docs/REGISTRY_PUBLISHING.md)

Local container smoke test:

```bash
docker build -t word-ai:local .
docker run --rm -i \
  -v "$PWD:/workspace" \
  -v "$HOME/Downloads:/documents/Downloads" \
  word-ai:local
```

The MCP Registry release uses a public MCPB artifact for one-click-friendly local server installation. The MCPB package requires Python 3.10+ and bootstraps a local virtual environment on first run. The Dockerfile remains available for local or self-hosted builds. For full Office.js live-session editing, use the local install path because the Word taskpane and localhost bridge must run on the user's machine.

## Secondary npm Channel

npm is a convenience fallback for clients that do not yet consume MCP Registry/MCPB packages, for CI smoke tests, and for quick no-clone stdio server startup. It is not the primary discovery channel.

Recommended scoped package:

```bash
npm exec --yes --package @flyfish-dev/word-ai -- word-ai-mcp --root "$PWD" --allow-root "$HOME/Downloads"
npm exec --yes --package @flyfish-dev/word-ai -- word-ai --root "$PWD" doctor
```

Unscoped compatibility package:

```bash
npx -y word-ai-mcp --root "$PWD"
npm exec --yes --package word-ai-mcp -- word-ai --root "$PWD" doctor
npm exec --yes --package word-ai-mcp -- word-ai-mcp --root "$PWD"
```

After a global install, the same commands are available directly:

```bash
npm install -g @flyfish-dev/word-ai
word-ai --root "$PWD" doctor
word-ai-mcp --root "$PWD" --allow-root "$HOME/Downloads"
```

The first npm run creates a local Python virtual environment under the user cache and installs Word AI Python dependencies automatically. Set `WORD_AI_PYTHON=/path/to/python3.10+` if Python discovery needs help.

## Use With Codex MCP

The installer writes a ready-to-merge MCP configuration snippet:

```bash
cat .wordai/codex-config.toml
```

Add it to your Codex MCP config. The generated snippet includes write-tool approval gates. A minimal manual version is:

```toml
[mcp_servers.word_ai]
command = "/absolute/path/to/word-ai/.venv/bin/python"
args = [
  "-m", "word_ai_mcp.server",
  "--root", "/absolute/path/to/word-ai",
  "--allow-root", "/Users/you/Downloads",
  "--allow-root", "/Users/you/Documents"
]
enabled = true
startup_timeout_sec = 30

[mcp_servers.word_ai.env]
PYTHONPATH = "/absolute/path/to/word-ai"
```

Secondary npm-based Codex setup, for hosts that cannot install from MCP Registry/MCPB yet:

```toml
[mcp_servers.word_ai]
command = "npm"
args = [
  "exec",
  "--yes",
  "--package",
  "@flyfish-dev/word-ai",
  "--",
  "word-ai-mcp",
  "--root",
  "/absolute/path/to/workspace",
  "--allow-root",
  "/Users/you/Downloads",
  "--allow-root",
  "/Users/you/Documents"
]
enabled = true
startup_timeout_sec = 60
```

You can replace `@flyfish-dev/word-ai` with the unscoped compatibility package `word-ai-mcp` in the npm-based Codex config.

`--root` is the primary workspace for relative paths and Word AI sidecars. Repeat `--allow-root` for external document folders you want Codex to edit, such as Downloads, Documents, or a team project folder. The installer-generated `.wordai/codex-config.toml` includes common user document folders automatically.

Recommended approval policy for write tools:

- `docx_dry_run_patchset`
- `docx_apply_patchset`
- `docx_backup`
- `docx_restore_backup`
- `docx_rollback`
- `word_session_apply_patchset`
- `word_session_wrap_selection`
- `word_session_rollback`
- sidecar export tools

Example prompt:

```text
Use word_ai to inspect examples/sample_contract.docx, list content controls, read WORD-AI:SRS:1.0:overview, and prepare a PatchSet. Run assess and dry-run before applying.
```

For the currently open Word document, load the Office add-in, connect the bridge, then ask Codex:

```text
Use word_ai to list active Word sessions, read WORD-AI:SRS:1.0:overview from the live Word session, preview a PatchSet through Office.js, then apply it to the open document and return the audit plus rollback PatchSet.
```

## Office.js Bridge And Live Word Sessions

The easiest path is:

```bash
bash scripts/start.sh
```

For manual startup, start the local bridge:

```bash
PYTHONPATH=. .venv/bin/python -m word_ai_mcp.server_http \
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

Once connected inside Word, the taskpane registers a live session under `.wordai/sessions`. Codex can then use:

- `word_session_list`
- `word_session_snapshot`
- `word_session_read_content_control`
- `word_session_preview_patchset`
- `word_session_apply_patchset`
- `word_session_wrap_selection`
- `word_session_rollback`
- `word_session_command_status`

This path edits the currently open Word document through Office.js. `word_session_apply_patchset` performs a live preflight against the open document, checks `expected_old_sha256`, applies supported content-control operations, returns an audit object, and generates a rollback PatchSet. The offline DOCX path still uses `docx_*` tools and the OOXML/Open XML validator.

## OfficeCLI Compatibility Policy

Word AI can optionally use OfficeCLI as auxiliary evidence for read-only or low-risk checks: `view html`, `view screenshot`, `view issues`, `query --json`, and `validate`. OfficeCLI mutation commands such as `set`, `add`, `remove`, `raw-set`, `batch`, and `merge` are not part of the default Word AI workflow unless they are wrapped by Word AI PatchSet, dry-run, audit, rollback, and explicit approval gates.

The MCP server exposes this integration only through allowlisted wrappers: `officecli_view_html`, `officecli_view_screenshot`, `officecli_view_issues`, `officecli_query`, and `officecli_validate`. If OfficeCLI is not installed, these tools return `available=false` and the core Word AI workflow continues to use `docx_*` and `word_session_*`.

Word AI borrows useful OfficeCLI design ideas such as schema/help-first usage, semantic paths, watch/render evidence, template merge concepts, and dump/batch inspection. The authoritative write model remains Word AI PatchSet.

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
- [Word AI Codex Skill](skills/word-ai/SKILL.md)
- [Architecture](docs/ARCHITECTURE.md)
- [Architecture in English](docs/ARCHITECTURE.en.md)
- [Tool Contract](docs/TOOL_CONTRACT.md)
- [Security Design](docs/SECURITY.md)
- [QA Report](docs/QA_REPORT.md)
- [Validation Matrix](docs/VALIDATION_MATRIX.md)
- [MCP Registry Publishing](docs/REGISTRY_PUBLISHING.md)
- [v0.8.1 Changelog](docs/CHANGELOG_V081.md)
- [v0.8.0 Changelog](docs/CHANGELOG_V080.md)
- [v0.7.1 Changelog](docs/CHANGELOG_V071.md)
- [v0.7 Changelog](docs/CHANGELOG_V07.md)

## Repository Status

Word AI is currently a local-first developer tool. It is suitable for controlled DOCX editing experiments, agent integration, and internal workflow pilots. For production remote MCP deployments, use proper MCP Streamable HTTP transport, authentication, network controls, audit storage, and render/visual diff infrastructure.

## License

GNU Affero General Public License v3.0 or later (`AGPL-3.0-or-later`). See [LICENSE](LICENSE).

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

- **63 个 MCP tools**，覆盖 DOCX 检查、锚点、标题、段落、表格、字段、图片、批注、修订、PatchSet 规划、dry-run、正式写入、验证、回滚、diff、Word 打开会话内编辑，以及可选 OfficeCLI 只读证据能力。
- **所有正式写入收口到 PatchSet**，默认不覆盖源文件。
- **优先使用内容控件 tag**，例如 `WORD-AI:SRS:1.0:overview`。
- **并发安全前置条件**：`source_sha256`、`expected_old_sha256`、`expected_old_text`。
- **结构验证**：保护 package parts、内容控件、表格、段落、字段、图片、批注、修订痕迹和正文块顺序。
- **Python MCP Server**，便于本地 Agent 集成。
- **.NET 8 Open XML SDK 引擎**，面向生产级 typed Open XML 处理。
- **Office.js taskpane**，支持创建/列出锚点、构建 PatchSet、预览、dry-run、apply，以及对当前打开的 Word 文档进行 hash 校验后的内容控件写入。
- **Word 会话 MCP 工具**，Codex 可以通过 `word_session_*` 读取当前打开文档、预览 PatchSet、调用 Office.js 写入并获取审计和 rollback PatchSet。

## 快速开始

请优先使用 Agent host 原生支持的分发方式：

1. **优先 MCP Registry / MCPB**：在支持 MCP Registry 或 MCP marketplace 的客户端里安装 `io.github.flyfish-dev/word-ai`。
2. **同时安装 Agent Skill**：安装 `word-ai` Skill，让 Codex、Claude Code 等 Agent 明确知道何时使用离线 `docx_*`，何时使用 live `word_session_*`。
3. **完整 Word 会话能力使用本地源码安装**：需要 Office.js taskpane、localhost bridge、.NET Open XML 回归或开发调试时使用。
4. **npm 是第二渠道**：仅作为 MCP host 暂不支持 Registry/MCPB，或需要免 clone stdio server 命令时的便利入口。

MCP Registry 信息：

- MCP server 名称：`io.github.flyfish-dev/word-ai`
- Registry 元数据：[server.json](server.json)
- MCPB 包：`https://github.com/flyfish-dev/word-ai/releases/download/v0.8.1/word-ai-0.8.1.mcpb`
- Registry latest API：`https://registry.modelcontextprotocol.io/v0.1/servers/io.github.flyfish-dev%2Fword-ai/versions/latest`

安装 Skill 和完整本地运行时：

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

bash scripts/install.sh
bash scripts/start.sh
```

安装脚本会安装 Python MCP 依赖、构建 Office.js taskpane、在可用时构建 .NET Open XML 引擎、生成 `.wordai/codex-config.toml`，并把正式 `word-ai` Skill 安装到 Codex、Claude Code 以及已检测到的兼容 Agent 客户端。

只安装或刷新 Agent Skill：

```bash
python3 scripts/install_agent_skills.py
```

浏览器调试 taskpane：

```bash
bash scripts/start.sh --http
```

Windows PowerShell：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\install.ps1
powershell -ExecutionPolicy Bypass -File scripts\start.ps1
```

健康检查：

```bash
.venv/bin/word-ai --root "$PWD" doctor
```

开发者验证：

```bash
PYTHONPATH=. .venv/bin/python scripts/run_smoke_test.py
PYTHONPATH=. .venv/bin/python scripts/run_structure_regression.py
```

构建 .NET 引擎：

```bash
dotnet --version  # 需要 .NET SDK 8
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py
```

构建 Office 加载项：

```bash
cd office-addin
npm install
npm run build
```

## Agent Skill 一键安装

Word AI 内置正式 `word-ai` Agent Skill。即使 MCP Server 已通过 MCP Registry 安装，也建议同时安装 Skill，因为它定义了安全编辑流程、OfficeCLI 边界和 live/offline 选择规则。默认安装会自动写入：

- Codex 官方用户 skill：`~/.agents/skills/word-ai`
- Codex app 兼容 skill：`~/.codex/skills/word-ai`
- Claude Code 个人 skill：`~/.claude/skills/word-ai`
- 如果本机已存在 Cursor、Windsurf、GitHub Copilot、OpenClaw 的 skill 目录，也会自动安装

只安装或刷新 skill：

```bash
python3 scripts/install_agent_skills.py
```

安装到所有已知目标或项目级目录：

```bash
python3 scripts/install_agent_skills.py --agents all
python3 scripts/install_agent_skills.py --project
```

安装后新开 Agent 会话；如果客户端未立即显示 skill，重启客户端即可。之后可直接用 `word-ai` / `$word-ai` 调用，也可以让 Agent 在 Word、DOCX、Office.js、内容控件、PatchSet、验证、回滚、审计等任务中自动识别。

## 全球 MCP 分发

Word AI 已通过官方 MCP Registry 和 MCPB 包进行全球发现和安装。对于 MCP host，这是优先推荐渠道，因为它提供标准化 server 元数据、版本、transport 与来源验证：

- MCP server 名称：`io.github.flyfish-dev/word-ai`
- MCPB 包：`https://github.com/flyfish-dev/word-ai/releases/download/v0.8.1/word-ai-0.8.1.mcpb`
- Registry 元数据：[server.json](server.json)
- 发布说明：[MCP Registry Publishing](docs/REGISTRY_PUBLISHING.md)

本地容器 smoke test：

```bash
docker build -t word-ai:local .
docker run --rm -i \
  -v "$PWD:/workspace" \
  -v "$HOME/Downloads:/documents/Downloads" \
  word-ai:local
```

MCP Registry 正式发布使用公开 MCPB 资产，便于本地 MCP Server 一键安装。MCPB 包要求 Python 3.10+，首次运行会自动创建本地虚拟环境并安装依赖。Dockerfile 继续保留给本地或自托管构建使用。完整 Office.js live session 编辑仍建议使用本地安装方式，因为 Word taskpane 和 localhost bridge 需要运行在用户机器上。

## npm 第二渠道

npm 是便利备用渠道，适合暂不支持 MCP Registry/MCPB 的客户端、CI smoke test、或免 clone 启动 stdio server。它不是首选发现渠道。

推荐 scoped 包：

```bash
npm exec --yes --package @flyfish-dev/word-ai -- word-ai-mcp --root "$PWD" --allow-root "$HOME/Downloads"
npm exec --yes --package @flyfish-dev/word-ai -- word-ai --root "$PWD" doctor
```

非 scope 兼容包：

```bash
npx -y word-ai-mcp --root "$PWD"
npm exec --yes --package word-ai-mcp -- word-ai --root "$PWD" doctor
npm exec --yes --package word-ai-mcp -- word-ai-mcp --root "$PWD"
```

全局安装后可直接使用：

```bash
npm install -g @flyfish-dev/word-ai
word-ai --root "$PWD" doctor
word-ai-mcp --root "$PWD" --allow-root "$HOME/Downloads"
```

首次 npm 启动会在用户缓存目录创建 Python venv，并自动安装 Word AI Python 依赖。如需指定 Python，可设置 `WORD_AI_PYTHON=/path/to/python3.10+`。

## 在 Codex 中使用

安装脚本会生成可合并到 Codex 的 MCP 配置片段：

```bash
cat .wordai/codex-config.toml
```

将其加入 Codex MCP 配置。生成的配置已经把写入类工具设置为审批模式。最小手动配置如下：

```toml
[mcp_servers.word_ai]
command = "/absolute/path/to/word-ai/.venv/bin/python"
args = [
  "-m", "word_ai_mcp.server",
  "--root", "/absolute/path/to/word-ai",
  "--allow-root", "/Users/you/Downloads",
  "--allow-root", "/Users/you/Documents"
]
enabled = true
startup_timeout_sec = 30

[mcp_servers.word_ai.env]
PYTHONPATH = "/absolute/path/to/word-ai"
```

npm 版 Codex 配置适合作为第二渠道，用于暂不能通过 MCP Registry/MCPB 安装的 host：

```toml
[mcp_servers.word_ai]
command = "npm"
args = [
  "exec",
  "--yes",
  "--package",
  "@flyfish-dev/word-ai",
  "--",
  "word-ai-mcp",
  "--root",
  "/absolute/path/to/workspace",
  "--allow-root",
  "/Users/you/Downloads",
  "--allow-root",
  "/Users/you/Documents"
]
enabled = true
startup_timeout_sec = 60
```

如果偏好非 scope 包名，也可以在 npm 版 Codex 配置中把 `@flyfish-dev/word-ai` 替换为 `word-ai-mcp`。

`--root` 是相对路径和 `.wordai` sidecar 的主工作区。需要编辑 Downloads、Documents 或团队目录中的原始 DOCX 时，重复添加 `--allow-root`。安装脚本生成的 `.wordai/codex-config.toml` 会自动加入常见用户文档目录。

推荐将写入类工具设置为需要审批，尤其是 `docx_apply_patchset`、`docx_restore_backup`、`docx_rollback`。
打开的 Word 文档会话写入也应审批：`word_session_apply_patchset`、`word_session_wrap_selection`、`word_session_rollback`。

对于当前已在 Word 中打开的文档，先加载 Office add-in 并连接 bridge，然后可以让 Codex 使用：

```text
使用 word_ai 列出 active Word sessions，读取 live session 中 WORD-AI:SRS:1.0:overview 的内容控件文本，先通过 Office.js preview PatchSet，再 apply 到当前打开文档，并返回 audit 与 rollback PatchSet。
```

## OfficeCLI 兼容策略

Word AI 可以把 OfficeCLI 作为辅助证据来源，但默认只允许只读/低风险能力：`view html`、`view screenshot`、`view issues`、`query --json`、`validate`。OfficeCLI 的 `set`、`add`、`remove`、`raw-set`、`batch`、`merge` 等写入/变更命令，不作为 Word AI 默认流程开放；除非未来被包进 Word AI 的 PatchSet、dry-run、audit、rollback 和显式审批机制。

MCP Server 只通过白名单 wrapper 暴露该集成：`officecli_view_html`、`officecli_view_screenshot`、`officecli_view_issues`、`officecli_query`、`officecli_validate`。如果本机未安装 OfficeCLI，这些工具会返回 `available=false`，核心 Word AI 流程仍继续使用 `docx_*` 和 `word_session_*`。

Word AI 会借鉴 OfficeCLI 的 schema/help-first、semantic path、watch/render、template merge、dump/batch 等设计，但正式写入模型仍以 Word AI PatchSet 为准。

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
- [Word AI Codex Skill](skills/word-ai/SKILL.md)
- [架构说明](docs/ARCHITECTURE.md)
- [英文架构说明](docs/ARCHITECTURE.en.md)
- [工具契约](docs/TOOL_CONTRACT.md)
- [安全设计](docs/SECURITY.md)
- [QA 验证报告](docs/QA_REPORT.md)
- [验证矩阵](docs/VALIDATION_MATRIX.md)
- [MCP Registry 发布说明](docs/REGISTRY_PUBLISHING.md)
- [v0.8.1 变更记录](docs/CHANGELOG_V081.md)
- [v0.8.0 变更记录](docs/CHANGELOG_V080.md)
- [v0.7.1 变更记录](docs/CHANGELOG_V071.md)
- [v0.7 变更记录](docs/CHANGELOG_V07.md)

## 许可证

GNU Affero General Public License v3.0 or later（`AGPL-3.0-or-later`）。详见 [LICENSE](LICENSE)。
