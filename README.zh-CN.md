# Word AI

<!-- mcp-name: io.github.flyfish-dev/word-ai -->

| 文档 | 预览 |
| --- | --- |
| [中文](README.zh-CN.md) | 面向 AI Agent 的 Word DOCX 结构稳定编辑；支持 .NET 后端、Office.js 会话和安全 PatchSet。 |
| [English](README.md) | Structure-preserving Word DOCX editing MCP server with a .NET Open XML backend and Office.js live sessions. |

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
- **Python MCP facade 与 bridge runtime**，负责本地 Agent 集成、路径策略、session 队列和分发兼容。
- **.NET 8 Open XML SDK 引擎**，作为离线 DOCX 事务的权威后端；优先使用打包 native 二进制或 Release DLL。
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

安装脚本会安装 Python MCP facade 依赖、构建 Office.js taskpane、在可用时构建 .NET Open XML 后端、生成 `.wordai/codex-config.toml`，并把正式 `word-ai` Skill 安装到 Codex、Claude Code 以及已检测到的兼容 Agent 客户端。

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
PYTHONPATH=. .venv/bin/python scripts/run_outline_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_engine_selection_regression.py
```

构建 .NET 引擎：

```bash
dotnet --version  # 需要 .NET SDK 8
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
scripts/publish_dotnet.sh   # 可选：生成 dist/native/<rid> native 二进制
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py
```

## 离线引擎选择

离线文件事务默认优先使用 .NET Open XML 后端。选择顺序：

1. `WORD_AI_DOTNET_EXE` 或 `native/<rid>/`、`dist/native/<rid>/` 下的 native executable。
2. `WORD_AI_DOTNET_DLL` 或本地 Release DLL：`dotnet/WordAi.OpenXml/bin/Release/net8.0/WordAi.OpenXml.dll`。
3. 本地源码工程：`dotnet run --project dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj`。
4. 只有在 `WORD_AI_ENGINE=auto` 且 .NET 不可用时，才回退到 Python OOXML。

可通过 `WORD_AI_ENGINE=auto|dotnet|python` 控制，也可在 `docx_assess_patchset`、`docx_dry_run_patchset`、`docx_apply_patchset`、`docx_validate` 调用中传入 `engine`。生产环境建议设置 `WORD_AI_ENGINE=dotnet`，让后端缺失时直接失败，而不是静默回退。

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
- 发布说明：[MCP Registry 发布说明](docs/REGISTRY_PUBLISHING.zh-CN.md)

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

- [文档导航](docs/README.zh-CN.md)
- [快速上手](docs/GETTING_STARTED.zh-CN.md)
- [Word AI Codex Skill](skills/word-ai/SKILL.md)
- [架构说明](docs/ARCHITECTURE.md)
- [英文架构说明](docs/ARCHITECTURE.en.md)
- [工具契约](docs/TOOL_CONTRACT.md)
- [安全设计](docs/SECURITY.md)
- [QA 验证报告](docs/QA_REPORT.md)
- [验证矩阵](docs/VALIDATION_MATRIX.md)
- [MCP Registry 发布说明](docs/REGISTRY_PUBLISHING.zh-CN.md)
- [v0.8.1 变更记录](docs/CHANGELOG_V081.md)
- [v0.8.0 变更记录](docs/CHANGELOG_V080.md)
- [v0.7.1 变更记录](docs/CHANGELOG_V071.md)
- [v0.7 变更记录](docs/CHANGELOG_V07.md)

## 许可证

GNU Affero General Public License v3.0 or later（`AGPL-3.0-or-later`）。详见 [LICENSE](LICENSE)。
