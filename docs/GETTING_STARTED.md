# Getting Started / 快速上手

## Requirements

- Python 3.10 or newer.
- .NET SDK 8 for the Open XML SDK engine.
- Node.js and npm for the Office.js taskpane.
- GitHub CLI is optional, only needed for repository publishing workflows.

## Install

No-clone npm quick check:

```bash
npx -y @flyfish-dev/word-ai --root "$PWD" doctor
```

Run the MCP stdio server through npm:

```bash
npm exec --yes --package @flyfish-dev/word-ai -- word-ai-mcp --root "$PWD" --allow-root "$HOME/Downloads"
```

For Codex without cloning the repository:

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

The first npm run creates a Python virtual environment under the user cache and installs the Word AI Python dependencies automatically. Set `WORD_AI_PYTHON=/path/to/python3.10+` if Python discovery needs help.

Recommended one-command setup:

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

bash scripts/install.sh
```

This installs the Python MCP package, builds the Office.js taskpane, builds the .NET Open XML engine when .NET SDK 8 is available, writes `.wordai/codex-config.toml`, and installs the formal `word-ai` skill into Codex, Claude Code, and detected compatible agent clients.

Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\install.ps1
```

Manual setup:

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Agent Skill Auto-Install

The installer writes the same formal `word-ai` skill to the local discovery paths used by agent clients:

- Codex official user skills: `~/.agents/skills/word-ai`
- Codex app compatibility skills: `~/.codex/skills/word-ai`
- Claude Code personal skills: `~/.claude/skills/word-ai`
- Cursor, Windsurf, GitHub Copilot, and OpenClaw when their skill homes already exist

To install or refresh only the skills:

```bash
python3 scripts/install_agent_skills.py
```

To install into all known client folders or add repo-scoped skills:

```bash
python3 scripts/install_agent_skills.py --agents all
python3 scripts/install_agent_skills.py --project
```

Start a new agent session after installation. If a client does not show the skill immediately, restart that client.

## Run Local Checks

```bash
.venv/bin/word-ai --root "$PWD" doctor
PYTHONPATH=. .venv/bin/python -m compileall word_ai_mcp scripts
PYTHONPATH=. .venv/bin/python scripts/run_smoke_test.py
PYTHONPATH=. .venv/bin/python scripts/run_structure_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_word_session_smoke.py
PYTHONPATH=. .venv/bin/python scripts/validate_word_ai_skill.py
```

Build and test the .NET engine:

```bash
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py
```

Build and validate the Office add-in:

```bash
cd office-addin
npm install
npm run build
npx office-addin-manifest validate manifest.xml
```

## Run The MCP Server

```bash
PYTHONPATH=. .venv/bin/python -m word_ai_mcp.server --root "$PWD"
```

The server uses stdio transport and is meant to be launched by Codex or another MCP client.

## Configure Codex

The installer writes a ready-to-merge config snippet:

```bash
cat .wordai/codex-config.toml
```

Add it to your Codex config. The generated snippet includes write-tool approvals. A minimal manual config is:

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

Use `--allow-root` for external folders that contain source DOCX files. The installer-generated config includes `~/Downloads`, `~/Documents`, and `~/Desktop` when those folders exist. You can add more roots with:

```bash
.venv/bin/word-ai --root "$PWD" codex-config --allow-root "/path/to/team/docs" --output .wordai/codex-config.toml
```

For production-like use, require approval on all write tools:

- `docx_preflight_patchset`
- `docx_dry_run_patchset`
- `docx_backup`
- `docx_apply_patchset`
- `docx_restore_backup`
- `docx_rollback`
- `word_session_apply_patchset`
- `word_session_wrap_selection`
- `word_session_rollback`
- sidecar export tools

## Run The Office Bridge

Recommended:

```bash
bash scripts/start.sh
```

This starts both the local bridge and the Office.js taskpane. Use `bash scripts/start.sh --http` for browser-only taskpane debugging.

Windows PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\start.ps1
```

Manual bridge startup:

```bash
PYTHONPATH=. .venv/bin/python -m word_ai_mcp.server_http --root "$PWD" --host 127.0.0.1 --port 8765
```

The bridge prints a local token. Put that token into the Word taskpane.

Run the taskpane:

```bash
cd office-addin
npm run dev
```

Sideload `office-addin/manifest.xml` in Word. For browser-only debugging:

```bash
npm run dev:http
```

Browser-only debugging can verify the taskpane UI and bridge connectivity, but it cannot execute real `Word.run(...)` Office.js operations. For live document editing, the taskpane must be loaded in Microsoft Word.

## Optional OfficeCLI Evidence

OfficeCLI is not required for the core Word AI workflow. If `officecli` is installed, Word AI exposes only allowlisted auxiliary wrappers:

- `officecli_view_html`
- `officecli_view_screenshot`
- `officecli_view_issues`
- `officecli_query`
- `officecli_validate`

Use these for rendering, issues, query, and validation evidence after the Word AI PatchSet validation path. Do not use OfficeCLI mutation commands as the default editing path.

## Edit The Open Word Document Through Codex

1. Start `word_ai_mcp.server_http`.
2. Start the Office taskpane dev server.
3. Sideload `office-addin/manifest.xml` in Microsoft Word.
4. Open a DOCX that contains content controls, or select text and use `Wrap Selection`.
5. Connect the bridge in the taskpane and confirm the session status shows `Live session`.
6. In Codex, use:

```text
Use word_ai to call word_session_list, read WORD-AI:SRS:1.0:overview from the active Word session, preview a PatchSet, apply it through Office.js, and return the audit plus rollback PatchSet.
```

The live path uses these MCP tools:

- `word_session_list`
- `word_session_snapshot`
- `word_session_read_content_control`
- `word_session_preview_patchset`
- `word_session_apply_patchset`
- `word_session_rollback`

`word_session_apply_patchset` only supports content-control text operations and performs a live preflight before writing.

## 中文快速上手

### 环境要求

- Python 3.10 或更高版本。
- .NET SDK 8，用于 Open XML SDK 引擎。
- Node.js 和 npm，用于 Office.js taskpane。

### 安装

npm 免 clone 快速检查：

```bash
npx -y @flyfish-dev/word-ai --root "$PWD" doctor
```

通过 npm 运行 MCP stdio server：

```bash
npm exec --yes --package @flyfish-dev/word-ai -- word-ai-mcp --root "$PWD" --allow-root "$HOME/Downloads"
```

Codex 也可以不 clone 仓库，直接通过 npm 启动：

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

首次 npm 启动会在用户缓存目录创建 Python venv，并自动安装 Word AI Python 依赖。如需指定 Python，可设置 `WORD_AI_PYTHON=/path/to/python3.10+`。

推荐一键安装：

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

bash scripts/install.sh
```

该命令会安装 Python MCP 包、构建 Office.js taskpane、在可用时构建 .NET Open XML 引擎、生成 `.wordai/codex-config.toml`，并把正式 `word-ai` Skill 安装到 Codex、Claude Code 以及已检测到的兼容 Agent 客户端。

Windows PowerShell：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\install.ps1
```

手动安装：

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Agent Skill 一键安装

默认安装会自动写入这些本地识别目录：

- Codex 官方用户 skill：`~/.agents/skills/word-ai`
- Codex app 兼容 skill：`~/.codex/skills/word-ai`
- Claude Code 个人 skill：`~/.claude/skills/word-ai`
- 如果 Cursor、Windsurf、GitHub Copilot、OpenClaw 的 skill 目录已存在，也会同步安装

只安装或刷新 skill：

```bash
python3 scripts/install_agent_skills.py
```

安装到所有已知客户端目录，或额外安装到项目级目录：

```bash
python3 scripts/install_agent_skills.py --agents all
python3 scripts/install_agent_skills.py --project
```

安装后新开 Agent 会话；如果客户端没有立即显示，重启客户端即可。

### 本地验证

```bash
.venv/bin/word-ai --root "$PWD" doctor
PYTHONPATH=. .venv/bin/python -m compileall word_ai_mcp scripts
PYTHONPATH=. .venv/bin/python scripts/run_smoke_test.py
PYTHONPATH=. .venv/bin/python scripts/run_structure_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_word_session_smoke.py
PYTHONPATH=. .venv/bin/python scripts/validate_word_ai_skill.py
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py
```

### Codex 配置

先查看安装脚本生成的配置：

```bash
cat .wordai/codex-config.toml
```

将其中的 `mcp_servers.word_ai` 配置加入 Codex 配置文件。安装脚本生成的配置会在存在时自动加入 `~/Downloads`、`~/Documents` 和 `~/Desktop`。如果还要访问其他目录，可使用：

```bash
.venv/bin/word-ai --root "$PWD" codex-config --allow-root "/path/to/team/docs" --output .wordai/codex-config.toml
```

新增 MCP server 或修改 `--allow-root` 后，通常需要新开 Codex 会话或重启 Codex 才会加载。

### Office Bridge

启动本地 bridge：

```bash
bash scripts/start.sh
```

上面的命令会同时启动 bridge 和 taskpane。若只做浏览器调试，可以使用 `bash scripts/start.sh --http`。

Windows PowerShell：

```powershell
powershell -ExecutionPolicy Bypass -File scripts\start.ps1
```

手动启动 bridge：

```bash
PYTHONPATH=. .venv/bin/python -m word_ai_mcp.server_http --root "$PWD" --host 127.0.0.1 --port 8765
```

启动 taskpane：

```bash
cd office-addin
npm run dev
```

在 Word 中加载 `office-addin/manifest.xml`，并在 taskpane 中填入 bridge 启动时打印的 token。

浏览器调试只能验证 taskpane UI 和 bridge 连接，不能真正执行 Word host 的 `Word.run(...)`。要让 Codex 编辑当前打开文档，必须在 Microsoft Word 中加载 taskpane，连接 bridge，确认出现 `Live session`，然后让 Codex 调用 `word_session_*` 工具。

### 可选 OfficeCLI 证据

OfficeCLI 不是 Word AI 核心流程的必装依赖。如果本机安装了 `officecli`，Word AI 只开放以下白名单辅助 wrapper：

- `officecli_view_html`
- `officecli_view_screenshot`
- `officecli_view_issues`
- `officecli_query`
- `officecli_validate`

这些工具只用于渲染、issues、query、validate 等辅助证据。默认编辑路径仍是 Word AI PatchSet，不使用 OfficeCLI 写入/变更命令。
