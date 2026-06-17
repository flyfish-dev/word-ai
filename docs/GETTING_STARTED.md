# Getting Started

| Language | Preview |
| --- | --- |
| [English](GETTING_STARTED.md) | Install, configure, run, validate, and connect Word AI with MCP Registry, Agent Skills, and local Office.js sessions. |
| [中文](GETTING_STARTED.zh-CN.md) | 安装、配置、启动、验证 Word AI，并连接 MCP Registry、Agent Skill 与本地 Office.js 会话。 |

## Requirements

- Python 3.10 or newer.
- .NET SDK 8 for the Open XML SDK engine.
- Node.js and npm for the Office.js taskpane.
- GitHub CLI is optional, only needed for repository publishing workflows.

## Install

Recommended installation order:

1. **MCP Registry / MCPB**: install server `io.github.flyfish-dev/word-ai` from your MCP host or marketplace when supported.
2. **Agent Skill**: install the `word-ai` Skill so Codex, Claude Code, and compatible agents follow the safe Word editing workflow.
3. **Local source install**: use this for full Office.js live Word sessions, the localhost bridge, .NET Open XML regression checks, and development.
4. **npm secondary channel**: use npm only as a no-clone fallback for clients that cannot consume MCP Registry/MCPB yet.

MCP Registry details:

- Server name: `io.github.flyfish-dev/word-ai`
- Registry metadata: [server.json](../server.json)
- MCPB package: `https://github.com/flyfish-dev/word-ai/releases/download/v0.8.1/word-ai-0.8.1.mcpb`
- Registry latest API: `https://registry.modelcontextprotocol.io/v0.1/servers/io.github.flyfish-dev%2Fword-ai/versions/latest`

Recommended one-command setup:

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

bash scripts/install.sh
```

This installs the Python MCP facade, builds the Office.js taskpane, builds the .NET Open XML backend when .NET SDK 8 is available, writes `.wordai/codex-config.toml`, and installs the formal `word-ai` skill into Codex, Claude Code, and detected compatible agent clients.

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

The installer writes the formal `word-ai` skill to the local discovery paths used by agent clients. Install the Skill even when the MCP server is installed through the MCP Registry, because it carries the offline/live mode selection rules and safety workflow:

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

## Secondary npm Channel

npm is a convenience fallback for clients that do not yet consume MCP Registry/MCPB packages, for CI smoke tests, and for quick no-clone stdio startup. Prefer MCP Registry/MCPB for MCP host discovery.

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

Secondary npm-based Codex config:

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

You can replace `@flyfish-dev/word-ai` with `word-ai-mcp` in this npm-based config. The first npm run creates a Python virtual environment under the user cache and installs the Word AI Python dependencies automatically. Set `WORD_AI_PYTHON=/path/to/python3.10+` if Python discovery needs help.

## Run Local Checks

```bash
.venv/bin/word-ai --root "$PWD" doctor
PYTHONPATH=. .venv/bin/python -m compileall word_ai_mcp scripts
PYTHONPATH=. .venv/bin/python scripts/run_smoke_test.py
PYTHONPATH=. .venv/bin/python scripts/run_structure_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_outline_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_engine_selection_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_word_session_smoke.py
PYTHONPATH=. .venv/bin/python scripts/validate_word_ai_skill.py
```

Build and test the .NET engine:

```bash
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
scripts/publish_dotnet.sh
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py
```

## Offline Engine Selection

Word AI does not require Python to implement the production DOCX writer. Python is the MCP facade and Office.js bridge runtime. Offline file transactions use the .NET Open XML backend by default when available:

1. `WORD_AI_DOTNET_EXE` or packaged native executable in `native/<rid>/` or `dist/native/<rid>/`.
2. `WORD_AI_DOTNET_DLL` or local Release DLL.
3. Source project through `dotnet run --project`.
4. Python OOXML fallback only when `WORD_AI_ENGINE=auto` and no .NET backend is available.

Set `WORD_AI_ENGINE=dotnet` in production to fail fast if the .NET backend is missing. Use `WORD_AI_ENGINE=python` only for fallback comparison or development.

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

Sideload `office-addin/manifest.xml` in Word. The default live-session taskpane URL is `https://localhost:3100/taskpane.html`. For browser-only debugging:

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
