# Word AI

<!-- mcp-name: io.github.flyfish-dev/word-ai -->

| Language | Preview |
| --- | --- |
| [English](README.md) | Structure-preserving Word DOCX editing MCP server with a .NET Open XML backend and Office.js live sessions. |
| [中文](README.zh-CN.md) | 面向 AI Agent 的 Word DOCX 结构稳定编辑；支持 .NET 后端、Office.js 会话和安全 PatchSet。 |

**Structure-preserving Word (DOCX) editing for AI agents.**

Word AI is an open-source MCP server and Office.js bridge for safe, auditable, incremental editing of Microsoft Word documents. It is designed for Codex, OpenAI Agents, and other MCP clients that need to edit `.docx` files without rebuilding the document or damaging styles, numbering, tables, images, fields, headers, footers, and relationships.

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
- **Python MCP facade and bridge runtime** for local agent integration, path policy, session queues, and distribution compatibility.
- **.NET 8 Open XML SDK engine** as the authoritative offline DOCX transaction backend, using a packaged native binary or Release DLL when available.
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
        +--> Python MCP facade / read indexes / Office bridge
        +--> .NET Open XML SDK backend for offline PatchSet transactions
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

This installs the Python MCP facade, builds the Office.js taskpane, builds the .NET Open XML backend when .NET SDK 8 is available, writes `.wordai/codex-config.toml`, and installs the `word-ai` skill into Codex, Claude Code, and detected compatible agent clients.

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
PYTHONPATH=. .venv/bin/python scripts/run_outline_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_engine_selection_regression.py
```

Build the .NET engine:

```bash
dotnet --version  # requires .NET SDK 8
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
scripts/publish_dotnet.sh   # optional: native binary in dist/native/<rid>
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py
```

## Offline Engine Selection

Offline file transactions use the .NET Open XML backend by default when it is available. Selection order is:

1. `WORD_AI_DOTNET_EXE` or a packaged native executable under `native/<rid>/` or `dist/native/<rid>/`.
2. `WORD_AI_DOTNET_DLL` or the local Release DLL at `dotnet/WordAi.OpenXml/bin/Release/net8.0/WordAi.OpenXml.dll`.
3. Local source project via `dotnet run --project dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj`.
4. Python OOXML fallback only when .NET is unavailable and `WORD_AI_ENGINE=auto`.

Control it with `WORD_AI_ENGINE=auto|dotnet|python`, or pass `engine` to `docx_assess_patchset`, `docx_dry_run_patchset`, `docx_apply_patchset`, and `docx_validate`. Use `WORD_AI_ENGINE=dotnet` in production to fail fast instead of silently falling back.

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

Then sideload `office-addin/manifest.xml` in Word. The taskpane runs at `https://localhost:3100/taskpane.html` by default and proxies `/bridge/*` to the local bridge. The bridge prints a local token at startup. Use that token in the taskpane.

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
- [Architecture](docs/ARCHITECTURE.en.md)
- [Tool Contract](docs/TOOL_CONTRACT.md)
- [Security Design](docs/SECURITY.en.md)
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
