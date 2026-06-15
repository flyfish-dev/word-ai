# Getting Started / 快速上手

## Requirements

- Python 3.10 or newer.
- .NET SDK 8 for the Open XML SDK engine.
- Node.js and npm for the Office.js taskpane.
- GitHub CLI is optional, only needed for repository publishing workflows.

## Install

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run Local Checks

```bash
PYTHONPATH=. python -m compileall word_ai_mcp scripts
PYTHONPATH=. python scripts/run_smoke_test.py
PYTHONPATH=. python scripts/run_structure_regression.py
PYTHONPATH=. python scripts/run_word_session_smoke.py
```

Build and test the .NET engine:

```bash
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. python scripts/run_dotnet_regression.py
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

Add this to your Codex config:

```toml
[mcp_servers.word_ai]
command = "/absolute/path/to/word-ai/.venv/bin/python"
args = ["-m", "word_ai_mcp.server", "--root", "/absolute/path/to/word-ai"]
enabled = true
startup_timeout_sec = 30

[mcp_servers.word_ai.env]
PYTHONPATH = "/absolute/path/to/word-ai"
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

```bash
PYTHONPATH=. python -m word_ai_mcp.server_http --root "$PWD" --host 127.0.0.1 --port 8765
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

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 本地验证

```bash
PYTHONPATH=. python -m compileall word_ai_mcp scripts
PYTHONPATH=. python scripts/run_smoke_test.py
PYTHONPATH=. python scripts/run_structure_regression.py
PYTHONPATH=. python scripts/run_word_session_smoke.py
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
PYTHONPATH=. python scripts/run_dotnet_regression.py
```

### Codex 配置

将上面的 `mcp_servers.word_ai` 配置加入 Codex 配置文件。新增 MCP server 后，通常需要新开 Codex 会话或重启 Codex 才会加载。

### Office Bridge

启动本地 bridge：

```bash
PYTHONPATH=. python -m word_ai_mcp.server_http --root "$PWD" --host 127.0.0.1 --port 8765
```

启动 taskpane：

```bash
cd office-addin
npm run dev
```

在 Word 中加载 `office-addin/manifest.xml`，并在 taskpane 中填入 bridge 启动时打印的 token。

浏览器调试只能验证 taskpane UI 和 bridge 连接，不能真正执行 Word host 的 `Word.run(...)`。要让 Codex 编辑当前打开文档，必须在 Microsoft Word 中加载 taskpane，连接 bridge，确认出现 `Live session`，然后让 Codex 调用 `word_session_*` 工具。
