# 快速上手

| 文档 | 预览 |
| --- | --- |
| [中文](GETTING_STARTED.zh-CN.md) | 安装、配置、启动、验证 Word AI，并连接 MCP Registry、Agent Skill 与本地 Office.js 会话。 |
| [English](GETTING_STARTED.md) | Install, configure, run, validate, and connect Word AI with MCP Registry, Agent Skills, and local Office.js sessions. |

## 环境要求

- Python 3.10 或更高版本。
- .NET SDK 8，用于 Open XML SDK 引擎。
- Node.js 和 npm，用于 Office.js taskpane。

## 安装

推荐安装顺序：

1. **MCP Registry / MCPB**：客户端支持时，优先从 MCP host 或 marketplace 安装 `io.github.flyfish-dev/word-ai`。
2. **Agent Skill**：安装 `word-ai` Skill，让 Codex、Claude Code 和兼容 Agent 遵循安全 Word 编辑流程。
3. **本地源码安装**：需要完整 Office.js live Word session、localhost bridge、.NET Open XML 回归或开发调试时使用。
4. **npm 第二渠道**：仅作为暂不支持 MCP Registry/MCPB 的客户端、CI smoke test 或免 clone stdio 启动的备用方式。

MCP Registry 信息：

- MCP server 名称：`io.github.flyfish-dev/word-ai`
- Registry 元数据：[server.json](../server.json)
- MCPB 包：`https://github.com/flyfish-dev/word-ai/releases/download/v0.8.4/word-ai-0.8.4.mcpb`
- Registry latest API：`https://registry.modelcontextprotocol.io/v0.1/servers/io.github.flyfish-dev%2Fword-ai/versions/latest`

推荐一键安装：

```bash
git clone https://github.com/flyfish-dev/word-ai.git
cd word-ai

bash scripts/install.sh
```

该命令会安装 Python MCP facade、构建 Office.js taskpane、在可用时构建 .NET Open XML 后端、生成 `.wordai/codex-config.toml`，并把正式 `word-ai` Skill 安装到 Codex、Claude Code 以及已检测到的兼容 Agent 客户端。

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

## Agent Skill 一键安装

安装器会把正式 `word-ai` Skill 写入 Agent 客户端的本地识别目录。即使 MCP Server 已经通过 MCP Registry 安装，也建议安装 Skill，因为它携带离线/实时模式选择规则和安全流程：

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

## npm 第二渠道

npm 是便利备用渠道，适合暂不支持 MCP Registry/MCPB 的客户端、CI smoke test、或免 clone 启动 stdio server。MCP host 发现能力优先使用 MCP Registry/MCPB。

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

npm 版 Codex 配置：

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

也可以把该 npm 配置中的 `@flyfish-dev/word-ai` 替换为 `word-ai-mcp`。首次 npm 启动会在用户缓存目录创建 Python venv，并自动安装 Word AI Python 依赖。如需指定 Python，可设置 `WORD_AI_PYTHON=/path/to/python3.10+`。

## 本地验证

```bash
.venv/bin/word-ai --root "$PWD" doctor
PYTHONPATH=. .venv/bin/python -m compileall word_ai_mcp scripts
PYTHONPATH=. .venv/bin/python scripts/run_smoke_test.py
PYTHONPATH=. .venv/bin/python scripts/run_structure_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_outline_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_engine_selection_regression.py
PYTHONPATH=. .venv/bin/python scripts/run_word_session_smoke.py
PYTHONPATH=. .venv/bin/python scripts/validate_word_ai_skill.py
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
scripts/publish_dotnet.sh           # 当前主机 RID
scripts/publish_dotnet.sh --all     # 所有支持的 release RID
PYTHONPATH=. .venv/bin/python scripts/run_dotnet_regression.py
```

## 离线引擎选择

Word AI 并不是要求 Python 实现生产级 DOCX writer。Python 是 MCP facade 与 Office.js bridge runtime；离线文件事务默认优先使用 .NET Open XML 后端：

1. `WORD_AI_DOTNET_EXE` 或 `native/<rid>/`、`dist/native/<rid>/` 下的 native executable。
2. `WORD_AI_DOTNET_DLL` 或本地 Release DLL。
3. 通过 `dotnet run --project` 使用源码工程。
4. 只有在 `WORD_AI_ENGINE=auto` 且 .NET 后端不可用时，才回退 Python OOXML。

MCPB 与 GitHub Release 资产包含 `osx-arm64`、`osx-x64`、`linux-x64`、`linux-arm64`、`linux-musl-x64`、`linux-musl-arm64`、`win-x64` 和 `win-arm64`。Word AI 会自动选择当前平台。npm 启动器只会从 GitHub Release 下载当前平台 native 压缩包，校验 SHA-256 后缓存到用户缓存目录；只有自定义打包时才需要设置 `WORD_AI_DOTNET_RID`、`WORD_AI_DOTNET_EXE` 或 `WORD_AI_DOTNET_NATIVE_DIR`。

生产环境建议设置 `WORD_AI_ENGINE=dotnet`，让 .NET 后端缺失时直接失败。`WORD_AI_ENGINE=python` 只建议用于 fallback 对照或开发调试。

## Codex 配置

先查看安装脚本生成的配置：

```bash
cat .wordai/codex-config.toml
```

将其中的 `mcp_servers.word_ai` 配置加入 Codex 配置文件。安装脚本生成的配置会在存在时自动加入 `~/Downloads`、`~/Documents` 和 `~/Desktop`。如果还要访问其他目录，可使用：

```bash
.venv/bin/word-ai --root "$PWD" codex-config --allow-root "/path/to/team/docs" --output .wordai/codex-config.toml
```

新增 MCP server 或修改 `--allow-root` 后，通常需要新开 Codex 会话或重启 Codex 才会加载。

## Office Bridge

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

在 Word 中加载 `office-addin/manifest.xml`，默认 live-session taskpane 地址为 `https://localhost:3100/taskpane.html`，并在 taskpane 中填入 bridge 启动时打印的 token。

浏览器调试只能验证 taskpane UI 和 bridge 连接，不能真正执行 Word host 的 `Word.run(...)`。要让 Codex 编辑当前打开文档，必须在 Microsoft Word 中加载 taskpane，连接 bridge，确认出现 `Live session`，然后让 Codex 调用 `word_session_*` 工具。

## 可选 OfficeCLI 证据

OfficeCLI 不是 Word AI 核心流程的必装依赖。如果本机安装了 `officecli`，Word AI 只开放以下白名单辅助 wrapper：

- `officecli_view_html`
- `officecli_view_screenshot`
- `officecli_view_issues`
- `officecli_query`
- `officecli_validate`

这些工具只用于渲染、issues、query、validate 等辅助证据。默认编辑路径仍是 Word AI PatchSet，不使用 OfficeCLI 写入/变更命令。
