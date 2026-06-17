# 分发说明

| 文档 | 预览 |
| --- | --- |
| [中文](DISTRIBUTION.zh-CN.md) | 发布渠道、单文件二进制、快速集成包、Agent Skill 安装和 npm 备用渠道。 |
| [English](DISTRIBUTION.md) | Release channels, standalone binaries, quickstart bundles, Agent Skill install, and npm fallback. |

Word AI 有两个主分发路径和一个备用路径：

1. **MCP Registry / MCPB**：用于 MCP host 发现、标准元数据和校验。
2. **Standalone quickstart 包**：用于不想 clone 源码、不想安装 Python venv/pip/.NET SDK 的用户，解压后只有一个本地命令。
3. **npm**：第二渠道，给暂不支持 MCP Registry/MCPB 的客户端或 CI 使用。

## 推荐路径

支持 MCP Registry 的客户端，优先安装 `io.github.flyfish-dev/word-ai`。这样 server identity、transport、version 和 package checksum 都由标准元数据承载。

Codex、Claude Code 和兼容 Agent 客户端建议同时安装 `word-ai` Agent Skill。Skill 会明确告诉 Agent：什么时候使用离线 `docx_*`，什么时候使用 live `word_session_*`，以及为什么默认只把 OfficeCLI 当作只读证据来源。

如果用户只需要快速得到一个本地命令，从 GitHub Release 下载当前平台 quickstart 包：

```bash
tar -xzf word-ai-quickstart-0.8.5-osx-arm64.tar.gz
cd word-ai-quickstart-0.8.5-osx-arm64

./word-ai install-skill
./word-ai codex-config --output .wordai/codex-config.toml
./word-ai mcp --root "$PWD" --allow-root "$HOME/Downloads" --allow-root "$HOME/Documents"
```

Windows 使用 `.zip` 包和 `word-ai.exe`：

```powershell
.\word-ai.exe install-skill
.\word-ai.exe codex-config --output .wordai\codex-config.toml
.\word-ai.exe mcp --root "$PWD" --allow-root "$HOME\Documents"
```

## 单文件二进制约定

standalone executable 内置：

- Python MCP facade 和 Python runtime 依赖。
- 当前平台 .NET Open XML 后端。
- PatchSet schemas。
- 官方 `word-ai` Agent Skill 模板。
- quickstart、doctor、Codex config、stdio MCP 和 HTTP bridge 入口。

支持命令：

```bash
word-ai --version
word-ai doctor
word-ai install-skill
word-ai codex-config --output .wordai/codex-config.toml
word-ai mcp --root "$PWD"
word-ai http --root "$PWD" --host 127.0.0.1 --port 8765
```

离线 DOCX 编辑不要求本地 Python、pip install、.NET SDK 或源码 checkout。完整 Office.js taskpane 开发仍使用源码安装路径，因为 Word 加载项需要在本机构建和 sideload。

## Release 资产

每个版本发布：

- `word-ai-<version>.mcpb`，用于 MCP Registry/MCPB 安装。
- .NET Open XML native 后端：`osx-arm64`、`osx-x64`、`linux-x64`、`linux-arm64`、`linux-musl-x64`、`linux-musl-arm64`、`win-x64`、`win-arm64`。
- 标准 hosted 平台 standalone 单文件：`linux-x64`、`linux-arm64`、`osx-arm64`、`osx-x64`、`win-x64`、`win-arm64`。
- 每个 standalone 对应的 quickstart 包，包含 README、manifest 和 LICENSE。
- SHA-256 checksum 文件。

MCPB 会包含所有 native Open XML 后端。standalone 每个平台一个 artifact。npm 包保持轻量，首次运行只下载当前平台 native 后端。

## 本地构建

构建当前平台 standalone：

```bash
python -m pip install -r requirements.txt -r requirements-standalone.txt
PYTHONPATH=. python scripts/publish_native_backends.py "$(PYTHONPATH=. python - <<'PY'
from scripts.publish_native_backends import host_rid
print(host_rid())
PY
)"
PYTHONPATH=. python scripts/build_standalone.py --json
```

打 quickstart 包：

```bash
PYTHONPATH=. python scripts/package_quickstart.py --version 0.8.5 --rid osx-arm64 --json
```

烟测：

```bash
dist/standalone/osx-arm64/word-ai --version
dist/standalone/osx-arm64/word-ai install-skill --dry-run
dist/standalone/osx-arm64/word-ai codex-config --output /tmp/word-ai-codex.toml
```

## 发布

发布由 tag 触发：

```bash
git tag v0.8.5
git push origin v0.8.5
```

Release workflow 会构建并烟测 MCPB、native 后端、standalone 二进制、quickstart 包、Docker、npm payload 和 MCP Registry 元数据。npm 发布需要 trusted publishing 或有效的 `NPM_TOKEN`；如果 npm 权限未配置，workflow 会 warning 后继续，保证 MCP Registry 与 GitHub Release 仍可用。
