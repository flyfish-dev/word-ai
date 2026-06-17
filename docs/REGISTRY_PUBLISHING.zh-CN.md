# MCP Registry 发布说明

| 文档 | 预览 |
| --- | --- |
| [中文](REGISTRY_PUBLISHING.zh-CN.md) | 全球 MCP Registry、MCPB、npm 第二渠道、发布流程和验证步骤。 |
| [English](REGISTRY_PUBLISHING.md) | Global MCP Registry, MCPB, npm fallback, release flow, and verification steps. |

Word AI 通过官方 MCP Registry 和公开 GitHub Release MCPB 资产实现全球发现。npm 包是第二便利渠道，不应替代面向 MCP host 的 Registry/MCPB 元数据。

## Server 身份

- MCP server 名称：`io.github.flyfish-dev/word-ai`
- 源码仓库：`https://github.com/flyfish-dev/word-ai`
- 包类型：MCPB
- Release 资产：`https://github.com/flyfish-dev/word-ai/releases/download/v<version>/word-ai-<version>.mcpb`
- Transport：`stdio`
- License：`AGPL-3.0-or-later`

## 本地 Bundle 与容器 Smoke Test

构建确定性的 MCPB bundle：

```bash
scripts/publish_dotnet.sh --all
PYTHONPATH=. python scripts/verify_native_backends.py --require-all
PYTHONPATH=. python scripts/package_native_assets.py --version 0.8.3
PYTHONPATH=. python scripts/build_mcpb.py --version 0.8.3 --out dist/word-ai-0.8.3.mcpb
shasum -a 256 dist/word-ai-0.8.3.mcpb
```

正式 release RID 集合为 `osx-arm64`、`osx-x64`、`linux-x64`、`linux-arm64`、`linux-musl-x64`、`linux-musl-arm64`、`win-x64` 和 `win-arm64`。如果构建 MCPB 前已经存在 `dist/native/<rid>/WordAi.OpenXml` 或 `WordAi.OpenXml.exe`，native .NET Open XML 后端会被打包，并优先于 DLL、项目和 Python fallback 被选择。npm 包故意不内嵌所有 native 后端；启动器会从 GitHub Release 下载当前平台压缩包，校验 SHA-256 后缓存到用户缓存目录，并通过 `WORD_AI_DOTNET_NATIVE_DIR` 暴露给运行时。MCPB bootstrap 仍要求 Python 3.10+ 运行 MCP facade；首次启动会在安装后的 bundle 目录内创建 `.word-ai-mcpb-venv`，安装 Word AI 依赖，然后启动 stdio MCP server。Bootstrap 安装日志写入 stderr，stdout 保留给 MCP JSON-RPC。

GitHub Release 还会包含每个 RID 的 native 压缩包：`word-ai-openxml-<version>-<rid>.tar.gz` 或 `.zip`，以及 `word-ai-openxml-<version>-checksums.sha256`。

构建可选本地容器：

```bash
docker build -t word-ai:local .
docker run --rm -i \
  -v "$PWD:/workspace" \
  -v "$HOME/Downloads:/documents/Downloads" \
  word-ai:local
```

默认容器命令使用：

```bash
python -m word_ai_mcp.server --root /workspace --allow-root /documents
```

这会让 MCP 进程保持本地运行，同时允许用户显式挂载文档目录。

## npm 第二渠道

npm 适合快速免 clone 启动，以及暂不能消费 MCP Registry/MCPB 包的客户端。用户文档应先推荐 Registry/MCPB，再把 npm 作为备用入口：

- 推荐 scoped 包：`@flyfish-dev/word-ai`
- 非 scope 兼容包：`word-ai-mcp`
- CLI 命令：`word-ai`、`word-ai-mcp`、`word-ai-http`
- Native 后端分发：npm 只下载当前 RID 的 GitHub Release 压缩包，校验 checksum 后缓存。需要安装后完全离线的环境可设置 `WORD_AI_SKIP_NATIVE_DOWNLOAD=1` 并自行提供 native 后端。

示例 smoke check：

```bash
npm exec --yes --package @flyfish-dev/word-ai -- word-ai --root "$PWD" doctor
npm exec --yes --package word-ai-mcp -- word-ai-mcp --root "$PWD"
```

## 发布流程

1. 将 `pyproject.toml`、`word_ai_mcp/__init__.py`、`server.json`、README 和 changelog 更新到发布版本。
2. 合并到 `main`。
3. 创建并推送版本 tag：

   ```bash
   git tag v0.8.3
   git push origin v0.8.3
   ```

4. GitHub Actions 为所有 release RID 发布 native .NET 后端，并验证当前平台二进制可执行。
5. 工作流构建 `word-ai-<version>.mcpb`、计算 `fileSha256`，并验证 npm payload 保持轻量且不内嵌过大的 native 二进制。
6. 工作流上传 MCPB、每平台 native 压缩包和 checksums 到 GitHub Release。
7. 工作流在对应版本尚不存在时发布 `@flyfish-dev/word-ai` 与 `word-ai-mcp` 到 npm。
8. 工作流通过 GitHub OIDC 登录 MCP Registry。
9. `mcp-publisher validate` 和 `mcp-publisher publish` 基于 `server.json` 发布版本元数据。

## 验证

发布工作流完成后执行：

```bash
curl "https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.flyfish-dev/word-ai"
curl "https://registry.modelcontextprotocol.io/v0.1/servers/io.github.flyfish-dev%2Fword-ai/versions/latest"
```

如果 registry 发布步骤因为 namespace 权限失败，请确认工作流使用的 GitHub 身份可以发布 `flyfish-dev` 组织命名空间，或切换为自定义域名的 DNS 验证方式。
