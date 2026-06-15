# v0.8.1 Changelog - MCPB Registry Release Path

v0.8.1 switches the official MCP Registry package path from GHCR OCI to a public GitHub Release MCPB artifact. This avoids organization-level GHCR visibility friction and gives users a more one-click-friendly local MCP server package.

## Added

- Added deterministic `scripts/build_mcpb.py` for reproducible `.mcpb` bundle creation.
- Added `mcpb_bootstrap.py`, which creates a local virtual environment on first run, installs Word AI dependencies, and preserves clean MCP stdout by routing installation logs to stderr.
- Added release workflow steps to build `word-ai-<version>.mcpb`, compute `fileSha256`, upload it to the GitHub Release, validate `server.json`, and publish to the MCP Registry.

## Changed

- Bumped package and server version to `0.8.1`.
- Changed `server.json` to use `registryType: "mcpb"` as the default global Registry package.
- Kept Dockerfile support for local/self-hosted builds, without relying on GHCR visibility for Registry publication.

## 中文

v0.8.1 将官方 MCP Registry 发布包从 GHCR OCI 切换为公开 GitHub Release MCPB 资产，避免组织级 GHCR 可见性限制，同时提供更接近一键安装的本地 MCP Server 分发格式。

## 新增

- 新增确定性 `scripts/build_mcpb.py`，用于生成可复现的 `.mcpb` bundle。
- 新增 `mcpb_bootstrap.py`，首次运行创建本地虚拟环境、安装 Word AI 依赖，并把安装日志写入 stderr，保持 MCP stdout JSON 干净。
- 发布 workflow 会构建 `word-ai-<version>.mcpb`、计算 `fileSha256`、上传 GitHub Release、验证 `server.json`，并发布到 MCP Registry。

## 变更

- 版本提升到 `0.8.1`。
- `server.json` 默认使用 `registryType: "mcpb"` 作为全球 Registry 发布包。
- Dockerfile 继续用于本地/自托管构建，但不再依赖 GHCR 可见性完成 Registry 发布。
