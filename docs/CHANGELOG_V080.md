# v0.8.0 Changelog - Strict License And Global MCP Distribution

v0.8.0 prepares Word AI for stricter open-source distribution and official MCP Registry indexing.

## Added

- Added `server.json` for official MCP Registry metadata under `io.github.flyfish-dev/word-ai`.
- Added a Docker/OCI packaging path with MCP ownership label `io.modelcontextprotocol.server.name`.
- Added GitHub Actions release workflow for GHCR image publishing and MCP Registry publishing through `mcp-publisher` with GitHub OIDC.
- Added registry publishing documentation and container smoke-test instructions.

## Changed

- Changed the project license from MIT to `AGPL-3.0-or-later`.
- Bumped package and server version to `0.8.0`.

## 中文

v0.8.0 面向更严格的开源分发和全球 MCP 索引能力。

## 新增

- 新增 `server.json`，使用官方 MCP Registry 元数据格式声明 `io.github.flyfish-dev/word-ai`。
- 新增 Docker/OCI 发布路径，并加入 MCP 所有权验证标签 `io.modelcontextprotocol.server.name`。
- 新增 GitHub Actions 发布工作流，通过 GHCR 发布镜像，并使用 `mcp-publisher` + GitHub OIDC 发布到 MCP Registry。
- 新增 Registry 发布说明和容器 smoke test 文档。

## 变更

- 项目许可证从 MIT 切换为 `AGPL-3.0-or-later`。
- 版本提升到 `0.8.0`。
