# v0.8.3 Changelog - Cross-Platform Native Distribution

v0.8.3 makes native .NET Open XML distribution explicit and release-grade. The official MCPB and GitHub Release assets carry self-contained native backends for the mainstream desktop/server platforms, while the npm packages download the current-platform backend on first run. Python remains the MCP facade and Office.js bridge runtime.

## Added

- Added `scripts/publish_native_backends.py` to cross-publish self-contained `WordAi.OpenXml` binaries for:
  - `osx-arm64`
  - `osx-x64`
  - `linux-x64`
  - `linux-arm64`
  - `linux-musl-x64`
  - `linux-musl-arm64`
  - `win-x64`
  - `win-arm64`
- Added `scripts/verify_native_backends.py` to verify all expected native files, executable bits, file sizes, and current-platform execution.
- Added `scripts/package_native_assets.py` to produce per-RID GitHub Release assets and SHA-256 checksums.
- Added release workflow steps to build all native backends before MCPB packaging, package per-RID native assets, smoke-check lightweight npm payloads, and publish both `@flyfish-dev/word-ai` and `word-ai-mcp`.

## Changed

- The runtime now detects Linux glibc vs musl and searches the matching RID first, with a conservative Linux fallback.
- Native executable selection now supports `WORD_AI_DOTNET_RID` and `WORD_AI_DOTNET_NATIVE_DIR` overrides.
- The npm launcher now downloads only the current-platform native archive from GitHub Releases, verifies SHA-256, caches it, and leaves MCP stdout clean during bootstrap.
- MCPB packaging now preserves executable file modes and includes the split English/Chinese documentation.
- CI now verifies native packaging and MCPB payload contents for the Linux x64 path.

## 中文

v0.8.3 将 .NET Open XML native 后端的分发做成明确、可验证的发布链路。官方 MCPB 与 GitHub Release 资产携带主流桌面/服务器平台的 self-contained native 二进制；npm 包首次运行时按当前平台下载后端。Python 继续作为 MCP facade 与 Office.js bridge runtime。

## 新增

- 新增 `scripts/publish_native_backends.py`，交叉发布以下 RID 的 `WordAi.OpenXml` self-contained 二进制：
  - `osx-arm64`
  - `osx-x64`
  - `linux-x64`
  - `linux-arm64`
  - `linux-musl-x64`
  - `linux-musl-arm64`
  - `win-x64`
  - `win-arm64`
- 新增 `scripts/verify_native_backends.py`，校验所有 native 文件、可执行权限、文件大小以及当前平台可执行 smoke。
- 新增 `scripts/package_native_assets.py`，生成每个 RID 的 GitHub Release 下载资产和 SHA-256 校验文件。
- Release workflow 现在会先构建全部 native 后端，再构建 MCPB 包，生成每平台 native 资产，检查轻量 npm payload，并发布 `@flyfish-dev/word-ai` 与 `word-ai-mcp` 两个 npm 包。

## 变更

- 运行时会区分 Linux glibc 与 musl，优先加载匹配 RID，并提供保守 Linux fallback。
- Native executable 选择支持 `WORD_AI_DOTNET_RID` 和 `WORD_AI_DOTNET_NATIVE_DIR` 覆盖。
- npm 启动器只下载当前平台 GitHub Release native 压缩包，校验 SHA-256 后缓存，并保证 bootstrap 阶段不污染 MCP stdout。
- MCPB 打包会保留可执行权限，并包含拆分后的中英文文档。
- CI 增加 Linux x64 native packaging 和 MCPB payload 验证。
