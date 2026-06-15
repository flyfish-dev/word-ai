# MCP Registry Publishing

Word AI is prepared for global discovery through the official MCP Registry and GitHub Container Registry (GHCR).

## Server Identity

- MCP server name: `io.github.flyfish-dev/word-ai`
- Source repository: `https://github.com/flyfish-dev/word-ai`
- Package type: MCPB
- Release asset: `https://github.com/flyfish-dev/word-ai/releases/download/v<version>/word-ai-<version>.mcpb`
- Transport: `stdio`
- License: `AGPL-3.0-or-later`

## Local Bundle And Container Smoke Tests

Build the deterministic MCPB bundle:

```bash
PYTHONPATH=. python scripts/build_mcpb.py --version 0.8.1 --out dist/word-ai-0.8.1.mcpb
shasum -a 256 dist/word-ai-0.8.1.mcpb
```

The MCPB bootstrap requires Python 3.10+ on the target machine. On first run it creates `.word-ai-mcpb-venv` inside the installed bundle directory, installs Word AI dependencies there, and then starts the stdio MCP server. Bootstrap installation logs are written to stderr so stdout remains reserved for MCP JSON-RPC.

Build the optional local container:

```bash
docker build -t word-ai:local .
docker run --rm -i \
  -v "$PWD:/workspace" \
  -v "$HOME/Downloads:/documents/Downloads" \
  word-ai:local
```

The default container command starts the MCP server with:

```bash
python -m word_ai_mcp.server --root /workspace --allow-root /documents
```

This keeps the MCP process local to the client while allowing users to mount document folders explicitly.

## Release Flow

1. Update `pyproject.toml`, `word_ai_mcp/__init__.py`, `server.json`, README, and changelog to the release version.
2. Merge to `main`.
3. Create and push a version tag:

   ```bash
   git tag v0.8.1
   git push origin v0.8.1
   ```

4. GitHub Actions builds `word-ai-<version>.mcpb` and computes its `fileSha256`.
5. The workflow uploads the MCPB file to the GitHub Release.
6. The workflow authenticates to the MCP Registry with GitHub OIDC.
7. `mcp-publisher validate` and `mcp-publisher publish` publish the version metadata from `server.json`.

## Verification

After the release workflow finishes:

```bash
curl "https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.flyfish-dev/word-ai"
curl "https://registry.modelcontextprotocol.io/v0.1/servers/io.github.flyfish-dev%2Fword-ai/versions/latest"
```

If the registry publish step fails with a namespace permission error, verify that the GitHub identity used by the workflow can publish for the `flyfish-dev` organization namespace, or switch the registry authentication method to DNS verification for a custom domain.

## 中文说明

Word AI 已按官方 MCP Registry 的元数据格式准备全球索引能力。项目通过 GitHub Release 发布公开 MCPB 资产，通过 `server.json` 声明 MCP server 名称、仓库、版本、安装包和 stdio transport。

正式发布时，只需要推送 `v*` tag。GitHub Actions 会完成 MCPB 构建、sha256 计算、GitHub Release 上传、GitHub OIDC 登录 MCP Registry 和 `mcp-publisher publish`。发布完成后，MCP Registry 及其下游聚合器即可按 `io.github.flyfish-dev/word-ai` 检索。
