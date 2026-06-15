# MCP Registry Publishing

Word AI is prepared for global discovery through the official MCP Registry and GitHub Container Registry (GHCR).

## Server Identity

- MCP server name: `io.github.flyfish-dev/word-ai`
- Source repository: `https://github.com/flyfish-dev/word-ai`
- Package type: OCI image
- Image: `ghcr.io/flyfish-dev/word-ai:<version>`
- Transport: `stdio`
- License: `AGPL-3.0-or-later`

## Local Container Smoke Test

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
   git tag v0.8.0
   git push origin v0.8.0
   ```

4. GitHub Actions builds and pushes `ghcr.io/flyfish-dev/word-ai:<version>`.
5. The workflow authenticates to the MCP Registry with GitHub OIDC.
6. `mcp-publisher publish` publishes the version metadata from `server.json`.

## Verification

After the release workflow finishes:

```bash
curl "https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.flyfish-dev/word-ai"
curl "https://registry.modelcontextprotocol.io/v0.1/servers/io.github.flyfish-dev%2Fword-ai/versions/latest"
```

If the registry publish step fails with a namespace permission error, verify that the GitHub identity used by the workflow can publish for the `flyfish-dev` organization namespace, or switch the registry authentication method to DNS verification for a custom domain.

## 中文说明

Word AI 已按官方 MCP Registry 的元数据格式准备全球索引能力。项目通过 GHCR 发布 OCI 镜像，通过 `server.json` 声明 MCP server 名称、仓库、版本、安装包和 stdio transport。

正式发布时，只需要推送 `v*` tag。GitHub Actions 会完成镜像构建、GHCR 推送、GitHub OIDC 登录 MCP Registry 和 `mcp-publisher publish`。发布完成后，MCP Registry 及其下游聚合器即可按 `io.github.flyfish-dev/word-ai` 检索。
