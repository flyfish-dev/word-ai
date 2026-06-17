# MCP Registry Publishing

| Language | Preview |
| --- | --- |
| [English](REGISTRY_PUBLISHING.md) | Global MCP Registry, MCPB, npm fallback, release flow, and verification steps. |
| [中文](REGISTRY_PUBLISHING.zh-CN.md) | 全球 MCP Registry、MCPB、npm 第二渠道、发布流程和验证步骤。 |

Word AI is published for global discovery through the official MCP Registry using a public GitHub Release MCPB artifact. npm packages are secondary convenience channels and should not replace Registry/MCPB metadata for MCP host discovery.

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
scripts/publish_dotnet.sh osx-arm64      # repeat for each release RID in CI
scripts/publish_dotnet.sh linux-x64
scripts/publish_dotnet.sh linux-arm64
powershell -ExecutionPolicy Bypass -File scripts\publish_dotnet.ps1 -RuntimeIdentifier win-x64
PYTHONPATH=. python scripts/build_mcpb.py --version 0.8.1 --out dist/word-ai-0.8.1.mcpb
shasum -a 256 dist/word-ai-0.8.1.mcpb
```

When `dist/native/<rid>/WordAi.OpenXml` exists before the MCPB build, the native .NET Open XML backend is bundled and selected before DLL/project/Python fallback. The MCPB bootstrap still requires Python 3.10+ for the MCP facade. On first run it creates `.word-ai-mcpb-venv` inside the installed bundle directory, installs Word AI dependencies there, and then starts the stdio MCP server. Bootstrap installation logs are written to stderr so stdout remains reserved for MCP JSON-RPC.

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

## Secondary npm Channel

npm is useful for quick no-clone startup and clients that cannot consume MCP Registry/MCPB packages yet. Keep Registry/MCPB first in user-facing installation docs, then present npm as a fallback:

- Recommended scoped package: `@flyfish-dev/word-ai`
- Unscoped compatibility package: `word-ai-mcp`
- CLI commands: `word-ai`, `word-ai-mcp`, and `word-ai-http`

Example smoke check:

```bash
npm exec --yes --package @flyfish-dev/word-ai -- word-ai --root "$PWD" doctor
npm exec --yes --package word-ai-mcp -- word-ai-mcp --root "$PWD"
```

## Release Flow

1. Update `pyproject.toml`, `word_ai_mcp/__init__.py`, `server.json`, README, and changelog to the release version.
2. Merge to `main`.
3. Create and push a version tag:

   ```bash
   git tag v0.8.1
   git push origin v0.8.1
   ```

4. GitHub Actions publishes native .NET backends for target RIDs, builds `word-ai-<version>.mcpb`, and computes its `fileSha256`.
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
