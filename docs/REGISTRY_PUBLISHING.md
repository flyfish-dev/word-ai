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
scripts/publish_dotnet.sh --all
PYTHONPATH=. python scripts/verify_native_backends.py --require-all
PYTHONPATH=. python scripts/package_native_assets.py --version 0.8.5
PYTHONPATH=. python scripts/build_mcpb.py --version 0.8.5 --out dist/word-ai-0.8.5.mcpb
shasum -a 256 dist/word-ai-0.8.5.mcpb
```

The release RID set is `osx-arm64`, `osx-x64`, `linux-x64`, `linux-arm64`, `linux-musl-x64`, `linux-musl-arm64`, `win-x64`, and `win-arm64`. When `dist/native/<rid>/WordAi.OpenXml` or `WordAi.OpenXml.exe` exists before the MCPB build, the native .NET Open XML backend is bundled and selected before DLL/project/Python fallback. The npm package intentionally does not embed all native backends; its launcher downloads the current-platform GitHub Release archive, verifies SHA-256, caches it under the user cache, and exposes it through `WORD_AI_DOTNET_NATIVE_DIR`. The MCPB bootstrap still requires Python 3.10+ for the MCP facade. On first run it creates `.word-ai-mcpb-venv` inside the installed bundle directory, installs Word AI dependencies there, and then starts the stdio MCP server. Bootstrap installation logs are written to stderr so stdout remains reserved for MCP JSON-RPC.

GitHub Releases also include per-RID native archives named `word-ai-openxml-<version>-<rid>.tar.gz` or `.zip`, standalone binaries named `word-ai-standalone-<version>-<rid>`, quickstart bundles named `word-ai-quickstart-<version>-<rid>`, and checksum files.

Build the current-platform standalone and quickstart package:

```bash
python -m pip install -r requirements-standalone.txt
PYTHONPATH=. python scripts/build_standalone.py --rid osx-arm64 --json
PYTHONPATH=. python scripts/package_quickstart.py --version 0.8.5 --rid osx-arm64 --json
dist/standalone/osx-arm64/word-ai --version
dist/standalone/osx-arm64/word-ai install-skill --dry-run
```

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
- Native backend delivery: npm downloads only the current RID archive from the GitHub Release, verifies the checksum, and caches it. Use `WORD_AI_SKIP_NATIVE_DOWNLOAD=1` when an environment must stay offline after install.

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
   git tag v0.8.5
   git push origin v0.8.5
   ```

4. GitHub Actions publishes native .NET backends for all release RIDs and verifies the current-platform binary.
5. The workflow builds `word-ai-<version>.mcpb`, computes its `fileSha256`, and verifies the npm payload is lightweight and does not embed oversized native binaries.
6. The workflow uploads the MCPB file, per-RID native archives, and checksums to the GitHub Release.
7. The standalone workflow builds and smokes single-file binaries plus quickstart bundles for `linux-x64`, `linux-arm64`, `osx-arm64`, `osx-x64`, `win-x64`, and `win-arm64`, then uploads them to the same GitHub Release.
8. The workflow publishes `@flyfish-dev/word-ai` and `word-ai-mcp` to npm when that version is not already present.
9. The workflow authenticates to the MCP Registry with GitHub OIDC.
10. `mcp-publisher validate` and `mcp-publisher publish` publish the version metadata from `server.json`.

## Verification

After the release workflow finishes:

```bash
curl "https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.flyfish-dev/word-ai"
curl "https://registry.modelcontextprotocol.io/v0.1/servers/io.github.flyfish-dev%2Fword-ai/versions/latest"
```

If the registry publish step fails with a namespace permission error, verify that the GitHub identity used by the workflow can publish for the `flyfish-dev` organization namespace, or switch the registry authentication method to DNS verification for a custom domain.
