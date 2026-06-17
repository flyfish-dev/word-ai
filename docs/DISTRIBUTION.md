# Distribution

| Language | Preview |
| --- | --- |
| [English](DISTRIBUTION.md) | Release channels, standalone binaries, quickstart bundles, Agent Skill install, and npm fallback. |
| [中文](DISTRIBUTION.zh-CN.md) | 发布渠道、单文件二进制、快速集成包、Agent Skill 安装和 npm 备用渠道。 |

Word AI has two primary distribution paths and one secondary fallback:

1. **MCP Registry / MCPB** for MCP host discovery and standardized metadata.
2. **Standalone quickstart bundles** for users who need a single local command with no source checkout, Python venv, pip install, or .NET SDK.
3. **npm** as a secondary convenience channel for hosts that cannot consume MCP Registry/MCPB yet.

## Recommended Paths

For MCP-capable clients, install `io.github.flyfish-dev/word-ai` from the MCP Registry first. This keeps server identity, transport, version, and package checksum in one standardized place.

For Codex, Claude Code, and compatible agent clients, install the `word-ai` Agent Skill as well. The Skill tells the agent when to choose offline `docx_*` tools, when to choose live `word_session_*` tools, and why OfficeCLI is only read-only evidence in the default workflow.

For users who just need a local command quickly, download the quickstart bundle for the current platform from the GitHub Release:

```bash
tar -xzf word-ai-quickstart-0.8.5-osx-arm64.tar.gz
cd word-ai-quickstart-0.8.5-osx-arm64

./word-ai install-skill
./word-ai codex-config --output .wordai/codex-config.toml
./word-ai mcp --root "$PWD" --allow-root "$HOME/Downloads" --allow-root "$HOME/Documents"
```

On Windows, use the `.zip` bundle and `word-ai.exe`:

```powershell
.\word-ai.exe install-skill
.\word-ai.exe codex-config --output .wordai\codex-config.toml
.\word-ai.exe mcp --root "$PWD" --allow-root "$HOME\Documents"
```

## Standalone Binary Contract

The standalone executable bundles:

- Python MCP facade and Python runtime dependencies.
- The current-platform .NET Open XML backend.
- PatchSet schemas.
- The official `word-ai` Agent Skill template.
- The quickstart, doctor, Codex config, stdio MCP, and HTTP bridge entrypoints.

It supports:

```bash
word-ai --version
word-ai doctor
word-ai install-skill
word-ai codex-config --output .wordai/codex-config.toml
word-ai mcp --root "$PWD"
word-ai http --root "$PWD" --host 127.0.0.1 --port 8765
```

Offline DOCX editing does not require a local Python install, pip install, .NET SDK, or source checkout. Full Office.js taskpane development still uses the source install path because the Word add-in project must be built and sideloaded locally.

## Release Assets

Each release publishes:

- `word-ai-<version>.mcpb` for MCP Registry/MCPB installation.
- Native .NET Open XML backend archives for `osx-arm64`, `osx-x64`, `linux-x64`, `linux-arm64`, `linux-musl-x64`, `linux-musl-arm64`, `win-x64`, and `win-arm64`.
- Standalone single-file binaries for standard hosted platforms: `linux-x64`, `linux-arm64`, `osx-arm64`, `osx-x64`, `win-x64`, and `win-arm64`.
- Quickstart bundles wrapping each standalone binary with a README, manifest, and license.
- SHA-256 checksum files for release verification.

The MCPB path includes all native Open XML backends. The standalone path packages one platform per artifact. The npm path keeps the package lightweight and downloads only the current-platform native backend on first run.

## Build Locally

Build the current-platform standalone binary:

```bash
python -m pip install -r requirements.txt -r requirements-standalone.txt
PYTHONPATH=. python scripts/publish_native_backends.py "$(PYTHONPATH=. python - <<'PY'
from scripts.publish_native_backends import host_rid
print(host_rid())
PY
)"
PYTHONPATH=. python scripts/build_standalone.py --json
```

Package the quickstart bundle:

```bash
PYTHONPATH=. python scripts/package_quickstart.py --version 0.8.5 --rid osx-arm64 --json
```

Smoke test:

```bash
dist/standalone/osx-arm64/word-ai --version
dist/standalone/osx-arm64/word-ai install-skill --dry-run
dist/standalone/osx-arm64/word-ai codex-config --output /tmp/word-ai-codex.toml
```

## Publishing

Publishing is tag-driven:

```bash
git tag v0.8.5
git push origin v0.8.5
```

The release workflows build and smoke test MCPB, native backends, standalone binaries, quickstart bundles, Docker, npm payloads, and MCP Registry metadata. npm publication requires trusted publishing or a valid `NPM_TOKEN`; if npm authorization is not configured, the workflow warns and continues so MCP Registry/GitHub Release remain available.
