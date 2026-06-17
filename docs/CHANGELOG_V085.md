# Changelog v0.8.5

Release focus: make Word AI easier to install, easier to hand to another agent user, and safer to run through the .NET-first PatchSet path.

## Added

- Single-file `word-ai` standalone executable entrypoint for stdio MCP, HTTP bridge, readiness checks, Codex config generation, and Agent Skill installation.
- `scripts/build_standalone.py` for PyInstaller one-file packaging with bundled schemas, skills, docs, and the current-platform .NET Open XML backend.
- `scripts/package_quickstart.py` for quickstart archives containing the executable, README, manifest, and AGPL license.
- `requirements-standalone.txt` for optional standalone build dependencies.
- GitHub Actions workflow for standalone binaries and quickstart bundles across `linux-x64`, `linux-arm64`, `osx-arm64`, `osx-x64`, `win-x64`, and `win-arm64`.
- Distribution documentation in English and Chinese.

## Changed

- Quickstart and Codex config generation can now target a standalone executable directly, producing `word-ai mcp --root ...` configs instead of requiring a local source checkout.
- Runtime resource discovery works inside PyInstaller one-file extraction, so bundled skills, schemas, and native .NET backends are found from the executable.
- Agent Skill installation can copy from bundled standalone resources as well as from a source checkout.
- Release documentation now treats MCP Registry/MCPB as the primary MCP discovery channel, standalone quickstart as the lowest-friction local command path, and npm as the secondary fallback.
- Release smoke tests now verify standalone `--version`, `install-skill --dry-run`, MCP handshake, and `tools/list` with 63 tools.

## Verified Locally

- Default offline engine selected the `.NET native` backend.
- Agent-style PatchSet aliases such as `operation`, `target_tag`, `new_text`, `text_sha256`, and camelCase operation names passed assess, dry-run, apply, validation, and audit through the default .NET route.
- The `osx-arm64` standalone binary ran `--version`, `install-skill --dry-run`, `doctor`, `codex-config`, MCP initialize, and `tools/list`.
- The `osx-arm64` quickstart archive extracted and ran the same MCP smoke checks.
