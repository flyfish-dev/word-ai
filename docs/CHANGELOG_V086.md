# Changelog v0.8.6

Release focus: reduce customer-facing release clutter and make Open XML backend delivery invisible.

## Changed

- GitHub Releases now expose only the assets ordinary users should choose from: one MCPB plus one quickstart bundle per supported platform.
- Per-RID `word-ai-openxml-*` backend archives and bare `word-ai-standalone-*` binaries are no longer uploaded to the default release. The Open XML backend is bundled inside MCPB and linked into each quickstart executable.
- The npm launcher now downloads the current-platform quickstart bundle and executes its bundled standalone `word-ai`, so the default npm path no longer creates a Python venv or downloads a separate Open XML backend.
- The legacy npm source bootstrap path remains available with `WORD_AI_NPM_USE_SOURCE_BOOTSTRAP=1`.
- Documentation now tells users to pick MCP Registry/MCPB or their platform quickstart bundle, not individual backend files.

## 验证重点

- Open XML 后端仍内置在 MCPB 和 quickstart executable 中。
- 普通 GitHub Release 页面不再展示一大片后端散件。
- npm 默认路径直接复用 quickstart 包，降低依赖安装和运行时差异。
