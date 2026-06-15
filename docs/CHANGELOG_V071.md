# v0.7.1 Changelog - Multi-Root Path Access

## English

v0.7.1 improves real-world usability for Codex and local agents that need to edit source DOCX files outside the repository.

### Added

- Multi-root path policy for the MCP server and HTTP bridge:
  - `--root` remains the primary workspace and relative-path base.
  - `--allow-root <dir>` can be repeated for external document folders.
  - `WORD_AI_ALLOWED_ROOTS` can also provide extra roots using the platform path separator.
- Installer-generated Codex config now includes common user document folders when they exist:
  - `~/Downloads`
  - `~/Documents`
  - `~/Desktop`
- Startup scripts pass the same common allowed roots to the Office bridge.
- `doctor` prints the default allowed roots.
- Smoke coverage verifies that paths outside root are rejected unless explicitly allowlisted.

### Security Model

This does not allow arbitrary disk access. Every absolute path must still resolve inside the primary root or an explicit allowed root.

## 中文

v0.7.1 提升了 Codex 和本地智能体处理仓库外原始 DOCX 文件的可用性。

### 新增

- MCP server 和 HTTP bridge 支持多根目录路径策略：
  - `--root` 仍是主工作区和相对路径基准。
  - `--allow-root <dir>` 可重复传入，用于外部文档目录。
  - `WORD_AI_ALLOWED_ROOTS` 也可用平台路径分隔符提供额外目录。
- 安装脚本生成的 Codex 配置会在存在时自动加入常用用户文档目录：
  - `~/Downloads`
  - `~/Documents`
  - `~/Desktop`
- 启动脚本会给 Office bridge 传入同样的常用 allowed roots。
- `doctor` 会打印默认 allowed roots。
- smoke test 覆盖：未授权越界路径仍拒绝，显式 allowlist 后允许。

### 安全模型

这不是任意磁盘访问。所有绝对路径仍必须解析到主 root 或显式允许的 root 内。
