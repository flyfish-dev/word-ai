# Word AI 文档导航

| 文档 | 预览 |
| --- | --- |
| [中文](README.zh-CN.md) | 安装、架构、工具契约、安全、验证和发布相关文档导航。 |
| [English](README.md) | Documentation index for installation, architecture, contracts, security, QA, and release operations. |

本目录包含 Word AI 的设计、工具契约、安全策略和验证材料。

## 中文

- [快速上手](GETTING_STARTED.zh-CN.md)：安装、构建、运行和 Codex 配置；优先 MCP Registry/Skill，npm 作为第二渠道。
- [分发说明](DISTRIBUTION.zh-CN.md)：MCP Registry、standalone 二进制、quickstart 包、Agent Skill 安装和 npm 备用渠道。
- [Word AI Codex Skill](../skills/word-ai/SKILL.md)：正式 Codex Skill，定义离线 `docx_*`、实时 `word_session_*` 与 OfficeCLI 只读辅助规则。
- [架构说明](ARCHITECTURE.md)：系统架构和组件职责。
- [英文架构说明](ARCHITECTURE.en.md)：面向国际协作的英文架构文档。
- [工具契约](TOOL_CONTRACT.md)：MCP 工具链路与 PatchSet 规则。
- [安全设计](SECURITY.md)：路径安全、写入安全、Office bridge token 和部署建议。
- [QA 报告](QA_REPORT.md)：本地验证证据。
- [验证矩阵](VALIDATION_MATRIX.md)：结构验证覆盖范围。
- [MCP Registry 发布说明](REGISTRY_PUBLISHING.zh-CN.md)：MCPB 资产和官方 MCP Registry 发布流程。
- [稳定性策略](STABILITY_POLICY.md)：编辑风险分级与安全边界。
- [开发计划](DEVELOPMENT_PLAN.md)：生产化路线图。
- [工具清单](CODEX_TOOL_CATALOG.md)：MCP tools 目录。
- [v0.8.6 变更记录](CHANGELOG_V086.md)：简化客户发布资产，并让 npm 直接使用 quickstart launcher。
- [v0.8.5 变更记录](CHANGELOG_V085.md)：standalone quickstart 包和更完整的发布烟测。
- [v0.8.4 变更记录](CHANGELOG_V084.md)：Agent 友好的 PatchSet 别名归一化。
- [v0.8.3 变更记录](CHANGELOG_V083.md)：跨平台 native 分发与完整发布产物。
- [v0.8.1 变更记录](CHANGELOG_V081.md)：MCPB Registry 发布路径。
- [v0.8.0 变更记录](CHANGELOG_V080.md)：严格许可证和全球 MCP 分发。
- [v0.7.1 变更记录](CHANGELOG_V071.md)：面向外部文档目录的多 root 路径访问。
- [v0.7 变更记录](CHANGELOG_V07.md)：正式 Codex Skill、OfficeCLI 边界和一键启动。
- [v0.6 变更记录](CHANGELOG_V06.md)：Word 会话 MCP 闭环。

## 推荐阅读顺序

1. 先读 [快速上手](GETTING_STARTED.zh-CN.md)。
2. 选择安装或发布渠道时阅读 [分发说明](DISTRIBUTION.zh-CN.md)。
3. 构建 Agent 工作流前阅读 [工具契约](TOOL_CONTRACT.md)。
4. 启用写入工具前阅读 [安全设计](SECURITY.md)。
5. 修改内部实现前阅读 [架构说明](ARCHITECTURE.md)。
