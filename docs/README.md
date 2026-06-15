# Word AI Documentation

This directory contains the design, tool contract, safety policy, and validation material for Word AI.

中文文档导航见下方：[中文](#中文文档导航)。

## English

- [Getting Started](GETTING_STARTED.md): install, build, run, and configure Codex.
- [Architecture](ARCHITECTURE.en.md): system architecture and component responsibilities.
- [Tool Contract](TOOL_CONTRACT.md): MCP tool lifecycle and PatchSet rules.
- [Security Design](SECURITY.md): path safety, write safety, Office bridge token model, and deployment notes.
- [QA Report](QA_REPORT.md): local verification evidence.
- [Validation Matrix](VALIDATION_MATRIX.md): validation coverage and expectations.
- [Stability Policy](STABILITY_POLICY.md): editing risk classes and safety boundaries.
- [Development Plan](DEVELOPMENT_PLAN.md): roadmap for production hardening.
- [Code Tool Catalog](CODEX_TOOL_CATALOG.md): generated MCP tool catalog.

## 中文文档导航

- [快速上手](GETTING_STARTED.md)：安装、构建、运行和 Codex 配置。
- [架构说明](ARCHITECTURE.md)：系统架构和组件职责。
- [英文架构说明](ARCHITECTURE.en.md)：面向国际协作的英文架构文档。
- [工具契约](TOOL_CONTRACT.md)：MCP 工具链路与 PatchSet 规则。
- [安全设计](SECURITY.md)：路径安全、写入安全、Office bridge token 和部署建议。
- [QA 报告](QA_REPORT.md)：本地验证证据。
- [验证矩阵](VALIDATION_MATRIX.md)：结构验证覆盖范围。
- [稳定性策略](STABILITY_POLICY.md)：编辑风险分级与安全边界。
- [开发计划](DEVELOPMENT_PLAN.md)：生产化路线图。
- [工具清单](CODEX_TOOL_CATALOG.md)：MCP tools 目录。

## Recommended Reading Order

1. Start with [Getting Started](GETTING_STARTED.md).
2. Read [Tool Contract](TOOL_CONTRACT.md) before building an agent workflow.
3. Read [Security Design](SECURITY.md) before enabling write tools.
4. Read [Architecture](ARCHITECTURE.en.md) or [架构说明](ARCHITECTURE.md) before changing internals.

