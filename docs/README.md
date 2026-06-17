# Word AI Documentation

| Language | Preview |
| --- | --- |
| [English](README.md) | Documentation index for installation, architecture, contracts, security, QA, and release operations. |
| [中文](README.zh-CN.md) | 安装、架构、工具契约、安全、验证和发布相关文档导航。 |

This directory contains the design, tool contract, safety policy, and validation material for Word AI.

## English

- [Getting Started](GETTING_STARTED.md): install, build, run, and configure Codex with MCP Registry/Skill first and npm as a secondary channel.
- [Word AI Codex Skill](../skills/word-ai/SKILL.md): official Codex Skill rules for offline `docx_*`, live `word_session_*`, and optional read-only OfficeCLI evidence.
- [Architecture](ARCHITECTURE.en.md): system architecture and component responsibilities.
- [Tool Contract](TOOL_CONTRACT.md): MCP tool lifecycle and PatchSet rules.
- [Security Design](SECURITY.en.md): path safety, write safety, Office bridge token model, and deployment notes.
- [QA Report](QA_REPORT.md): local verification evidence.
- [Validation Matrix](VALIDATION_MATRIX.md): validation coverage and expectations.
- [MCP Registry Publishing](REGISTRY_PUBLISHING.md): MCPB asset and official MCP Registry release flow.
- [Stability Policy](STABILITY_POLICY.md): editing risk classes and safety boundaries.
- [Development Plan](DEVELOPMENT_PLAN.md): roadmap for production hardening.
- [Code Tool Catalog](CODEX_TOOL_CATALOG.md): generated MCP tool catalog.
- [v0.8.3 Changelog](CHANGELOG_V083.md): cross-platform native distribution and full release artifacts.
- [v0.8.1 Changelog](CHANGELOG_V081.md): MCPB Registry release path.
- [v0.8.0 Changelog](CHANGELOG_V080.md): strict license and global MCP distribution.
- [v0.7.1 Changelog](CHANGELOG_V071.md): multi-root path access for external document folders.
- [v0.7 Changelog](CHANGELOG_V07.md): formal Codex Skill, OfficeCLI boundaries, and one-command startup.
- [v0.6 Changelog](CHANGELOG_V06.md): live Word session MCP workflow.

## Recommended Reading Order

1. Start with [Getting Started](GETTING_STARTED.md).
2. Read [Tool Contract](TOOL_CONTRACT.md) before building an agent workflow.
3. Read [Security Design](SECURITY.en.md) before enabling write tools.
4. Read [Architecture](ARCHITECTURE.en.md) before changing internals.
