# Word AI 编辑系统架构

## 1. 总体目标

目标不是做一个“会写 Word 的聊天机器人”，而是做一个**以 Word 原结构为边界的文档事务系统**：

- 模型负责理解需求、生成修改计划、生成目标文本。
- MCP Server 负责暴露安全工具契约。
- OOXML/Open XML SDK 内核负责定点编辑。
- Office.js 负责 Word 端锚点治理、人工审批和打开文档的轻量会话写入。
- 验证器负责阻断任何结构破坏。

## 2. 组件分工

### 2.1 Codex / Agent

职责：

- 读取文档索引和目标 chunk。
- 输出结构化 `PatchSet`。
- 不直接操作 DOCX 二进制或 ZIP 包。
- 不生成完整文档，只生成目标锚点的新文本或受限表格单元格内容。

### 2.2 MCP Server

职责：

- 暴露 63 个分层工具：丰富读取与定位能力，正式写入统一收口到 PatchSet。
- 限制文件根目录。
- 校验 JSON Schema。
- 默认不覆盖原文档。
- 每次写操作生成审计记录。
- 维护 `.wordai/sessions` 文件队列，把 Codex 的 `word_session_*` 命令转交给 Office.js taskpane。

建议工具分层：

- Read tools：`docx_health_check`、`docx_map`、`docx_list_anchors`、`docx_read_content_control`、`docx_read_table_cell`、`docx_list_fields`、`docx_list_images` 等。
- Plan tools：`docx_write_index`、`docx_assess_patchset`、`docx_plan_patchset`。
- Dry-run tools：`docx_preflight_patchset`、`docx_dry_run_patchset`。
- Write tools：`docx_backup`、`docx_apply_patchset`、`docx_restore_backup`、`docx_rollback`。
- Verify tools：`docx_validate`、`docx_compare_structure`、`docx_text_diff`。
- Word session tools：`word_session_list`、`word_session_read_content_control`、`word_session_preview_patchset`、`word_session_apply_patchset`、`word_session_rollback`。
- OfficeCLI auxiliary tools：`officecli_view_html`、`officecli_view_screenshot`、`officecli_view_issues`、`officecli_query`、`officecli_validate`，仅作为只读/低风险证据，不作为默认写入后端。

### 2.3 Open XML SDK 内核

职责：

- 解析 WordprocessingML。
- 用 `SdtElement`、`Paragraph`、`TableCell` 等对象定点回写。
- 使用 `OpenXmlReader` 做超大文档 streaming index。
- 保持 `styles.xml`、`numbering.xml`、`rels`、`media`、`settings.xml` 原样不动。

### 2.4 Office.js 加载项

职责：

- 在 Word 中把选区包装为内容控件。
- 给内容控件写入稳定 tag，例如：`WORD-AI:SRS:2.1:functional_scope`。
- 给锚点设置 title/alias，便于人工查看。
- 允许文档编写人员在交付模板中预埋可编辑区域。
- 连接本地 Office bridge，读取文件级内容控件锚点。
- 为内容控件文本修改生成带 `source_sha256` 和 `expected_old_sha256` 的 PatchSet。
- 在 Word taskpane 内执行 assess、dry-run、apply，并展示 output DOCX、audit JSON、validation 和 diff。
- 注册当前打开的 Word 文档为 live session，持续 heartbeat 并轮询本地命令队列。
- 对当前打开的 Word 文档执行内容控件级 `word_session_apply_patchset`，写入前先 live preflight，再用客户端 SHA-256 校验打开文档文本没有漂移。
- 返回 Office.js 写入审计、触达内容控件、校验结果和 rollback PatchSet。

Office.js 不应承担批量后台处理主职责，因为超大文档、批处理、字段刷新和复杂错误恢复更适合后端内核或 Word 桌面自动化。当前最佳分工是：Office.js 做会话交互和人工审批，本地 MCP/.NET/Python 内核做权威文件事务。

### 2.5 Office Bridge

本地 bridge 由 `word_ai_mcp.server_http` 提供：

- `/office/read`：聚合 inspect、health、content controls、tables、fields、comments、revisions。
- `/office/build-patchset`：按内容控件 tag 读取当前文本 hash，生成安全 PatchSet。
- `/office/assess-patchset`：执行风险评估。
- `/office/preview-patchset`：执行 dry-run 和结构验证，默认不保留临时输出。
- `/office/apply-patchset`：执行 health -> assess -> dry-run -> backup -> apply -> validate -> text diff。
- `/office/session/register`：taskpane 注册打开文档、能力和内容控件快照。
- `/office/session/heartbeat`：刷新打开文档快照和 session 状态。
- `/office/session/poll`：taskpane 拉取 Codex 排队的 session 命令。
- `/office/session/result`：taskpane 回传 Office.js 执行结果、错误、audit 和 rollback PatchSet。
- `/office/session/list`：列出当前本地 session。

打开文档编辑链路：

```text
Codex MCP client
  -> word_session_* tool
  -> .wordai/sessions/commands/*.json
  -> Office.js taskpane poll
  -> Word.run(...)
  -> /office/session/result
  -> Codex receives audit / rollback / error
```

安全边界：

- 所有 `/office/*` POST 必须携带 `X-Word-AI-Token`。
- JSON-RPC `/mcp` 写工具也要求 token；读工具保留本地开发可用性。
- CORS 仅允许 localhost/127.0.0.1 开发源，并支持同源 `/bridge/*` 代理。
- 路径仍由 MCP root 限制，所有 DOCX 路径必须在 root 内。

## 3. 内容控件锚点策略

推荐 tag 规范：

```text
WORD-AI:{doc_type}:{chapter_path}:{semantic_key}:{version?}
```

示例：

```text
WORD-AI:SRS:1.0:overview
WORD-AI:HLD:3.2:deployment_architecture
WORD-AI:DBD:4.1:user_table_fields
```

内容控件的粒度建议：

- 一级/二级章节不要整体包太大。
- 每个可变正文块、规则列表、接口说明表、数据库字段表分别设置锚点。
- 对复杂表格，优先锚定表格外部说明和单元格范围，不轻易重建整表。

## 4. 文档索引模型

索引对象：

- DOCX 文件哈希、大小、part 列表。
- 标题树。
- 内容控件列表。
- 书签列表。
- 段落索引、段落文本摘要、styleId、paraId。
- 表格索引、行列维度、首行表头。
- 图片关系、drawing anchor 信息。
- 字段和交叉引用摘要。
- 批注和修订痕迹摘要。

索引落地：

- 小文档：JSON sidecar。
- 大文档：SQLite + FTS5 + chunk hash。
- 生产系统：对象存储保存原始 DOCX、索引、patchset、审计报告和渲染快照。

## 5. 回写事务

```text
prepare
  -> backup source.docx
  -> parse patchset
  -> validate anchors exist
  -> apply changes to temp.docx
  -> validate unchanged invariants
  -> text diff / optional render diff
  -> commit output.docx or abort
```

所有写操作必须满足：

- 输出新文件，不覆盖原文件。
- 审计记录包含 source hash、target hash、patchset、changed parts、验证结果。
- 如果结构验证失败，目标文件仍保留但标记 `ok=false`，不得进入交付目录。

打开文档 live session 的写入事务不生成新 DOCX 文件；它作用于 Word 当前会话中的内容控件文本。该路径必须满足：

- 写入前执行 `word_session_preview_patchset` 或 apply 内置 live preflight。
- 每个 operation 带 `expected_old_sha256`，且必须与 Word 打开文档当前文本一致。
- 只允许内容控件文本类操作，不改样式、编号、关系、页眉页脚、图片或字段。
- 返回 audit 和 rollback PatchSet；回滚通过 `word_session_rollback` 再次走 Office.js session apply。

## 6. 性能设计

数百页、上万段落时，不应让 Agent 读全量正文：

- 第一次：索引全量文档，建立 sidecar。
- 交互时：Agent 只读取目标锚点和邻近标题上下文。
- 回写时：只打开并修改目标 XML part。
- 验证时：先做 ZIP part hash，再做关键 OOXML 结构计数，再按需做渲染 diff。

生产版 .NET 内核建议：

- `OpenXmlReader` streaming 扫描段落和 SDT。
- 对 `word/document.xml` 做局部 XPath/DOM 定位；对超大 XML 使用可定位 token 或 OpenXmlPowerTools 思路分段。
- PatchSet 合并执行，避免每个 operation 反复打开/保存 DOCX。

## 7. 最终推荐部署形态

### MVP

- Python stdio MCP + OOXML 定点修改。
- 内容控件锚点。
- JSON sidecar 索引。
- 结构验证和审计。

### Production

- MCP facade：TypeScript 或 Python。
- 核心引擎：.NET Open XML SDK。
- 索引库：SQLite/FTS 或 PostgreSQL。
- Word 桌面补强：Office.js 创建锚点、审批 PatchSet、当前会话内容控件写入；Word COM/Aspose/Syncfusion 刷新字段和 PDF。
- CI：样本文档库 + OpenXML validation + render diff。
