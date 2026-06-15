# 开发计划

## Phase 0：基线与样本文档库

交付：

- 需求规格说明书、概要设计、详细设计、数据库设计四类模板。
- 每类模板至少 10 个真实复杂样本。
- 样本覆盖：目录、自动编号、页眉页脚、分页符、分节符、批注、修订、交叉引用、图片、图题、表题、脚注、复杂表格。

验收：

- 所有样本能被 Word 正常打开。
- 索引器能输出标题树和内容控件清单。
- 渲染输出页数稳定。

## Phase 1：内容控件锚点体系

交付：

- Office.js 加载项。
- 本地 Office bridge：读取锚点、生成 PatchSet、assess、dry-run、apply、validate、diff。
- taskpane 操作台：包装选区、列出打开文档内容控件、文件级闭环应用、打开文档 hash 校验写入。
- Tag 命名规范。
- 模板预埋规范。
- 锚点巡检工具。

验收：

- 每个可 AI 编辑区域有唯一 tag。
- tag 不重复、不为空、不跨越危险结构。
- 表格、图片、字段附近的锚点有风险标记。
- Office bridge smoke 能完整跑通 read/build/assess/dry-run/apply/validate/diff。

## Phase 2：MCP 读工具

交付：

- `docx_inspect`
- `docx_list_anchors`
- `docx_read_content_control`
- `docx_write_index`

验收：

- Agent 能只读取目标 chunk 完成改写。
- 对 300 页文档，索引和读取单个锚点不需要加载给模型全篇正文。

## Phase 3：PatchSet 与写工具

交付：

- PatchSet JSON Schema。
- `docx_backup`
- `docx_apply_patchset`
- 审计 JSON。

验收：

- 默认输出新 DOCX。
- 只允许白名单操作。
- 优先支持内容控件文本替换、表格单元格文本替换、段落文本替换。

## Phase 4：结构验证

交付：

- ZIP part hash 对比。
- 内容控件集合对比。
- 表格、图片、字段、批注、修订计数对比。
- styles/numbering/settings/rels 不变性检查。

验收：

- 任何非授权 part 变化都阻断。
- 表格、图片、字段数量变化默认报错。
- 验证报告可被 CI 解析。

## Phase 5：Open XML SDK 生产内核

交付：

- .NET class library。
- OpenXmlReader streaming index。
- DOM/typed API patch applier。
- OpenXmlValidator 集成。
- 与 MCP facade 的进程内或 gRPC 集成。

验收：

- 100MB 级 DOCX 能索引。
- 大文档 patch 不超过可接受延迟。
- 内存峰值可控。

## Phase 6：Word 渲染与字段刷新

交付：

- Word COM 或 Aspose/Syncfusion post-processor。
- 更新 TOC、PAGEREF、REF、SEQ 字段。
- 导出 PDF。
- 渲染 diff。

验收：

- 目录页码和交叉引用刷新。
- PDF 与 DOCX 视觉快照一致。
- 表格分页、图片锚点无异常。

## Phase 7：企业化

交付：

- 权限边界：工作区 root、只读模式、写操作审批。
- 审计：请求、patch、输出 hash、操作者、模型、时间。
- 安全：工具 allowlist、路径防逃逸、敏感文档脱敏策略。
- 可观测：指标、耗时、失败原因分类。

验收：

- 所有生产写操作有可追溯记录。
- 支持回滚。
- 支持批量任务队列。
