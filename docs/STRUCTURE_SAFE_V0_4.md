# v0.4 结构稳定优先实现说明

本版本的目标不是“让模型自由编辑 Word”，而是把模型限制在可验证、可回滚、可审计的 OOXML 定点写入通道内。默认策略是：**读全局结构，写局部锚点；验证全局不变量；失败不提交最终 DOCX。**

## 一、硬性结构保护原则

1. **禁止整篇重建**：不允许 DOCX → Markdown/HTML/TXT → DOCX 的往返重建链路。
2. **禁止默认覆盖源文件**：`docx_apply_patchset` 默认生成新 DOCX；同名覆盖必须显式设置 `guard.allow_overwrite=true`，且应先执行 `docx_backup`。
3. **写入作用域白名单**：常规写入只允许修改 `word/document.xml`，添加批注是例外，会受控修改 `word/comments.xml`、`word/_rels/document.xml.rels` 和 `[Content_Types].xml`。
4. **锚点优先级**：内容控件 `w:tag` > 书签/标题 > `w14:paraId` > 段落序号 > 搜索结果。生产模板必须优先布置内容控件。
5. **模型只输出 PatchSet**：Codex/Agent 不应直接输出 OOXML，也不应直接生成最终 DOCX。

## 二、PatchSet 门禁

建议所有正式写入都启用：

```json
{
  "schema_version": "2.0",
  "strict": true,
  "source_sha256": "<docx_inspect.sha256>",
  "guard": {
    "require_preconditions": true,
    "allow_overwrite": false
  },
  "operations": []
}
```

关键门禁：

- `source_sha256`：源 DOCX 发生任何字节级变化时拒绝写入，防止 Codex 基于过期上下文落盘。
- `expected_old_sha256` / `expected_old_text`：每个破坏性写操作都应绑定目标旧文本；启用 `guard.require_preconditions=true` 后，缺失前置条件会成为错误。
- `strict=true`：未授权 part 增删、未授权对象计数变化、未授权对象 hash 变化均视为错误。
- `abort_on_validation_error=true`：验证失败时只保留 invalid audit，不提交最终 DOCX。

## 三、推荐 Codex 调用链

```text
1. docx_health_check
2. docx_map / docx_list_anchors / docx_list_content_controls / docx_search_text
3. docx_read_content_control / docx_read_paragraph / docx_read_table_cell
4. 生成 PatchSet，包含 source_sha256 与 expected_old_sha256
5. docx_plan_patchset 或 docx_assess_patchset
6. docx_preflight_patchset 或 docx_dry_run_patchset
7. docx_backup
8. docx_apply_patchset
9. docx_validate 或 docx_compare_structure
10. docx_text_diff
```

## 四、验证不变量

`docx_validate` 会检查：

- ZIP 包完整性。
- DOCX part 增删情况。
- 仅允许授权 part 发生 hash 变化。
- 表格数、图片数、字段数、批注数、批注引用数、修订痕迹数、内容控件数、标题数、书签数。
- 未触达内容控件的 XML hash。
- 未触达表格的 XML hash。
- 未触达 `paraId` 段落的 XML hash。
- 内容控件 tag 集合是否稳定。
- 表格行列是否被意外改变。

`docx_apply_patchset` 会把实际触达对象写入 audit JSON；独立调用 `docx_validate` / `docx_compare_structure` 时应显式传入 `touched_*` 参数，或由调用方从 audit JSON 中提取触达对象后传入，避免把授权改动误判为异常。

## 五、工具覆盖面

本版本提供 49 个可发现工具，覆盖以下场景：

- 包结构与 hash：`docx_package_manifest`、`docx_list_parts`、`docx_part_hashes`、`docx_structural_fingerprint`。
- 文档导航：`docx_map`、`docx_get_outline`、`docx_list_headings`、`docx_list_anchors`、`docx_search_text`。
- 目标读取：`docx_read_content_control`、`docx_read_anchor`、`docx_read_paragraph`、`docx_read_heading_section`、`docx_read_table`、`docx_read_table_cell`。
- 复杂对象检查：`docx_list_styles`、`docx_list_numbering`、`docx_list_sections`、`docx_list_fields`、`docx_list_images`、`docx_list_hyperlinks`、`docx_list_headers_footers`、`docx_list_notes`、`docx_list_comments`、`docx_list_revisions`。
- 写入生命周期：`docx_assess_patchset`、`docx_plan_patchset`、`docx_preflight_patchset`、`docx_dry_run_patchset`、`docx_backup`、`docx_apply_patchset`。
- 验证与恢复：`docx_validate`、`docx_compare_structure`、`docx_text_diff`、`docx_restore_backup`、`docx_rollback`。
- sidecar 导出：`docx_write_index`、`docx_export_plain_text`、`docx_export_table_csv`、`docx_table_to_csv`。

## 六、写操作白名单

| 操作 | 风险 | 默认建议 |
|---|---:|---|
| `replace_content_control_text` | 低 | 首选。替换内容控件内文本，保留锚点外结构。 |
| `replace_text_in_content_control` | 低 | 只改内容控件内短语。 |
| `append_content_control_text` / `prepend_content_control_text` | 中 | 会改变段落数，需验证。 |
| `replace_table_cell_text` | 中 | 只改指定单元格文本。 |
| `replace_paragraph_text` | 中 | 没有内容控件时兜底，必须带 hash 前置条件。 |
| `insert_paragraph_after` / `insert_paragraph_before` | 中高 | 会改变段落数，需人工确认插入位置。 |
| `append_table_row` | 高 | 会改变表格结构，必须审批。 |
| `wrap_paragraph_with_content_control` | 中高 | 锚点治理操作，会新增内容控件。 |
| `add_comment` | 中高 | 会新增批注相关 part/关系。 |

## 七、验收标准

一次正式 Word 编辑交付必须同时满足：

- `docx_apply_patchset` 返回的 validation 为 `ok=true`。
- `docx_validate` 或 `docx_compare_structure` 为 `ok=true`。
- 只有预期 DOCX part 发生变化。
- 未触达内容控件、表格、段落 hash 未变化。
- 图片、字段、页眉页脚、修订痕迹等复杂对象无未授权变化。
- audit JSON 可追溯每个操作、目标、旧 hash、新 hash、changed parts 和 diff。
