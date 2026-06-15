# AGENTS.md — Word AI MCP 使用规则

处理 Word DOCX 交付文档时，必须把结构稳定放在内容生成之前。本工程提供 63 个 MCP tools；读/分析工具可以充分调用，正式写入必须通过 PatchSet。

## 不可违反的原则

- 不重建整篇 DOCX。
- 不把 DOCX 转 Markdown/HTML 后再转回 DOCX。
- 不直接覆盖源文件；默认输出新文件。
- 优先使用内容控件 tag 作为写入锚点。
- 写入前必须 assess + dry-run；写入后必须 validate + diff。
- 不修改样式、编号、页眉页脚、目录、字段、关系文件、图片关系，除非用户明确授权并进入高风险审批模式。
- 离线文件编辑优先使用 `docx_*`；用户明确要求编辑当前 Word 窗口或已连接 Office.js taskpane 时，才使用 `word_session_*`。
- OfficeCLI 只能作为只读/低风险证据辅助，不作为默认写入路径。

## 标准 Codex 流程

1. `docx_health_check`：确认复杂对象、重复 tag、字段、批注、修订痕迹风险。
2. `docx_map` 或 `docx_list_anchors`：定位目标锚点。
3. `docx_read_content_control` / `docx_read_anchor` / `docx_read_table`：读取目标文本，记录 `text_sha256`。
4. 生成 PatchSet。优先操作：
   - `replace_content_control_text`
   - `append_content_control_text`
   - `prepend_content_control_text`
   - `replace_text_in_content_control`
5. `docx_assess_patchset`：解析目标、复杂对象和风险。
6. `docx_dry_run_patchset`：临时写入并验证候选文件。
7. `docx_backup`：创建备份。
8. `docx_apply_patchset`：生成新 DOCX 和 audit JSON。
9. `docx_validate` / `docx_compare_structure` / `docx_text_diff`：给出验证和差异。

## PatchSet 规则

推荐形式：

```json
{
  "schema_version": "2.0",
  "strict": true,
  "source_sha256": "可选：docx_inspect 返回的源文件 sha256",
  "reason": "说明本次修改原因",
  "guard": {
    "require_preconditions": true,
    "allow_overwrite": false
  },
  "operations": [
    {
      "op": "replace_content_control_text",
      "tag": "WORD-AI:SRS:1.0:overview",
      "expected_old_sha256": "目标锚点原文 hash",
      "text": "新文本",
      "preserve_style": true,
      "allow_complex_content": false
    }
  ]
}
```

## 高风险操作

以下操作只有在没有内容控件锚点、且已通过 `docx_assess_patchset` / `docx_dry_run_patchset` 时使用：

- `replace_paragraph_text`
- `insert_paragraph_before`
- `insert_paragraph_after`
- `replace_table_cell_text`
- `append_table_row`
- `add_comment`

高风险操作必须提供 `expected_old_text` 或 `expected_old_sha256`。表格操作前必须先调用 `docx_list_tables` 和 `docx_read_table`。

## 完成后向用户说明

- 修改了哪些锚点或对象。
- 是否只改变允许的 DOCX part。
- 表格、图片、字段、批注、修订痕迹、内容控件 tag 是否保持稳定。
- 新文件路径、audit JSON 路径和 diff 摘要。
