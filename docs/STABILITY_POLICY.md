# 结构稳定策略

## 目标

让 Codex 能最大化使用 Word 编辑能力，同时把写入风险压到最低：任何添加内容或改动都应只影响被授权的最小 OOXML 范围，其他内容、样式、编号、表格、图片、字段、页眉页脚、批注、修订痕迹和关系文件保持不变。

## 稳定性等级

### S0：只读分析

允许工具：所有 `docx_list_*`、`docx_read_*`、`docx_search_text`、`docx_map`、`docx_health_check`、`docx_text_diff`。

特点：不写任何文件，适合 Codex 自动调用。

### S1：sidecar 写入

允许工具：`docx_write_index`、`docx_export_plain_text`、`docx_export_table_csv`、`docx_table_to_csv`、`docx_backup`。

特点：不修改 DOCX，只写索引、CSV、TXT、备份等辅助文件。

### S2：内容控件内文本改动

允许操作：`replace_content_control_text`、`replace_text_in_content_control`、`append_content_control_text`、`prepend_content_control_text`。

要求：必须有 `source_sha256` 和 `expected_old_sha256`；默认禁止复杂内容控件。

### S3：局部结构改动

允许操作：`replace_paragraph_text`、`insert_paragraph_before/after`、`replace_table_cell_text`、`append_table_row`、`add_comment`、`wrap_paragraph_with_content_control`。

要求：必须 assess + dry-run + backup + apply + validate；新增内容控件 tag、表格维度变化、批注 part 变化必须显式授权。

### S4：高风险 Word 原生操作

包括刷新目录/字段、接受/拒绝修订、改页眉页脚、改分节、改编号体系、移动图片、处理复杂交叉引用。

当前 Python MCP 不直接执行 S4。生产系统应交给 Word COM、Office.js 人工界面、Aspose/Syncfusion 或 Open XML SDK 专用内核，并启用人工审批。

## 默认阻断条件

- 源文件 hash 与 `source_sha256` 不一致。
- 目标锚点不存在或不唯一。
- `expected_old_sha256` 不匹配。
- 写入后出现未授权 package part 变化。
- 未触达内容控件、表格或带 `paraId` 段落 hash 变化。
- 表格、图片、字段、批注、修订痕迹、内容控件数量出现未授权变化。
- 输出文件已存在但未显式 `allow_overwrite`。

## 推荐给 Codex 的写入策略

1. 先使用 `docx_health_check` 和 `docx_map` 判断文档是否适合 AI 编辑。
2. 目标有内容控件时，只使用 tag 写入。
3. 目标无内容控件时，先用 `wrap_paragraph_with_content_control` 做锚点治理，再进行后续正文改写。
4. 表格只先支持单元格替换；新增/删除行列需要人工审批。
5. 添加新章节时优先插入到已有标题锚点之后，并保留原标题编号体系。
6. 每次 PatchSet 尽量小：一个章节或一组相关内容控件，不要跨越过多文档区域。
