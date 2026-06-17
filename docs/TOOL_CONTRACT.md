# MCP 工具契约 v0.7：结构稳定优先

本契约面向 Codex / OpenAI Agents / MCP 客户端。设计目标是：让模型拥有足够丰富的读取、定位和规划能力，但把所有正式 DOCX 写入都收口到受控 PatchSet 事务中。离线文件编辑使用 `docx_*` 工具链；当前打开的 Word 文档编辑使用 Office.js taskpane 注册的 `word_session_*` 工具链。

## 总原则

1. **读工具丰富，写工具收口**：Codex 可以广泛读取结构、索引、锚点、表格、字段、批注和修订信息；文件级正文写入只能通过 `docx_apply_patchset`，打开的 Word 会话写入只能通过 `word_session_apply_patchset`。
2. **默认不覆盖源文件**：输出新 DOCX；覆盖必须显式授权，并先备份。
3. **候选写入再提交**：写工具先生成 candidate，结构验证通过后才提交最终文件。
4. **前置条件优先**：关键修改必须带 `source_sha256`、`expected_old_text` 或 `expected_old_sha256`。
5. **复杂对象默认保护**：目标含表格、图片、字段、批注、修订痕迹、嵌套内容控件时，替换操作默认风险上报或拒绝，除非显式 `allow_complex_content=true`。
6. **不变量验证阻断交付**：未授权 part 变化、非目标对象 hash 变化、表格/图片/字段/批注/修订计数异常，均导致 `ok=false`。
7. **多 root 路径白名单**：相对路径基于主 `--root` 解析；外部绝对路径必须位于显式 `--allow-root` 或 `WORD_AI_ALLOWED_ROOTS` 白名单内。

## 离线引擎选择

离线文件事务的默认后端是 .NET Open XML SDK。Python 负责 MCP facade、路径策略、读取索引、Office.js bridge 和 fallback/reference。

`WORD_AI_ENGINE=auto|dotnet|python` 控制默认行为：

- `auto`：优先 .NET；找不到 native executable、Release DLL 或源码 project 时才回退 Python。
- `dotnet`：强制 .NET，后端缺失或失败时直接报错，推荐生产环境使用。
- `python`：强制 Python fallback，主要用于对照测试和开发调试。

单次工具调用也可以在 `docx_assess_patchset`、`docx_dry_run_patchset`、`docx_apply_patchset`、`docx_validate` 中传入 `engine` 覆盖默认值。返回结果必须包含实际 `engine`，便于审计。

## 推荐调用链

```text
docx_health_check
  -> docx_map / docx_list_anchors / docx_search_text
  -> docx_read_content_control / docx_read_paragraph / docx_read_table_cell
  -> docx_assess_patchset 或 docx_plan_patchset
  -> docx_dry_run_patchset 或 docx_preflight_patchset
  -> docx_backup
  -> docx_apply_patchset
  -> docx_validate 或 docx_compare_structure
  -> docx_text_diff
```

## Office.js Bridge 调用链

Office taskpane 使用 `word_ai_mcp.server_http` 的 `/office/*` REST 封装。文件级流程底层仍调用本契约中的 `docx_*` MCP tools：

```text
/office/read
  -> docx_inspect + docx_health_check + docx_list_content_controls + complex-object lists
/office/build-patchset
  -> docx_inspect + docx_read_content_control
/office/assess-patchset
  -> docx_health_check + docx_assess_patchset
/office/preview-patchset
  -> docx_health_check + docx_assess_patchset + docx_dry_run_patchset
/office/apply-patchset
  -> docx_health_check
  -> docx_assess_patchset
  -> docx_dry_run_patchset
  -> docx_backup
  -> docx_apply_patchset
  -> docx_validate
  -> docx_text_diff
```

Bridge 不新增绕过 PatchSet 的文件写入口；所有文件级 DOCX 正文修改仍由 `docx_apply_patchset` 完成。

打开的 Word 会话由 taskpane 注册到 `.wordai/sessions`，Codex 通过 `word_session_*` MCP tools 把命令写入本地队列，taskpane 轮询后在真实 Word host 内执行 Office.js：

```text
Word taskpane -> /office/session/register
Codex -> word_session_list / word_session_snapshot
Codex -> word_session_read_content_control
Codex -> word_session_preview_patchset
Codex -> word_session_apply_patchset
  -> taskpane live preflight
  -> expected_old_sha256 校验
  -> Office.js contentControl.insertText
  -> audit + rollback PatchSet
Codex -> word_session_rollback
```

`word_session_apply_patchset` 只支持内容控件文本类 PatchSet 操作：`replace_content_control_text`、`append_content_control_text`、`prepend_content_control_text`、`replace_text_in_content_control`。它不刷新字段、不导出 PDF、不重建 DOCX，也不直接修改关系、样式、编号、页眉页脚或图片关系。

## OfficeCLI 辅助证据链路

OfficeCLI 是可选辅助后端，不是 Word AI 的权威写入路径。MCP 只暴露以下白名单 wrapper：

- `officecli_view_html`：`officecli view <file> html`，读取 HTML 渲染快照。
- `officecli_view_screenshot`：`officecli view <file> screenshot -o <output>`，生成 PNG sidecar；属于低风险 sidecar 写入，建议审批。
- `officecli_view_issues`：`officecli view <file> issues --json`，读取格式/内容/结构问题。
- `officecli_query`：`officecli query <file> <selector> --json`，读取语义路径查询结果。
- `officecli_validate`：`officecli validate <file> --json`，作为 Word AI validate 之外的辅助证据。

默认不暴露 OfficeCLI `set`、`add`、`remove`、`move`、`swap`、`batch`、`raw-set`、`create`、`merge`、`dump` 等变更能力。若未来要使用这些能力，必须先包进 Word AI PatchSet、assess、dry-run、audit、rollback 和显式 approval。

## 工具清单

当前 `tools/list` 暴露 **63 个工具**，包含常用别名，便于 Codex 以自然名称调用。

### 包结构与全局健康检查

| 工具 | 用途 |
|---|---|
| `docx_health_check` | 保守稳定性报告：重复 tag、复杂内容控件、字段、批注、修订、表格风险。 |
| `docx_inspect` | 基础计数、文件 hash、anchors、可选可见文本。 |
| `docx_package_manifest` | 列出 DOCX ZIP part、大小和可选 SHA-256。 |
| `docx_list_parts` | 轻量列出所有 DOCX package parts。 |
| `docx_part_hashes` | before/after 对比用 part hash 清单。 |
| `docx_structural_fingerprint` | 结构指纹：part hash、内容控件、段落、表格、大纲和健康检查。 |
| `docx_map` | Codex 友好的结构地图。 |

### 导航、检索与锚点

| 工具 | 用途 |
|---|---|
| `docx_get_outline` / `docx_outline` | 标题大纲。 |
| `docx_list_headings` | 标题列表、样式、层级和 hash。 |
| `docx_read_heading_section` | 读取某个标题下的章节块。 |
| `docx_list_anchors` | 内容控件、标题、书签、段落锚点。 |
| `docx_read_anchor` | 按 anchor_id 读取目标内容。 |
| `docx_search_text` | 搜索可见文本，支持 regex。 |
| `docx_list_paragraphs` | 段落索引、paraId、样式、scope、hash。 |
| `docx_read_paragraph` | 读取精确段落。 |
| `docx_list_content_controls` | 列内容控件、tag、title、复杂度和 hash。 |
| `docx_read_content_control` | 按内容控件 tag 读取目标文本。首选。 |
| `docx_extract_plain_text` | 提取全文纯文本，只用于检索/审阅，不用于重建。 |
| `docx_export_plain_text` | 导出全文 sidecar `.txt`，不修改 DOCX。 |

标题识别必须优先遵守 Word 的结构语义，而不是只看文本：

- `docx_get_outline` / `docx_list_headings` 会读取 `styles.xml`，支持 `Heading1`、`标题1`、数字 styleId 但样式名为 `heading 1` 的中文/WPS 文档，以及带 `w:outlineLvl` 的自定义标题样式。
- TOC/目录结果必须从 outline 中排除。`TOC1`、`toc 1`、`目录`、`WPSOffice 手动目录`、`TOC Heading` 样式，以及复杂 TOC 字段范围内的段落，不能作为 heading anchor。
- TOC 字段不能无限扩张：遇到缺失 `end` 的复杂 TOC 字段时，识别器只跳过显式 TOC 样式、TOC 指令或 `_Toc` 引用段落，不能把后续正文 heading 全部视为 TOC。
- `docx_list_paragraphs` / `docx_read_paragraph` 会返回 `style_name` 和 `is_toc`，用于区分目录结果与正文标题。`is_toc=true` 的段落不应作为编辑锚点。

### 表格、样式和复杂对象

| 工具 | 用途 |
|---|---|
| `docx_list_tables` | 表格定位、维度、预览和 hash。 |
| `docx_read_table` | 读取指定表格矩阵。 |
| `docx_read_table_cell` | 精读指定表格单元格并返回 hash。 |
| `docx_export_table_csv` | 表格导出 CSV，属于 sidecar 写入，不改 DOCX。 |
| `docx_table_to_csv` | 表格转 CSV，支持默认 sidecar 输出。 |
| `docx_list_styles` | 样式清单和依赖关系。 |
| `docx_list_numbering` | 编号定义和层级。 |
| `docx_list_sections` | 节属性、页边距、方向、页眉页脚引用。 |
| `docx_list_fields` | TOC、REF、PAGEREF、SEQ 等字段。 |
| `docx_list_images` | 图片和 drawing 对象。 |
| `docx_list_hyperlinks` | 超链接和关系 ID。 |
| `docx_list_comments` / `docx_extract_comments` | 批注。 |
| `docx_list_bookmarks` | 书签。 |
| `docx_list_headers_footers` | 页眉页脚 part、文本预览和 hash。 |
| `docx_list_notes` | 脚注/尾注文本和 hash。 |
| `docx_list_revisions` | 修订痕迹。 |

### PatchSet、写入、验证和回滚

| 工具 | 用途 |
|---|---|
| `docx_assess_patchset` | 静态评估 PatchSet 风险、目标解析和触达对象。 |
| `docx_plan_patchset` | `docx_assess_patchset` 的规划别名。 |
| `docx_preflight_patchset` | 临时写入并验证，默认删除临时输出。 |
| `docx_dry_run_patchset` | 与 preflight 等价，便于 Codex 显式 dry-run。 |
| `docx_write_index` | 写 `.wordai` sidecar index，不改 DOCX。 |
| `docx_backup` | 创建备份。 |
| `docx_restore_backup` | 将备份恢复到指定目标。 |
| `docx_rollback` | 从备份回滚，可先备份当前文件。 |
| `docx_apply_patchset` | 唯一正式正文写入入口。 |
| `docx_validate` | 结构验证，可传入 touched_* 说明授权修改范围。 |
| `docx_compare_structure` | 与 validate 同等的结构比较报告。 |
| `docx_text_diff` | 可见文本 unified diff。 |

### Word 会话与 Office.js 实时编辑

| 工具 | 用途 |
|---|---|
| `word_session_list` | 列出已由 Office.js taskpane 注册的 active Word sessions。 |
| `word_session_snapshot` | 返回打开文档最近一次心跳中的内容控件快照。 |
| `word_session_refresh` | 要求 taskpane 重新扫描当前打开文档的内容控件。 |
| `word_session_read_content_control` | 通过 Office.js 读取当前打开文档中的内容控件文本和 hash。 |
| `word_session_preview_patchset` | 在 Word host 内对 PatchSet 做 live preflight，不写入。 |
| `word_session_apply_patchset` | 在 Word host 内写入内容控件文本，写入前自动 live preflight，返回 audit 和 rollback PatchSet。 |
| `word_session_wrap_selection` | 将当前 Word 选区包装为带稳定 tag/title 的内容控件。 |
| `word_session_rollback` | 使用上一条 live apply 生成的 rollback PatchSet 回滚打开文档。 |
| `word_session_command_status` | 查询 session 命令队列中的命令状态、结果或错误。 |

### OfficeCLI 辅助证据

| 工具 | 用途 |
|---|---|
| `officecli_view_html` | 读取 HTML 渲染快照，用于视觉证据。 |
| `officecli_view_screenshot` | 输出 PNG screenshot sidecar，不修改 DOCX。 |
| `officecli_view_issues` | 读取 OfficeCLI issues JSON。 |
| `officecli_query` | 只读语义路径查询，强制 `--json`。 |
| `officecli_validate` | OfficeCLI OpenXML validate JSON，作为辅助验证。 |

## PatchSet 操作白名单

所有 PatchSet 入口都会先做轻量归一化，兼容部分 Agent 常见写法，例如 `operation` / `operation_type` / `type` / `action` 作为 `op`，`replaceContentControlText` / `replace-content-control-text` 作为 `replace_content_control_text`，以及 `target_tag`、`new_text`、`text_sha256` 等字段别名。归一化只发生在输入边界，不新增写入操作，不跳过 `assess`、`dry-run`、hash 前置条件、验证、审计和回滚。正式文档和技能仍应优先生成下列 canonical PatchSet。

### 内容控件替换

```json
{"op":"replace_content_control_text","tag":"WORD-AI:SRS:1.0:overview","expected_old_sha256":"...","text":"新文本"}
```

### 内容控件短语替换

```json
{"op":"replace_text_in_content_control","tag":"WORD-AI:SRS:1.0:overview","find":"旧词","replace":"新词","expected_old_sha256":"..."}
```

### 内容控件追加/前置

```json
{"op":"append_content_control_text","tag":"WORD-AI:SRS:3.0:nfr_performance","expected_old_sha256":"...","text":"新增约束。"}
```

### 段落替换/插入

```json
{"op":"replace_paragraph_text","paragraph_index":12,"expected_old_sha256":"...","text":"新段落"}
{"op":"insert_paragraph_after","paragraph_index":12,"text":"新增段落"}
```

### 表格单元格和行

```json
{"op":"replace_table_cell_text","table_index":2,"row":3,"col":4,"expected_old_sha256":"...","text":"新说明"}
{"op":"append_table_row","table_index":2,"template_row":3,"expected_old_sha256":"<table_text_sha256>","values":["字段","类型","约束","说明"]}
```

### 锚点治理和批注

```json
{"op":"wrap_paragraph_with_content_control","paragraph_index":18,"tag":"WORD-AI:SRS:4.1:database_design","title":"数据库设计正文锚点"}
{"op":"add_comment","paragraph_index":20,"text":"请人工确认该性能指标。","author":"Word AI"}
```

## 验证项目

- ZIP 完整性。
- part add/remove/change 是否在 allow-list 内。
- table/image/field/comment/comment-reference/tracked-change/content-control/heading 等计数是否符合预期。
- 内容控件 tag 集合是否稳定；新增 tag 必须通过 `allowed_added_content_control_tags` 显式授权。
- 未触达内容控件 XML hash 是否不变。
- 未触达表格 hash 是否不变。
- 未触达且有 `w14:paraId` 的段落 hash 是否不变。
- 写入审计是否包含 source/target hash、applied operations、validation、changed parts。

## Codex 硬规则

- 优先内容控件 tag。
- 离线文件写入必须先 `docx_assess_patchset` 和 `docx_dry_run_patchset`。
- Word 会话写入必须先 `word_session_preview_patchset`；即使直接调用 `word_session_apply_patchset`，taskpane 也必须先执行 live preflight。
- 离线文件关键交付内容、段落兜底操作、表格结构操作、批注和锚点治理必须带 `source_sha256` 和目标级 `expected_old_sha256`。
- Word 会话内容控件写入必须带目标级 `expected_old_sha256`，并由 taskpane 对打开文档实时计算 hash。
- 不允许全量重建 DOCX。
- 不允许无人工审批启用 `allow_complex_content=true`、`allow_overwrite=true` 或覆盖源文件。
