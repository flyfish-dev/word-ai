# v0.4 结构稳定优先升级记录

## 主要增强

- MCP tool surface 扩展到 49 个工具，覆盖包结构、健康检查、结构指纹、标题大纲、内容控件、段落、表格、字段、图片、超链接、批注、脚注尾注、修订痕迹、sidecar 导出、PatchSet 评估、dry-run、备份、正式写入、结构验证和回滚。
- `docx_apply_patchset` 改为候选写入事务：先写 candidate，验证通过后提交最终 DOCX；验证失败时阻断，并写失败审计。
- 增加 `source_sha256`、`expected_old_sha256`、`expected_old_text`、`guard.require_preconditions`、`guard.allow_overwrite` 等并发和误写防护。
- 增加非目标对象 hash 保护：未触达内容控件、表格、带 `w14:paraId` 的段落默认不允许变化。
- 增加受控结构操作：表格追加行、段落前后插入、批注添加、段落包装内容控件，用于满足更多 Word 编辑场景，但都需要通过 assess/dry-run/validate。
- 增加 `docs/STABILITY_POLICY.md`，明确 S0-S4 稳定等级和阻断条件。

## 已验证

- `scripts/run_smoke_test.py` 通过：内容控件替换只改变授权范围，验证 `ok=true`。
- `scripts/run_structure_regression.py` 通过：覆盖表格单元格替换、表格追加行、添加批注、段落插入、段落包装内容控件。
- MCP stdio `initialize`、`tools/list`、`tools/call` smoke test 通过。
- 示例 DOCX 已用 LibreOffice headless 渲染为 PNG 并完成视觉检查。
