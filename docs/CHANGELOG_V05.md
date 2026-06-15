# v0.5 Office.js 闭环升级记录

## 核心变化

- 新增 Office bridge REST 层：`/office/read`、`/office/build-patchset`、`/office/assess-patchset`、`/office/preview-patchset`、`/office/apply-patchset`。
- Office bridge 底层仍调用标准 MCP tools，不绕过 PatchSet、dry-run、backup、validate、diff。
- `/office/*` POST 和 `/mcp` 写工具要求 `X-Word-AI-Token`。
- 未显式传入 token 时，`server_http` 会在 `.wordai/bridge.token` 创建本地随机 token。
- CORS 限制到 localhost/127.0.0.1 开发源，并支持 taskpane 同源 `/bridge/*` 代理。
- Office taskpane 从锚点骨架升级为操作台：连接 bridge、读取锚点、生成 PatchSet、assess、dry-run、apply、查看验证和 diff。
- `Apply Open` 支持对当前打开 Word 文档中的内容控件文本执行安全写入，写入前校验 `expected_old_sha256`。
- 新增 `office-addin/scripts/dev-server.js`，默认 HTTPS 服务 taskpane，并把 `/bridge/*` 代理到本地 bridge。
- 新增 `scripts/run_office_bridge_smoke.py`，覆盖 token 拦截、read/build/assess/dry-run/apply/validate/diff。

## 验证

- `python -m compileall word_ai_mcp scripts`
- `PYTHONPATH=. python scripts/run_office_bridge_smoke.py`
- `npm run build`

完整回归仍应包含：

- `PYTHONPATH=. python scripts/run_smoke_test.py`
- `PYTHONPATH=. python scripts/run_structure_regression.py`
- `PYTHONPATH=. python scripts/run_dotnet_regression.py`
- `dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release`
