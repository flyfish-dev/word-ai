# QA Report — v0.5 Local Verification

Verification date: 2026-06-15.

## Environment

- Python: 3.14.5 in local `.venv`.
- Node.js: v26.0.0.
- npm: 11.12.1.
- .NET SDK: 8.0.128 via Homebrew `dotnet@8`; runtime 8.0.28; pinned by `global.json`.

## Automated Checks

Executed locally from the project root:

```bash
. .venv/bin/activate
python -m compileall word_ai_mcp scripts
PYTHONPATH=. python scripts/run_smoke_test.py
PYTHONPATH=. python scripts/run_structure_regression.py
PYTHONPATH=. python scripts/run_dotnet_regression.py
PYTHONPATH=. python scripts/run_office_bridge_smoke.py
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
```

Results:

- Python compilation: passed.
- Smoke test: passed.
- Structure regression: passed.
- .NET build: passed with 0 warnings and 0 errors.
- .NET PatchSet regression: passed.
- Office bridge smoke: passed.
- Advertised MCP tools: 49.
- Stdio MCP handshake and `tools/list`: passed.
- Basic HTTP adapter `/health` and `/mcp tools/list`: passed on `127.0.0.1:8765`.
- Office bridge endpoints verified through smoke: `/office/read`, `/office/build-patchset`, `/office/assess-patchset`, `/office/preview-patchset`, `/office/apply-patchset`.

## Structural Validation Evidence

The standard content-control replacement workflow produced:

- `examples/sample_contract.edited.docx`
- `examples/sample_contract.edited.audit.json`

Validation summary:

- `ok = true`.
- Changed DOCX parts: `word/document.xml` only.
- Paragraph/table/image/field/comment/comment-reference/tracked-change/content-control/heading counts stayed stable.
- Protected body block sequence stayed stable outside touched scopes.
- Untouched content controls, tables, table cells, and protected paragraphs passed hash checks.

## Regression Coverage

The regression script now exercises:

- `replace_table_cell_text`
- `replace_paragraph_text`
- `append_table_row`
- `add_comment`
- `insert_paragraph_after`
- `wrap_paragraph_with_content_control`

All scenarios passed validation with explicit `source_sha256` and target-level `expected_old_sha256` for high-risk operations.

## .NET Open XML SDK Checks

Executed locally:

```bash
dotnet --info
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
dotnet run --project dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release -- assess examples/sample_contract.docx examples/patches/replace_srs_sections.json
dotnet run --project dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release -- dry-run examples/sample_contract.docx examples/patches/replace_srs_sections.json false
dotnet run --project dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release -- apply examples/sample_contract.docx examples/patches/replace_srs_sections.json examples/sample_contract.dotnet.edited.docx
dotnet run --project dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release -- validate examples/sample_contract.docx examples/sample_contract.dotnet.edited.docx
```

Results:

- SDK/runtime: `.NET SDK 8.0.128`, `Microsoft.NETCore.App 8.0.28`, `Microsoft.AspNetCore.App 8.0.28`, `osx-arm64`.
- Build: passed, 0 warnings, 0 errors.
- `assess`: `ok=true`, 3 content-control operations resolved.
- `dry-run`: `validation.ok=true`, output removed when `keep_output=false`.
- `apply`: wrote `examples/sample_contract.dotnet.edited.docx` and `examples/sample_contract.dotnet.edited.audit.json`.
- `validate`: `ok=true`, changed parts `["word/document.xml"]`, new OpenXmlValidator errors `0`.
- `scripts/run_dotnet_regression.py`: passed all C# PatchSet operations:
  `replace_content_control_text`, `replace_text_in_content_control`, `append_content_control_text`, `prepend_content_control_text`, `replace_table_cell_text`, `replace_paragraph_text`, `append_table_row`, `add_comment`, `insert_paragraph_after`, `insert_paragraph_before`, `wrap_paragraph_with_content_control`.

## MCP Safety Checks

- Read-only server mode rejects sidecar and write tools: `docx_write_index`, `docx_backup`, `docx_export_plain_text`.
- `docx_validate` can auto-load audit scope for table-cell edits and paragraph-index edits when the document has no `w14:paraId`.
- Missing high-risk operation preconditions are reported as `missing_precondition` errors during `docx_assess_patchset`.

## Office Add-in Checks

Executed from `office-addin/`:

```bash
npm install
npm run build
npx office-addin-manifest validate manifest.xml
npm audit --omit=dev --registry=https://registry.npmjs.org --json
```

Results:

- TypeScript build: passed.
- Office manifest validation: passed.
- Production dependency audit: 0 vulnerabilities.
- Full dev dependency audit: 13 dev-only advisories remain through `office-addin-debugging` and its transitive dependencies.

## Office Bridge / Taskpane Checks

Executed locally:

```bash
PYTHONPATH=. python -m word_ai_mcp.server_http --root "$PWD" --host 127.0.0.1 --port 8765 --token <local-dev-token>
cd office-addin
npm run dev:http
```

Automated smoke:

- Missing token on `/office/read`: rejected with `401`.
- `/office/read`: returned sample document health and 3 content controls.
- `/office/build-patchset`: generated `source_sha256` and operation-level `expected_old_sha256`.
- `/office/assess-patchset`: `ok=true`.
- `/office/preview-patchset`: dry-run validation `ok=true`.
- `/office/apply-patchset`: wrote `examples/.wordai/bridge-smoke/sample_contract.bridge.docx` and adjacent audit JSON.
- Final validation: `ok=true`, changed parts include `word/document.xml`.
- JSON-RPC `/mcp` write tool without token: rejected with `401`.

Browser verification against `http://localhost:3000/taskpane.html`:

- Page loaded with no console errors or warnings.
- Bridge connected through same-origin `/bridge` proxy.
- `Load Anchors` populated all 3 sample content controls.
- `Build` generated a strict PatchSet with source and target hashes.
- `Assess` and `Dry Run` completed successfully from the taskpane UI.

## Not Covered In This Pass

- Automated Word rendering / visual diff: not run in this environment.
- Word desktop sideload and true Office.js host operations (`Wrap`, `List Open`, `Apply Open`) were not automated in this browser-only pass.
- HTTPS dev server certificate flow (`npm run dev`) was not exercised because the automated check used `npm run dev:http`; manifest HTTPS validation passed.
- True MCP Streamable HTTP compatibility: current HTTP adapter is a basic local JSON-RPC adapter, not a production remote MCP transport.
