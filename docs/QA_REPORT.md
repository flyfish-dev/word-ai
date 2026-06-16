# QA Report — v0.8.1 Local Verification

Verification date: 2026-06-15.

## Environment

- Python: 3.14.5 in local `.venv`.
- Node.js: v26.0.0.
- npm: 11.12.1.
- .NET SDK: 8.0.128 via Homebrew `dotnet@8`; runtime 8.0.28; pinned by `global.json`.
- Docker: 29.3.1.

## Automated Checks

Executed locally from the project root:

```bash
. .venv/bin/activate
python -m compileall word_ai_mcp scripts
PYTHONPATH=. python scripts/run_smoke_test.py
PYTHONPATH=. python scripts/run_structure_regression.py
PYTHONPATH=. python scripts/run_outline_regression.py
PYTHONPATH=. python scripts/run_word_session_smoke.py
PYTHONPATH=. python scripts/run_dotnet_regression.py
PYTHONPATH=. python scripts/run_office_bridge_smoke.py
PYTHONPATH=. python scripts/validate_word_ai_skill.py
dotnet build dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj -c Release
bash scripts/install.sh
WORD_AI_BRIDGE_PORT=8876 WORD_AI_TASKPANE_PORT=3100 bash scripts/start.sh --http
mcp-publisher validate
docker build -t word-ai:0.8.1-local .
docker run --rm -i -v "$PWD:/workspace" word-ai:0.8.1-local
PYTHONPATH=. python scripts/build_mcpb.py --version 0.8.1 --out dist/word-ai-0.8.1.mcpb
npx -y @anthropic-ai/mcpb validate <extracted-manifest.json>
```

Results:

- Python compilation: passed.
- Smoke test: passed.
- Structure regression: passed.
- Outline regression: passed for localized heading styles, numeric heading style IDs, direct `w:outlineLvl`, and TOC field exclusion.
- Word session command-queue smoke: passed.
- .NET build: passed with 0 warnings and 0 errors.
- .NET PatchSet regression: passed.
- Office bridge smoke: passed.
- Word AI Skill validation: passed.
- One-command install script: passed.
- One-command start script: passed on alternate ports `8876` and `3100`.
- Multi-root path policy smoke: passed; external absolute paths are rejected by default and allowed only with `--allow-root`.
- Advertised MCP tools: 63.
- Stdio MCP handshake and `tools/list`: passed.
- Basic HTTP adapter `/health` and `/mcp tools/list`: passed on `127.0.0.1:8765`.
- Office bridge endpoints verified through smoke: `/office/read`, `/office/build-patchset`, `/office/assess-patchset`, `/office/preview-patchset`, `/office/apply-patchset`.
- Live Word session queue endpoints verified through smoke: session register, queued command claim, command completion, command status, apply command enqueue, rollback command enqueue.
- Optional OfficeCLI wrappers are present as allowlisted tools: `officecli_view_html`, `officecli_view_screenshot`, `officecli_view_issues`, `officecli_query`, `officecli_validate`.
- Official MCP Registry metadata validation: passed for `server.json`.
- Local OCI image build: passed for `word-ai:0.8.1-local`.
- OCI image labels: `io.modelcontextprotocol.server.name=io.github.flyfish-dev/word-ai`, `org.opencontainers.image.licenses=AGPL-3.0-or-later`.
- Container stdio MCP handshake: passed with `serverInfo.version=0.8.1`.
- Container `tools/list`: passed with 63 tools.
- Deterministic MCPB build: passed for `dist/word-ai-0.8.1.mcpb`.
- MCPB manifest validation with `@anthropic-ai/mcpb`: passed.
- MCPB bootstrap first-run stdio test: passed with `serverInfo.version=0.8.1` and 63 tools. Dependency installation logs are routed to stderr so stdout remains valid MCP JSON.

## Global Distribution Checks

- `server.json` uses official schema `https://static.modelcontextprotocol.io/schemas/2025-12-11/server.schema.json`.
- MCP server name is `io.github.flyfish-dev/word-ai`.
- Package type is MCPB with identifier `https://github.com/flyfish-dev/word-ai/releases/download/v0.8.1/word-ai-0.8.1.mcpb`.
- MCPB `fileSha256` is `efc091250b81364cc7c02ba56966f9b6dc090e03c04f191a0f976be0e155f53f`.
- Transport is `stdio`.
- Release workflow `.github/workflows/release-mcp.yml` builds a deterministic MCPB, uploads it to the GitHub Release, then runs `mcp-publisher login github-oidc`, `mcp-publisher validate`, and `mcp-publisher publish`.
- Project license metadata is `AGPL-3.0-or-later`.
- MCP Registry/MCPB is the primary global distribution path for MCP host discovery.
- npm is a secondary convenience channel. Published packages were verified separately as `@flyfish-dev/word-ai@0.8.1` and unscoped `word-ai-mcp@0.8.2`.

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

## Outline Regression Evidence

The outline recognizer now uses `styles.xml` style names, paragraph/direct `w:outlineLvl`, and TOC field/style exclusion instead of relying only on raw style IDs.

Generated regression fixture:

- Numeric style IDs `1` / `2` with style names `heading 1` / `heading 2`: recognized as levels 1 and 2.
- Chinese style names `标题1` / `标题 2`: recognized as levels 1 and 2.
- Direct paragraph `w:outlineLvl=2`: recognized as level 3.
- TOC heading/result paragraphs and complex `TOC \o "1-3" \h \z \u` field range: excluded from heading anchors and marked `is_toc=true` in paragraph inventory.
- Malformed/unclosed complex TOC field: guarded so the leaked field state does not consume body headings after the visible TOC block.
- .NET `inspect` regression on the same fixture: `heading_count=6` with no TOC headings.

External real-document verification was also run against four local Chinese DOCX samples outside the repository:

- A user-manual sample with numeric heading style IDs: `docx_get_outline` returned 83 headings; TOC paragraphs 13-24 were `is_toc=true` and `heading_level=null`; the first real heading appeared after the TOC.
- A software-process specification sample: returned 18 headings using localized style names `标题1` / `标题2`; the TOC page was excluded.
- A large database-design sample: returned 106 headings; the TOC page was excluded while body headings beginning with Chinese words such as `目录...` remained valid when not inside TOC scope.
- A large interface-specification sample: returned 112 headings from numeric heading style IDs.

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

Word session smoke:

- A mock Office.js taskpane session was registered under `.wordai/sessions`.
- `word_session_list` and `word_session_snapshot` returned the active open-document snapshot.
- `word_session_refresh` queued a command and `word_session_command_status` observed completion.
- `word_session_apply_patchset` queued a live apply command for a content-control PatchSet.
- `word_session_rollback` loaded the apply command's generated rollback PatchSet and queued the rollback command.

Browser verification against `http://localhost:3000/taskpane.html`:

- Page loaded with no console errors or warnings.
- Bridge connected through same-origin `/bridge` proxy.
- `Load Anchors` populated all 3 sample content controls.
- `Build` generated a strict PatchSet with source and target hashes.
- `Assess` and `Dry Run` completed successfully from the taskpane UI.

## Not Covered In This Pass

- Automated Word rendering / visual diff: not run in this environment.
- Word desktop sideload and true Office.js host operations (`Wrap`, `List Open`, `Apply Open`, `word_session_apply_patchset`) were not fully automated in this browser-only pass. The TypeScript host code compiles, and the local queue protocol is covered by smoke tests; a real Word host is still required to execute `Word.run(...)`.
- HTTPS dev server certificate flow (`npm run dev`) was not exercised because the automated check used `npm run dev:http`; manifest HTTPS validation passed.
- True MCP Streamable HTTP compatibility: current HTTP adapter is a basic local JSON-RPC adapter, not a production remote MCP transport.
- Future MCP Registry publications depend on pushing a `v*` tag and the GitHub Actions release workflow completing successfully.
