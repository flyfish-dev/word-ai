type JsonObject = Record<string, any>;

type OpenControl = {
  id: number | string;
  tag: string;
  title: string;
  text: string;
  text_sha256: string;
};

type SessionCommand = {
  command_id: string;
  session_id: string;
  type: string;
  payload?: JsonObject;
};

let openControls: OpenControl[] = [];
let sessionId = localStorage.getItem("wordAiSessionId") || "";
let heartbeatTimer: number | undefined;
let pollTimer: number | undefined;
let pollingCommands = false;

function el<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) {
    throw new Error(`Missing element: ${id}`);
  }
  return node as T;
}

function value(id: string): string {
  return el<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(id).value.trim();
}

function setValue(id: string, next: string): void {
  el<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(id).value = next;
}

function setStatus(id: string, message: string, kind: "idle" | "ok" | "error" = "idle"): void {
  const node = el<HTMLDivElement>(id);
  node.className = kind === "idle" ? "status" : `status ${kind}`;
  node.textContent = message;
}

function setSessionStatus(message: string, kind: "idle" | "ok" | "error" = "idle"): void {
  setStatus("sessionStatus", message, kind);
}

function log(message: string, payload?: unknown): void {
  const out = el<HTMLTextAreaElement>("log");
  const text = payload === undefined ? message : `${message}\n${JSON.stringify(payload, null, 2)}`;
  out.value = `${new Date().toISOString()} ${text}\n\n${out.value}`;
}

function officeApi(): any {
  return (globalThis as any).Office;
}

function wordApi(): any {
  return (globalThis as any).Word;
}

function wordAvailable(): boolean {
  return Boolean(officeApi() && wordApi() && typeof wordApi().run === "function");
}

async function runAction(name: string, action: () => Promise<void>): Promise<void> {
  try {
    setStatus("resultStatus", `${name}...`);
    await action();
    setStatus("resultStatus", `${name} complete`, "ok");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setStatus("resultStatus", message, "error");
    log(`${name} failed: ${message}`);
  }
}

async function sha256(text: string): Promise<string> {
  const data = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

function bridgeBase(): string {
  const base = value("bridgeUrl") || "/bridge";
  return base.replace(/\/+$/, "");
}

async function bridgeFetch(path: string, body?: JsonObject, method = "POST"): Promise<JsonObject> {
  const token = value("bridgeToken");
  const headers: Record<string, string> = {"Content-Type": "application/json"};
  if (token) {
    headers["X-Word-AI-Token"] = token;
  }
  const response = await fetch(`${bridgeBase()}${path}`, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const text = await response.text();
  let payload: JsonObject = {};
  if (text) {
    try {
      payload = JSON.parse(text);
    } catch {
      payload = {raw: text};
    }
  }
  if (!response.ok || payload.ok === false) {
    throw new Error(payload.error || payload.message || `HTTP ${response.status}`);
  }
  return payload;
}

function parsePatchset(): JsonObject {
  const raw = value("patchset");
  if (!raw) {
    throw new Error("PatchSet is empty");
  }
  return JSON.parse(raw);
}

function patchsetOperations(patchset: JsonObject): JsonObject[] {
  const operations = patchset.operations;
  if (!Array.isArray(operations) || operations.length === 0) {
    throw new Error("PatchSet operations are empty");
  }
  return operations as JsonObject[];
}

function currentDocxPath(): string {
  const path = value("docxPath");
  if (!path) {
    throw new Error("Source DOCX is required");
  }
  return path;
}

function currentTag(): string {
  const tag = value("fileTag") || value("tag");
  if (!tag) {
    throw new Error("Target tag is required");
  }
  return tag;
}

function renderFileControls(items: JsonObject[]): void {
  const select = el<HTMLSelectElement>("fileTag");
  select.innerHTML = "";
  for (const item of items) {
    const option = document.createElement("option");
    option.value = item.tag || "";
    option.textContent = `${item.tag || "(untagged)"} - ${(item.alias || item.text_preview || "").slice(0, 60)}`;
    select.appendChild(option);
  }
  if (items.length > 0) {
    select.value = items[0].tag || "";
  }
}

function renderOpenControls(): void {
  const list = el<HTMLUListElement>("openControls");
  list.innerHTML = "";
  for (const item of openControls) {
    const row = document.createElement("li");
    row.textContent = `${item.tag || "(untagged)"} - ${(item.title || item.text || "").slice(0, 80)}`;
    row.onclick = () => {
      setValue("tag", item.tag);
      const select = el<HTMLSelectElement>("fileTag");
      for (const option of Array.from(select.options)) {
        if (option.value === item.tag) {
          select.value = item.tag;
          break;
        }
      }
      log(`Selected open tag: ${item.tag}`);
    };
    list.appendChild(row);
  }
}

function officeDocumentInfo(): JsonObject {
  const office = officeApi();
  return {
    host: office?.context?.host || "Word",
    platform: office?.context?.platform || null,
    url: office?.context?.document?.url || null,
    taskpane_url: window.location.href,
  };
}

async function collectOpenControls(): Promise<OpenControl[]> {
  if (!wordAvailable()) {
    throw new Error("Word host is not available");
  }
  const controls: OpenControl[] = [];
  await wordApi().run(async (context: any) => {
    const collection = context.document.contentControls;
    collection.load("items/id,items/tag,items/title,items/text");
    await context.sync();
    for (const item of collection.items || []) {
      controls.push({
        id: item.id,
        tag: item.tag || "",
        title: item.title || "",
        text: item.text || "",
        text_sha256: await sha256(item.text || ""),
      });
    }
  });
  openControls = controls;
  renderOpenControls();
  return controls;
}

async function connectBridge(): Promise<void> {
  const payload = await bridgeFetch("/office/capabilities", undefined, "GET");
  localStorage.setItem("wordAiBridgeUrl", value("bridgeUrl"));
  localStorage.setItem("wordAiBridgeToken", value("bridgeToken"));
  setStatus("bridgeStatus", `Connected: ${payload.name}`, "ok");
  log("Bridge capabilities", payload);
  if (wordAvailable()) {
    await registerSession();
  }
}

function startSessionLoops(): void {
  if (heartbeatTimer !== undefined) {
    window.clearInterval(heartbeatTimer);
  }
  if (pollTimer !== undefined) {
    window.clearInterval(pollTimer);
  }
  heartbeatTimer = window.setInterval(() => void heartbeatSession(), 5000);
  pollTimer = window.setInterval(() => void pollSessionCommands(), 1500);
}

async function registerSession(): Promise<void> {
  if (!wordAvailable()) {
    throw new Error("Word host is not available");
  }
  const controls = await collectOpenControls();
  const payload = await bridgeFetch("/office/session/register", {
    session_id: sessionId || undefined,
    client: {
      user_agent: navigator.userAgent,
      language: navigator.language,
      started_at: new Date().toISOString(),
    },
    document: officeDocumentInfo(),
    capabilities: {
      live_session: true,
      supported_patchset_ops: [
        "replace_content_control_text",
        "append_content_control_text",
        "prepend_content_control_text",
        "replace_text_in_content_control",
      ],
      supports_wrap_selection: true,
      supports_rollback_patchset: true,
    },
    content_controls: controls,
  });
  sessionId = payload.session?.session_id || sessionId;
  localStorage.setItem("wordAiSessionId", sessionId);
  setSessionStatus(`Live session: ${sessionId}`, "ok");
  startSessionLoops();
  log("Registered live Word session", payload.session);
}

async function heartbeatSession(): Promise<void> {
  if (!sessionId || !wordAvailable()) {
    return;
  }
  try {
    const controls = await collectOpenControls();
    await bridgeFetch("/office/session/heartbeat", {
      session_id: sessionId,
      status: "active",
      document: officeDocumentInfo(),
      content_controls: controls,
    });
    setSessionStatus(`Live session: ${sessionId}`, "ok");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setSessionStatus(`Session heartbeat failed: ${message}`, "error");
  }
}

async function wrapSelection(): Promise<void> {
  const tag = value("tag");
  const title = value("title") || tag;
  if (!tag) {
    throw new Error("Tag is required");
  }
  const result = await wrapSelectionInOpenDocument(tag, title);
  log(`Wrapped selection as ${tag}`, result);
}

async function listOpenControls(): Promise<void> {
  const controls = await collectOpenControls();
  log("Open document controls", controls);
}

async function loadFileControls(): Promise<void> {
  const payload = await bridgeFetch("/office/read", {docx_path: currentDocxPath()});
  const items = payload.content_controls?.content_controls || [];
  renderFileControls(items);
  log("Loaded DOCX anchors", {
    sha256: payload.inspect?.sha256,
    controls: items.length,
    fields: payload.inspect?.field_count,
    revisions: payload.inspect?.tracked_change_count,
  });
}

async function copyOpenTag(): Promise<void> {
  const tag = value("tag");
  if (!tag) {
    throw new Error("Open tag is empty");
  }
  const select = el<HTMLSelectElement>("fileTag");
  let found = false;
  for (const option of Array.from(select.options)) {
    if (option.value === tag) {
      found = true;
      break;
    }
  }
  if (!found) {
    const option = document.createElement("option");
    option.value = tag;
    option.textContent = tag;
    select.appendChild(option);
  }
  select.value = tag;
}

async function buildPatchset(): Promise<void> {
  const operation = value("operation");
  const body: JsonObject = {
    docx_path: currentDocxPath(),
    tag: currentTag(),
    operation,
    text: value("newText"),
    find: value("findText"),
    replace: value("replaceText"),
  };
  const payload = await bridgeFetch("/office/build-patchset", body);
  setValue("patchset", JSON.stringify(payload.patchset, null, 2));
  log("Built PatchSet", payload.patchset);
}

async function assessPatchset(): Promise<void> {
  const payload = await bridgeFetch("/office/assess-patchset", {
    docx_path: currentDocxPath(),
    patchset: parsePatchset(),
  });
  log("PatchSet assessment", payload);
}

async function dryRunPatchset(): Promise<void> {
  const payload = await bridgeFetch("/office/preview-patchset", {
    docx_path: currentDocxPath(),
    patchset: parsePatchset(),
    keep_output: false,
  });
  log("PatchSet dry-run", payload);
  if (!payload.ok) {
    throw new Error("Dry-run validation failed");
  }
}

async function applyPatchset(): Promise<void> {
  if (!window.confirm("Apply PatchSet and write a new DOCX?")) {
    return;
  }
  const body: JsonObject = {
    docx_path: currentDocxPath(),
    patchset: parsePatchset(),
  };
  const outputPath = value("outputPath");
  if (outputPath) {
    body.output_path = outputPath;
  }
  const payload = await bridgeFetch("/office/apply-patchset", body);
  log("PatchSet apply result", payload);
  if (!payload.ok) {
    throw new Error(`Apply stopped at ${payload.stage || "unknown"}`);
  }
  setStatus("resultStatus", `Wrote ${payload.output_path}`, "ok");
}

function supportedLiveOperation(op: JsonObject): boolean {
  return [
    "replace_content_control_text",
    "append_content_control_text",
    "prepend_content_control_text",
    "replace_text_in_content_control",
  ].includes(op.op);
}

function nextText(current: string, op: JsonObject): string {
  if (op.op === "replace_content_control_text") {
    return String(op.text ?? "");
  }
  if (op.op === "append_content_control_text") {
    return `${current}${current ? "\n" : ""}${String(op.text ?? "")}`;
  }
  if (op.op === "prepend_content_control_text") {
    return `${String(op.text ?? "")}${current ? "\n" : ""}${current}`;
  }
  if (op.op === "replace_text_in_content_control") {
    const find = String(op.find ?? "");
    if (!find) {
      throw new Error("Find text is required");
    }
    if (!current.includes(find) && op.require_match !== false) {
      throw new Error(`Find text not found in ${op.tag}`);
    }
    const replace = String(op.replace ?? "");
    return current.split(find).join(replace);
  }
  throw new Error(`Unsupported live operation: ${op.op}`);
}

async function previewPatchsetInOpenDocument(patchset: JsonObject): Promise<JsonObject> {
  const previews: JsonObject[] = [];
  if (!wordAvailable()) {
    throw new Error("Word host is not available");
  }
  const beforeControls = await collectOpenControls();
  await wordApi().run(async (context: any) => {
    for (const op of patchsetOperations(patchset)) {
      if (!supportedLiveOperation(op)) {
        throw new Error(`Unsupported live operation: ${op.op}`);
      }
      if (!op.expected_old_sha256) {
        throw new Error(`Missing expected_old_sha256 for ${op.tag}`);
      }
      const control = context.document.contentControls.getByTag(op.tag).getFirstOrNullObject();
      control.load("id,tag,title,text");
      await context.sync();
      if (control.isNullObject) {
        throw new Error(`Open document content control not found: ${op.tag}`);
      }
      const current = control.text || "";
      const currentHash = await sha256(current);
      if (currentHash !== op.expected_old_sha256) {
        throw new Error(`Open document hash mismatch for ${op.tag}`);
      }
      const next = nextText(current, op);
      previews.push({
        tag: op.tag,
        id: control.id,
        old_sha256: currentHash,
        new_sha256: await sha256(next),
        before: current.slice(0, 500),
        after: next.slice(0, 500),
      });
    }
  });
  return {
    ok: true,
    mode: "word_session",
    session_id: sessionId || null,
    previewed_at: new Date().toISOString(),
    operation_count: previews.length,
    previews,
    validation: {
      ok: true,
      mode: "word_session_preflight",
      checks: [
        "supported_patchset_operations",
        "content_control_exists",
        "expected_old_sha256_matches_open_document",
      ],
      before_content_control_count: beforeControls.length,
    },
  };
}

async function applyPatchsetToOpenDocument(patchset: JsonObject, rollbackOfCommandId?: string): Promise<JsonObject> {
  if (!wordAvailable()) {
    throw new Error("Word host is not available");
  }
  const preflight = await previewPatchsetInOpenDocument(patchset);
  const beforeControls = await collectOpenControls();
  const applied: JsonObject[] = [];
  await wordApi().run(async (context: any) => {
    for (const op of patchsetOperations(patchset)) {
      if (!supportedLiveOperation(op)) {
        throw new Error(`Unsupported live operation: ${op.op}`);
      }
      if (!op.expected_old_sha256) {
        throw new Error(`Missing expected_old_sha256 for ${op.tag}`);
      }
      const control = context.document.contentControls.getByTag(op.tag).getFirstOrNullObject();
      control.load("id,tag,title,text");
      await context.sync();
      if (control.isNullObject) {
        throw new Error(`Open document content control not found: ${op.tag}`);
      }
      const oldText = control.text || "";
      const oldHash = await sha256(oldText);
      if (oldHash !== op.expected_old_sha256) {
        throw new Error(`Open document hash mismatch for ${op.tag}`);
      }
      const newText = nextText(oldText, op);
      const newHash = await sha256(newText);
      control.insertText(newText, wordApi().InsertLocation.replace);
      applied.push({
        op: op.op,
        tag: op.tag,
        id: control.id,
        title: control.title || "",
        old_text: oldText,
        new_text: newText,
        old_sha256: oldHash,
        new_sha256: newHash,
      });
    }
    await context.sync();
  });
  const afterControls = await collectOpenControls();
  const rollbackOps = applied.map((item) => ({
    op: "replace_content_control_text",
    tag: item.tag,
    expected_old_sha256: item.new_sha256,
    text: item.old_text,
    preserve_style: true,
    allow_complex_content: false,
  }));
  const rollbackPatchset = {
    schema_version: "2.0",
    strict: true,
    reason: `Rollback Word session command${rollbackOfCommandId ? ` ${rollbackOfCommandId}` : ""}`,
    guard: {require_preconditions: true, allow_overwrite: false},
    operations: rollbackOps,
  };
  const validation = {
    ok: beforeControls.length === afterControls.length,
    mode: "word_session_apply",
    checks: [
      "supported_patchset_operations",
      "content_control_exists",
      "expected_old_sha256_matches_open_document",
      "content_control_count_stable",
      "rollback_patchset_generated",
    ],
    before_content_control_count: beforeControls.length,
    after_content_control_count: afterControls.length,
    touched_content_control_tags: applied.map((item) => item.tag),
  };
  const audit = {
    ok: validation.ok,
    mode: "word_session",
    session_id: sessionId || null,
    applied_at: new Date().toISOString(),
    rollback_of_command_id: rollbackOfCommandId || null,
    patchset,
    applied: applied.map((item) => ({
      op: item.op,
      tag: item.tag,
      id: item.id,
      title: item.title,
      old_sha256: item.old_sha256,
      new_sha256: item.new_sha256,
      old_preview: String(item.old_text).slice(0, 240),
      new_preview: String(item.new_text).slice(0, 240),
    })),
    validation,
    preflight,
    rollback_patchset: rollbackPatchset,
  };
  return {
    ok: validation.ok,
    mode: "word_session",
    audit,
    rollback_patchset: rollbackPatchset,
    content_controls: afterControls,
  };
}

async function wrapSelectionInOpenDocument(tag: string, title?: string): Promise<JsonObject> {
  if (!wordAvailable()) {
    throw new Error("Word host is not available");
  }
  if (!tag) {
    throw new Error("Tag is required");
  }
  const beforeControls = await collectOpenControls();
  await wordApi().run(async (context: any) => {
    const range = context.document.getSelection();
    const control = range.insertContentControl();
    control.tag = tag;
    control.title = title || tag;
    control.appearance = "BoundingBox";
    control.cannotDelete = true;
    control.cannotEdit = false;
    await context.sync();
  });
  const afterControls = await collectOpenControls();
  return {
    ok: afterControls.length === beforeControls.length + 1,
    mode: "word_session",
    operation: "wrap_selection",
    tag,
    title: title || tag,
    before_content_control_count: beforeControls.length,
    after_content_control_count: afterControls.length,
  };
}

async function previewOpenPatchset(): Promise<void> {
  const payload = await previewPatchsetInOpenDocument(parsePatchset());
  log("Open document preview", payload);
}

async function applyOpenPatchset(): Promise<void> {
  if (!window.confirm("Apply PatchSet to the currently open Word document?")) {
    return;
  }
  const result = await applyPatchsetToOpenDocument(parsePatchset());
  log("Applied PatchSet to open document", result.audit);
}

async function executeSessionCommand(command: SessionCommand): Promise<JsonObject> {
  const payload = command.payload || {};
  if (command.type === "list_content_controls") {
    const controls = await collectOpenControls();
    return {ok: true, content_controls: controls, count: controls.length};
  }
  if (command.type === "read_content_control") {
    const tag = String(payload.tag || "");
    if (!tag) {
      throw new Error("read_content_control requires tag");
    }
    const control = (await collectOpenControls()).find((item) => item.tag === tag);
    if (!control) {
      throw new Error(`Open document content control not found: ${tag}`);
    }
    return {
      ok: true,
      tag,
      id: control.id,
      title: control.title,
      text: control.text,
      text_sha256: control.text_sha256,
      validation: {
        ok: true,
        mode: "word_session_read",
        checks: ["content_control_exists", "text_sha256_calculated"],
      },
    };
  }
  if (command.type === "preview_patchset") {
    return await previewPatchsetInOpenDocument(payload.patchset);
  }
  if (command.type === "apply_patchset") {
    return await applyPatchsetToOpenDocument(payload.patchset, payload.rollback_of_command_id);
  }
  if (command.type === "wrap_selection") {
    return await wrapSelectionInOpenDocument(String(payload.tag || ""), payload.title ? String(payload.title) : undefined);
  }
  throw new Error(`Unsupported session command: ${command.type}`);
}

async function postSessionCommandResult(command: SessionCommand, result?: JsonObject, error?: unknown): Promise<void> {
  const snapshot = wordAvailable()
    ? {
        status: error ? "error" : "active",
        last_error: error instanceof Error ? error.message : error ? String(error) : null,
        document: officeDocumentInfo(),
        content_controls: await collectOpenControls().catch(() => openControls),
      }
    : undefined;
  await bridgeFetch("/office/session/result", {
    session_id: sessionId,
    command_id: command.command_id,
    result,
    error: error ? (error instanceof Error ? {message: error.message, name: error.name} : {message: String(error)}) : undefined,
    snapshot,
  });
}

async function pollSessionCommands(): Promise<void> {
  if (!sessionId || !wordAvailable() || pollingCommands) {
    return;
  }
  pollingCommands = true;
  try {
    const payload = await bridgeFetch("/office/session/poll", {session_id: sessionId, limit: 5});
    const commands = (payload.commands || []) as SessionCommand[];
    for (const command of commands) {
      try {
        const result = await executeSessionCommand(command);
        await postSessionCommandResult(command, result);
        log(`Session command completed: ${command.type}`, {command_id: command.command_id, result});
      } catch (error) {
        await postSessionCommandResult(command, undefined, error);
        const message = error instanceof Error ? error.message : String(error);
        log(`Session command failed: ${command.type}: ${message}`, {command_id: command.command_id});
      }
    }
  } finally {
    pollingCommands = false;
  }
}

function bind(id: string, action: () => Promise<void>): void {
  el<HTMLButtonElement>(id).onclick = () => void runAction(el<HTMLButtonElement>(id).textContent || id, action);
}

function init(): void {
  setValue("bridgeUrl", localStorage.getItem("wordAiBridgeUrl") || value("bridgeUrl"));
  setValue("bridgeToken", localStorage.getItem("wordAiBridgeToken") || value("bridgeToken"));
  bind("connectBridge", connectBridge);
  bind("registerSession", registerSession);
  bind("wrapSelection", wrapSelection);
  bind("listOpenControls", listOpenControls);
  bind("loadFileControls", loadFileControls);
  bind("copyOpenTag", copyOpenTag);
  bind("buildPatchset", buildPatchset);
  bind("assessPatchset", assessPatchset);
  bind("dryRunPatchset", dryRunPatchset);
  bind("applyPatchset", applyPatchset);
  bind("previewOpenPatchset", previewOpenPatchset);
  bind("applyOpenPatchset", applyOpenPatchset);
  el<HTMLButtonElement>("clearLog").onclick = () => {
    el<HTMLTextAreaElement>("log").value = "";
  };
  if (!wordAvailable()) {
    setStatus("bridgeStatus", "Word host unavailable; file bridge still works");
  }
}

if (officeApi() && typeof officeApi().onReady === "function") {
  officeApi().onReady(() => init());
} else {
  init();
}
