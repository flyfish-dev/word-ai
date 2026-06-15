type JsonObject = Record<string, any>;

type OpenControl = {
  id: number | string;
  tag: string;
  title: string;
  text: string;
  text_sha256: string;
};

let openControls: OpenControl[] = [];

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

async function connectBridge(): Promise<void> {
  const payload = await bridgeFetch("/office/capabilities", undefined, "GET");
  localStorage.setItem("wordAiBridgeUrl", value("bridgeUrl"));
  localStorage.setItem("wordAiBridgeToken", value("bridgeToken"));
  setStatus("bridgeStatus", `Connected: ${payload.name}`, "ok");
  log("Bridge capabilities", payload);
}

async function wrapSelection(): Promise<void> {
  if (!wordAvailable()) {
    throw new Error("Word host is not available");
  }
  const tag = value("tag");
  const title = value("title") || tag;
  if (!tag) {
    throw new Error("Tag is required");
  }
  await wordApi().run(async (context: any) => {
    const range = context.document.getSelection();
    const control = range.insertContentControl();
    control.tag = tag;
    control.title = title;
    control.appearance = "BoundingBox";
    control.cannotDelete = true;
    control.cannotEdit = false;
    await context.sync();
  });
  log(`Wrapped selection as ${tag}`);
}

async function listOpenControls(): Promise<void> {
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

async function previewOpenPatchset(): Promise<void> {
  const patchset = parsePatchset();
  const previews: JsonObject[] = [];
  if (!wordAvailable()) {
    throw new Error("Word host is not available");
  }
  await wordApi().run(async (context: any) => {
    for (const op of patchset.operations || []) {
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
  log("Open document preview", previews);
}

async function applyOpenPatchset(): Promise<void> {
  if (!window.confirm("Apply PatchSet to the currently open Word document?")) {
    return;
  }
  const patchset = parsePatchset();
  if (!wordAvailable()) {
    throw new Error("Word host is not available");
  }
  await wordApi().run(async (context: any) => {
    for (const op of patchset.operations || []) {
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
      control.insertText(next, wordApi().InsertLocation.replace);
    }
    await context.sync();
  });
  log("Applied PatchSet to open document");
  await listOpenControls();
}

function bind(id: string, action: () => Promise<void>): void {
  el<HTMLButtonElement>(id).onclick = () => void runAction(el<HTMLButtonElement>(id).textContent || id, action);
}

function init(): void {
  setValue("bridgeUrl", localStorage.getItem("wordAiBridgeUrl") || value("bridgeUrl"));
  setValue("bridgeToken", localStorage.getItem("wordAiBridgeToken") || value("bridgeToken"));
  bind("connectBridge", connectBridge);
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
