#!/usr/bin/env node
"use strict";

const childProcess = require("node:child_process");
const crypto = require("node:crypto");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const SUPPORTED_NATIVE_RIDS = new Set([
  "osx-arm64",
  "osx-x64",
  "linux-x64",
  "linux-arm64",
  "linux-musl-x64",
  "linux-musl-arm64",
  "win-x64",
  "win-arm64"
]);

function packageRoot() {
  return fs.realpathSync(path.resolve(__dirname, ".."));
}

function readPackage(root) {
  return JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
}

function cacheRoot() {
  if (process.env.WORD_AI_NPM_CACHE) {
    return path.resolve(process.env.WORD_AI_NPM_CACHE);
  }
  if (process.platform === "win32") {
    return path.join(process.env.LOCALAPPDATA || os.homedir(), "word-ai", "npm");
  }
  return path.join(process.env.XDG_CACHE_HOME || path.join(os.homedir(), ".cache"), "word-ai", "npm");
}

function venvPython(venvDir) {
  return process.platform === "win32"
    ? path.join(venvDir, "Scripts", "python.exe")
    : path.join(venvDir, "bin", "python");
}

function sleepSync(ms) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

function acquireLock(lockDir) {
  const start = Date.now();
  fs.mkdirSync(path.dirname(lockDir), {recursive: true});
  while (true) {
    try {
      fs.mkdirSync(lockDir, {recursive: false});
      fs.writeFileSync(path.join(lockDir, "pid"), `${process.pid}\n`, "utf8");
      return;
    } catch (error) {
      if (error && error.code !== "EEXIST") {
        throw error;
      }
      try {
        const stat = fs.statSync(lockDir);
        if (Date.now() - stat.mtimeMs > 10 * 60 * 1000) {
          fs.rmSync(lockDir, {recursive: true, force: true});
          continue;
        }
      } catch (_) {
        continue;
      }
      if (Date.now() - start > 5 * 60 * 1000) {
        throw new Error(`Timed out waiting for Word AI npm bootstrap lock: ${lockDir}`);
      }
      sleepSync(250);
    }
  }
}

function releaseLock(lockDir) {
  fs.rmSync(lockDir, {recursive: true, force: true});
}

function runChecked(command, args, options = {}) {
  const proc = childProcess.spawnSync(command, args, {
    stdio: options.stdio || ["ignore", 2, 2],
    env: options.env || process.env,
    cwd: options.cwd || process.cwd(),
    shell: false
  });
  if (proc.error) {
    throw proc.error;
  }
  if (proc.status !== 0) {
    throw new Error(`${command} ${args.join(" ")} exited with status ${proc.status}`);
  }
  return proc;
}

function commandExists(command) {
  const probe = process.platform === "win32" ? ["where", command] : ["sh", "-c", `command -v ${command}`];
  const proc = childProcess.spawnSync(probe[0], probe.slice(1), {
    stdio: "ignore",
    shell: false
  });
  return !proc.error && proc.status === 0;
}

function downloadFile(url, outputPath) {
  fs.mkdirSync(path.dirname(outputPath), {recursive: true});
  if (commandExists("curl")) {
    runChecked("curl", ["-fsSL", "--retry", "3", "--retry-delay", "2", "-o", outputPath, url]);
    return;
  }
  if (process.platform === "win32" && commandExists("powershell")) {
    runChecked("powershell", [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-Command",
      `$ErrorActionPreference='Stop'; Invoke-WebRequest -Uri '${url}' -OutFile '${outputPath.replace(/'/g, "''")}'`
    ]);
    return;
  }
  throw new Error("Neither curl nor PowerShell is available for downloading the Word AI native backend.");
}

function sha256File(filePath) {
  const hash = crypto.createHash("sha256");
  hash.update(fs.readFileSync(filePath));
  return hash.digest("hex");
}

function detectRid() {
  const arch = process.arch === "arm64" ? "arm64" : process.arch === "x64" ? "x64" : null;
  if (!arch) {
    return null;
  }
  if (process.platform === "darwin") {
    return `osx-${arch}`;
  }
  if (process.platform === "win32") {
    return `win-${arch}`;
  }
  if (process.platform === "linux") {
    const glibc = process.report && process.report.getReport
      ? process.report.getReport().header.glibcVersionRuntime
      : null;
    return `${glibc ? "linux" : "linux-musl"}-${arch}`;
  }
  return null;
}

function nativeExecutableName(rid) {
  return rid && rid.startsWith("win-") ? "WordAi.OpenXml.exe" : "WordAi.OpenXml";
}

function findBundledNativeRoot(root, rid) {
  for (const nativeRoot of [path.join(root, "native"), path.join(root, "dist", "native")]) {
    const exe = path.join(nativeRoot, rid, nativeExecutableName(rid));
    if (fs.existsSync(exe)) {
      return nativeRoot;
    }
  }
  return null;
}

function extractArchive(archivePath, destination) {
  fs.mkdirSync(destination, {recursive: true});
  if (archivePath.endsWith(".tar.gz")) {
    runChecked("tar", ["-xzf", archivePath, "-C", destination]);
    return;
  }
  if (archivePath.endsWith(".zip")) {
    if (process.platform === "win32" && commandExists("powershell")) {
      runChecked("powershell", [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        `$ErrorActionPreference='Stop'; Expand-Archive -LiteralPath '${archivePath.replace(/'/g, "''")}' -DestinationPath '${destination.replace(/'/g, "''")}' -Force`
      ]);
      return;
    }
    if (commandExists("unzip")) {
      runChecked("unzip", ["-q", "-o", archivePath, "-d", destination]);
      return;
    }
  }
  throw new Error(`No extractor is available for ${archivePath}`);
}

function ensureNativeBackend(root, version) {
  if (process.env.WORD_AI_DOTNET_EXE || process.env.WORD_AI_DOTNET_NATIVE_DIR || process.env.WORD_AI_SKIP_NATIVE_DOWNLOAD === "1") {
    return null;
  }
  const rid = process.env.WORD_AI_DOTNET_RID || detectRid();
  if (!rid) {
    return null;
  }
  if (!SUPPORTED_NATIVE_RIDS.has(rid)) {
    console.error(`Warning: unsupported Word AI native RID '${rid}'; skipping native download.`);
    return null;
  }
  const bundled = findBundledNativeRoot(root, rid);
  if (bundled) {
    return bundled;
  }

  const nativeRoot = path.join(cacheRoot(), version, "native");
  const exe = path.join(nativeRoot, rid, nativeExecutableName(rid));
  if (fs.existsSync(exe)) {
    return nativeRoot;
  }

  const lockDir = path.join(cacheRoot(), version, ".native.lock");
  acquireLock(lockDir);
  try {
    if (fs.existsSync(exe)) {
      return nativeRoot;
    }
    const extension = rid.startsWith("win-") ? "zip" : "tar.gz";
    const assetName = `word-ai-openxml-${version}-${rid}.${extension}`;
    const baseUrl = `https://github.com/flyfish-dev/word-ai/releases/download/v${version}`;
    const downloadDir = path.join(cacheRoot(), version, "downloads");
    const archivePath = path.join(downloadDir, assetName);
    const checksumsPath = path.join(downloadDir, `word-ai-openxml-${version}-checksums.sha256`);

    console.error(`Downloading Word AI native backend ${rid}...`);
    downloadFile(`${baseUrl}/word-ai-openxml-${version}-checksums.sha256`, checksumsPath);
    downloadFile(`${baseUrl}/${assetName}`, archivePath);
    const checksums = fs.readFileSync(checksumsPath, "utf8");
    const expectedLine = checksums.split(/\r?\n/).find((line) => line.trim().endsWith(`  ${assetName}`));
    if (!expectedLine) {
      throw new Error(`No checksum entry found for ${assetName}`);
    }
    const expected = expectedLine.trim().split(/\s+/)[0];
    const actual = sha256File(archivePath);
    if (actual !== expected) {
      throw new Error(`Checksum mismatch for ${assetName}: expected ${expected}, got ${actual}`);
    }

    fs.rmSync(path.join(nativeRoot, rid), {recursive: true, force: true});
    extractArchive(archivePath, nativeRoot);
    if (!fs.existsSync(exe)) {
      throw new Error(`Downloaded native backend did not contain ${rid}/${nativeExecutableName(rid)}`);
    }
    if (!rid.startsWith("win-")) {
      fs.chmodSync(exe, 0o755);
    }
    return nativeRoot;
  } catch (error) {
    console.error(`Warning: could not install Word AI native backend for ${rid}: ${error.message}`);
    return null;
  } finally {
    releaseLock(lockDir);
  }
}

function pythonCandidates() {
  const candidates = [];
  if (process.env.WORD_AI_PYTHON) {
    candidates.push(process.env.WORD_AI_PYTHON);
  }
  candidates.push("python3", "python");
  return [...new Set(candidates)];
}

function pythonVersion(command) {
  const proc = childProcess.spawnSync(
    command,
    ["-c", "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')"],
    {encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], shell: false}
  );
  if (proc.error || proc.status !== 0) {
    return null;
  }
  const text = proc.stdout.trim();
  const parts = text.split(".").map((part) => Number.parseInt(part, 10));
  if (parts.length < 2 || Number.isNaN(parts[0]) || Number.isNaN(parts[1])) {
    return null;
  }
  return {text, major: parts[0], minor: parts[1]};
}

function findPython() {
  for (const candidate of pythonCandidates()) {
    const version = pythonVersion(candidate);
    if (version && (version.major > 3 || (version.major === 3 && version.minor >= 10))) {
      return candidate;
    }
  }
  throw new Error("Word AI requires Python 3.10 or newer. Set WORD_AI_PYTHON=/path/to/python if needed.");
}

function ensureVenv(root, version) {
  const venvDir = path.join(cacheRoot(), version, "venv");
  const lockDir = path.join(cacheRoot(), version, ".bootstrap.lock");
  const python = venvPython(venvDir);
  const marker = path.join(venvDir, ".word-ai-npm-version");
  const requirements = path.join(root, "requirements.txt");
  if (
    fs.existsSync(python) &&
    fs.existsSync(marker) &&
    fs.readFileSync(marker, "utf8").trim() === version
  ) {
    return python;
  }

  fs.mkdirSync(path.dirname(venvDir), {recursive: true});
  acquireLock(lockDir);
  try {
    if (
      fs.existsSync(python) &&
      fs.existsSync(marker) &&
      fs.readFileSync(marker, "utf8").trim() === version
    ) {
      return python;
    }
    fs.rmSync(venvDir, {recursive: true, force: true});
    const systemPython = findPython();
    console.error(`Bootstrapping Word AI Python environment with ${systemPython}...`);
    runChecked(systemPython, ["-m", "venv", venvDir]);
    runChecked(python, ["-m", "pip", "install", "--upgrade", "pip"]);
    runChecked(python, ["-m", "pip", "install", "-r", requirements]);
    fs.writeFileSync(marker, `${version}\n`, "utf8");
  } finally {
    releaseLock(lockDir);
  }
  return python;
}

function main(moduleName) {
  const root = packageRoot();
  const pkg = readPackage(root);
  const nativeRoot = ensureNativeBackend(root, pkg.version);
  const python = ensureVenv(root, pkg.version);
  const env = {...process.env};
  env.PYTHONPATH = root + (env.PYTHONPATH ? path.delimiter + env.PYTHONPATH : "");
  if (nativeRoot && !env.WORD_AI_DOTNET_NATIVE_DIR) {
    env.WORD_AI_DOTNET_NATIVE_DIR = nativeRoot;
  }
  const child = childProcess.spawn(python, ["-m", moduleName, ...process.argv.slice(2)], {
    stdio: "inherit",
    env
  });
  child.on("error", (error) => {
    console.error(error.message);
    process.exit(1);
  });
  child.on("exit", (code, signal) => {
    if (signal) {
      process.kill(process.pid, signal);
      return;
    }
    process.exit(code === null ? 1 : code);
  });
}

module.exports = {main};
