#!/usr/bin/env node
"use strict";

const childProcess = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

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
    stdio: options.stdio || "inherit",
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
  const python = ensureVenv(root, pkg.version);
  const env = {...process.env};
  env.PYTHONPATH = root + (env.PYTHONPATH ? path.delimiter + env.PYTHONPATH : "");
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
