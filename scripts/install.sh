#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
INSTALL_SKILL=0
SKIP_NODE=0
SKIP_DOTNET=0

while [ "$#" -gt 0 ]; do
  case "$1" in
    --install-skill)
      INSTALL_SKILL=1
      ;;
    --skip-node)
      SKIP_NODE=1
      ;;
    --skip-dotnet)
      SKIP_DOTNET=1
      ;;
    -h|--help)
      cat <<EOF
Usage: scripts/install.sh [--install-skill] [--skip-node] [--skip-dotnet]

Installs Python dependencies, builds the Office.js taskpane, builds the .NET
Open XML engine, and writes .wordai/codex-config.toml.
EOF
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      exit 2
      ;;
  esac
  shift
done

cd "$ROOT"
mkdir -p "$ROOT/.wordai"

PYTHON_BIN="${PYTHON:-python3}"
if [ ! -x "$ROOT/.venv/bin/python" ]; then
  echo "Creating Python virtual environment..."
  "$PYTHON_BIN" -m venv "$ROOT/.venv"
fi

VENV_PY="$ROOT/.venv/bin/python"
echo "Installing Python package and MCP dependencies..."
"$VENV_PY" -m pip install --upgrade pip
"$VENV_PY" -m pip install -r "$ROOT/requirements.txt"
"$VENV_PY" -m pip install -e "$ROOT"

if [ "$SKIP_NODE" -eq 0 ]; then
  if command -v npm >/dev/null 2>&1; then
    echo "Installing and building Office.js taskpane..."
    (cd "$ROOT/office-addin" && npm install && npm run build)
  else
    echo "WARN: npm not found; skipped Office.js taskpane install." >&2
  fi
fi

if [ "$SKIP_DOTNET" -eq 0 ]; then
  if command -v dotnet >/dev/null 2>&1; then
    echo "Building .NET Open XML engine..."
    dotnet build "$ROOT/dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj" -c Release
  else
    echo "WARN: dotnet not found; install .NET SDK 8 for the Open XML backend." >&2
  fi
fi

echo "Writing Codex MCP config snippet..."
"$VENV_PY" -m word_ai_mcp.quickstart --root "$ROOT" codex-config --output "$ROOT/.wordai/codex-config.toml"

if [ "$INSTALL_SKILL" -eq 1 ]; then
  CODEX_HOME="${CODEX_HOME:-$HOME/.codex}"
  SKILL_DEST="$CODEX_HOME/skills/word-ai"
  echo "Installing Codex Skill to $SKILL_DEST ..."
  mkdir -p "$SKILL_DEST"
  cp -R "$ROOT/skills/word-ai/." "$SKILL_DEST/"
fi

echo
echo "Word AI install complete."
echo "Codex config snippet: $ROOT/.wordai/codex-config.toml"
echo "Start local bridge + taskpane: bash scripts/start.sh"
echo "Browser-only taskpane debug: bash scripts/start.sh --http"
echo "Word manifest: $ROOT/office-addin/manifest.xml"
