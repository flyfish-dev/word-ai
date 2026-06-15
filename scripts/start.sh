#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
USE_HTTP=0
BRIDGE_ONLY=0
TASKPANE_ONLY=0

while [ "$#" -gt 0 ]; do
  case "$1" in
    --http)
      USE_HTTP=1
      ;;
    --bridge-only)
      BRIDGE_ONLY=1
      ;;
    --taskpane-only)
      TASKPANE_ONLY=1
      ;;
    -h|--help)
      cat <<EOF
Usage: scripts/start.sh [--http] [--bridge-only] [--taskpane-only]

Starts the Word AI local HTTP bridge and Office.js taskpane.
Default taskpane mode is HTTPS for Word sideloading.
Use --http for browser-only debugging.
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

VENV_PY="${WORD_AI_PYTHON:-$ROOT/.venv/bin/python}"
if [ ! -x "$VENV_PY" ]; then
  echo "Missing $VENV_PY. Run: bash scripts/install.sh" >&2
  exit 1
fi

if [ ! -d "$ROOT/office-addin/node_modules" ] && [ "$BRIDGE_ONLY" -eq 0 ]; then
  echo "Missing office-addin/node_modules. Run: bash scripts/install.sh" >&2
  exit 1
fi

BRIDGE_HOST="${WORD_AI_BRIDGE_HOST:-127.0.0.1}"
BRIDGE_PORT="${WORD_AI_BRIDGE_PORT:-8765}"
TASKPANE_HOST="${WORD_AI_TASKPANE_HOST:-localhost}"
TASKPANE_PORT="${WORD_AI_TASKPANE_PORT:-3000}"
RUN_DIR="$ROOT/.wordai/run"
TOKEN_PATH="$ROOT/.wordai/bridge.token"
mkdir -p "$RUN_DIR" "$ROOT/.wordai"

ALLOWED_ROOT_ARGS=()
for candidate in "$HOME/Downloads" "$HOME/Documents" "$HOME/Desktop"; do
  if [ -d "$candidate" ]; then
    ALLOWED_ROOT_ARGS+=(--allow-root "$candidate")
  fi
done

TOKEN="${WORD_AI_TOKEN:-}"
if [ -z "$TOKEN" ] && [ -f "$TOKEN_PATH" ]; then
  TOKEN="$(tr -d '\r\n' < "$TOKEN_PATH")"
fi
if [ -z "$TOKEN" ]; then
  TOKEN="$("$VENV_PY" -c 'import secrets; print(secrets.token_urlsafe(32))')"
  printf '%s\n' "$TOKEN" > "$TOKEN_PATH"
  chmod 600 "$TOKEN_PATH" 2>/dev/null || true
fi

BRIDGE_PID=""
TASKPANE_PID=""

cleanup() {
  if [ -n "$TASKPANE_PID" ] && kill -0 "$TASKPANE_PID" 2>/dev/null; then
    kill "$TASKPANE_PID" 2>/dev/null || true
  fi
  if [ -n "$BRIDGE_PID" ] && kill -0 "$BRIDGE_PID" 2>/dev/null; then
    kill "$BRIDGE_PID" 2>/dev/null || true
  fi
  if [ -n "$TASKPANE_PID" ]; then
    rm -f "$RUN_DIR/taskpane.pid"
  fi
  if [ -n "$BRIDGE_PID" ]; then
    rm -f "$RUN_DIR/bridge.pid"
  fi
}
trap cleanup INT TERM EXIT

if [ "$TASKPANE_ONLY" -eq 0 ]; then
  BRIDGE_LOG="$RUN_DIR/bridge.log"
  echo "Starting Word AI bridge on http://$BRIDGE_HOST:$BRIDGE_PORT ..."
  "$VENV_PY" -m word_ai_mcp.server_http \
    --root "$ROOT" \
    --host "$BRIDGE_HOST" \
    --port "$BRIDGE_PORT" \
    --token "$TOKEN" \
    "${ALLOWED_ROOT_ARGS[@]}" > "$BRIDGE_LOG" 2>&1 &
  BRIDGE_PID="$!"
  printf '%s\n' "$BRIDGE_PID" > "$RUN_DIR/bridge.pid"

  i=0
  until curl -fsS "http://$BRIDGE_HOST:$BRIDGE_PORT/health" >/dev/null 2>&1; do
    i=$((i + 1))
    if [ "$i" -gt 60 ]; then
      echo "Bridge did not become healthy. Log:" >&2
      tail -n 80 "$BRIDGE_LOG" >&2 || true
      exit 1
    fi
    sleep 0.25
  done
fi

if [ "$BRIDGE_ONLY" -eq 0 ]; then
  TASKPANE_LOG="$RUN_DIR/taskpane.log"
  TASKPANE_CMD="dev"
  TASKPANE_SCHEME="https"
  if [ "$USE_HTTP" -eq 1 ]; then
    TASKPANE_CMD="dev:http"
    TASKPANE_SCHEME="http"
  fi
  echo "Starting Office.js taskpane on $TASKPANE_SCHEME://$TASKPANE_HOST:$TASKPANE_PORT ..."
  (
    cd "$ROOT/office-addin"
    PORT="$TASKPANE_PORT" \
    HOST="$TASKPANE_HOST" \
    WORD_AI_BRIDGE_URL="http://$BRIDGE_HOST:$BRIDGE_PORT" \
    npm run "$TASKPANE_CMD"
  ) > "$TASKPANE_LOG" 2>&1 &
  TASKPANE_PID="$!"
  printf '%s\n' "$TASKPANE_PID" > "$RUN_DIR/taskpane.pid"
fi

echo
echo "Word AI is running."
echo "Bridge: http://$BRIDGE_HOST:$BRIDGE_PORT"
if [ "$BRIDGE_ONLY" -eq 0 ]; then
  echo "Taskpane: $TASKPANE_SCHEME://$TASKPANE_HOST:$TASKPANE_PORT/taskpane.html"
fi
echo "Bridge token: $TOKEN"
echo "Manifest: $ROOT/office-addin/manifest.xml"
if [ "${#ALLOWED_ROOT_ARGS[@]}" -gt 0 ]; then
  echo "Additional allowed roots: ${ALLOWED_ROOT_ARGS[*]}"
fi
echo "Logs: $RUN_DIR"
echo "Stop with Ctrl-C or: bash scripts/stop.sh"
echo

while true; do
  if [ -n "$BRIDGE_PID" ] && ! kill -0 "$BRIDGE_PID" 2>/dev/null; then
    echo "Bridge process exited. Log:" >&2
    tail -n 80 "$RUN_DIR/bridge.log" >&2 || true
    exit 1
  fi
  if [ -n "$TASKPANE_PID" ] && ! kill -0 "$TASKPANE_PID" 2>/dev/null; then
    echo "Taskpane process exited. Log:" >&2
    tail -n 80 "$RUN_DIR/taskpane.log" >&2 || true
    exit 1
  fi
  sleep 2
done
