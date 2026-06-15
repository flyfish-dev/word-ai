#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RUN_DIR="$ROOT/.wordai/run"

stop_pid_file() {
  local name="$1"
  local file="$RUN_DIR/$name.pid"
  if [ ! -f "$file" ]; then
    return 0
  fi
  local pid
  pid="$(tr -d '\r\n' < "$file")"
  if [ -n "$pid" ] && kill -0 "$pid" 2>/dev/null; then
    echo "Stopping $name process $pid ..."
    kill "$pid" 2>/dev/null || true
  fi
  rm -f "$file"
}

stop_pid_file taskpane
stop_pid_file bridge

echo "Word AI local processes stopped."
