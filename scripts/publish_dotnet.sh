#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RID="${1:-}"

PYTHON="${PYTHON:-python3}"

if [ "$RID" = "--all" ] || [ "$RID" = "all" ]; then
  exec "$PYTHON" "$ROOT/scripts/publish_native_backends.py" --all --clean
fi

if [ -n "$RID" ]; then
  exec "$PYTHON" "$ROOT/scripts/publish_native_backends.py" "$RID" --clean
fi

exec "$PYTHON" "$ROOT/scripts/publish_native_backends.py" --clean
