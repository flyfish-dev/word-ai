#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RID="${1:-}"

if [ -z "$RID" ]; then
  OS="$(uname -s)"
  ARCH="$(uname -m)"
  case "$ARCH" in
    arm64|aarch64) ARCH_ID="arm64" ;;
    x86_64|amd64) ARCH_ID="x64" ;;
    *) echo "Unsupported architecture: $ARCH" >&2; exit 2 ;;
  esac
  case "$OS" in
    Darwin) RID="osx-$ARCH_ID" ;;
    Linux) RID="linux-$ARCH_ID" ;;
    *) echo "Unsupported OS: $OS" >&2; exit 2 ;;
  esac
fi

OUT="$ROOT/dist/native/$RID"
dotnet publish "$ROOT/dotnet/WordAi.OpenXml/WordAi.OpenXml.csproj" \
  -c Release \
  -r "$RID" \
  --self-contained true \
  -p:UseAppHost=true \
  -p:PublishSingleFile=true \
  -p:PublishTrimmed=false \
  -p:EnableCompressionInSingleFile=true \
  -o "$OUT"

echo "Published WordAi.OpenXml native backend to: $OUT"
