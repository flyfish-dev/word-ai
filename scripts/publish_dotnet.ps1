param(
  [string]$RuntimeIdentifier = "",
  [switch]$All
)

$ErrorActionPreference = "Stop"
$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$Python = if ($env:PYTHON) { $env:PYTHON } else { "python" }

if ($All) {
  & $Python (Join-Path $Root "scripts\publish_native_backends.py") --all --clean
  exit $LASTEXITCODE
}

if ($RuntimeIdentifier) {
  & $Python (Join-Path $Root "scripts\publish_native_backends.py") $RuntimeIdentifier --clean
  exit $LASTEXITCODE
}

& $Python (Join-Path $Root "scripts\publish_native_backends.py") --clean
exit $LASTEXITCODE
