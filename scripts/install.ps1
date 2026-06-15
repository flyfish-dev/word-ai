param(
  [switch]$InstallSkill,
  [switch]$InstallAgentSkills,
  [switch]$NoAgentSkills,
  [switch]$SkipNode,
  [switch]$SkipDotnet
)

$ErrorActionPreference = "Stop"
$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$VenvPython = Join-Path $Root ".venv\Scripts\python.exe"

Set-Location $Root
New-Item -ItemType Directory -Force -Path (Join-Path $Root ".wordai") | Out-Null

if (-not (Test-Path $VenvPython)) {
  Write-Host "Creating Python virtual environment..."
  $Python = if ($env:PYTHON) { $env:PYTHON } else { "python" }
  & $Python -m venv (Join-Path $Root ".venv")
}

Write-Host "Installing Python package and MCP dependencies..."
& $VenvPython -m pip install --upgrade pip
& $VenvPython -m pip install -r (Join-Path $Root "requirements.txt")
& $VenvPython -m pip install -e $Root

if (-not $SkipNode) {
  if (Get-Command npm -ErrorAction SilentlyContinue) {
    Write-Host "Installing and building Office.js taskpane..."
    Push-Location (Join-Path $Root "office-addin")
    & npm install
    & npm run build
    Pop-Location
  } else {
    Write-Warning "npm not found; skipped Office.js taskpane install."
  }
}

if (-not $SkipDotnet) {
  if (Get-Command dotnet -ErrorAction SilentlyContinue) {
    Write-Host "Building .NET Open XML engine..."
    & dotnet build (Join-Path $Root "dotnet\WordAi.OpenXml\WordAi.OpenXml.csproj") -c Release
  } else {
    Write-Warning "dotnet not found; install .NET SDK 8 for the Open XML backend."
  }
}

Write-Host "Writing Codex MCP config snippet..."
& $VenvPython -m word_ai_mcp.quickstart --root $Root codex-config --output (Join-Path $Root ".wordai\codex-config.toml")

if (-not $NoAgentSkills) {
  Write-Host "Installing Word AI agent skills..."
  & $VenvPython -m word_ai_mcp.quickstart --root $Root install-skills --agents auto
}

Write-Host ""
Write-Host "Word AI install complete."
Write-Host "Codex config snippet: $(Join-Path $Root '.wordai\codex-config.toml')"
Write-Host "Agent skill installer: $VenvPython -m word_ai_mcp.quickstart --root `"$Root`" install-skills"
Write-Host "Start local bridge + taskpane: powershell -ExecutionPolicy Bypass -File scripts\start.ps1"
Write-Host "Word manifest: $(Join-Path $Root 'office-addin\manifest.xml')"
