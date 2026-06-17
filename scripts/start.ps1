param(
  [switch]$Http,
  [switch]$BridgeOnly,
  [switch]$TaskpaneOnly
)

$ErrorActionPreference = "Stop"
$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$VenvPython = if ($env:WORD_AI_PYTHON) { $env:WORD_AI_PYTHON } else { Join-Path $Root ".venv\Scripts\python.exe" }

if (-not (Test-Path $VenvPython)) {
  throw "Missing $VenvPython. Run scripts\install.ps1 first."
}

$BridgeHost = if ($env:WORD_AI_BRIDGE_HOST) { $env:WORD_AI_BRIDGE_HOST } else { "127.0.0.1" }
$BridgePort = if ($env:WORD_AI_BRIDGE_PORT) { $env:WORD_AI_BRIDGE_PORT } else { "8765" }
$TaskpaneHost = if ($env:WORD_AI_TASKPANE_HOST) { $env:WORD_AI_TASKPANE_HOST } else { "localhost" }
$TaskpanePort = if ($env:WORD_AI_TASKPANE_PORT) { $env:WORD_AI_TASKPANE_PORT } else { "3100" }
$RunDir = Join-Path $Root ".wordai\run"
$TokenPath = Join-Path $Root ".wordai\bridge.token"
New-Item -ItemType Directory -Force -Path $RunDir | Out-Null

$AllowedRootArgs = @()
foreach ($Candidate in @((Join-Path $HOME "Downloads"), (Join-Path $HOME "Documents"), (Join-Path $HOME "Desktop"))) {
  if (Test-Path $Candidate) {
    $AllowedRootArgs += @("--allow-root", (Resolve-Path $Candidate).Path)
  }
}

$Token = $env:WORD_AI_TOKEN
if (-not $Token -and (Test-Path $TokenPath)) {
  $Token = (Get-Content $TokenPath -Raw).Trim()
}
if (-not $Token) {
  $Token = & $VenvPython -c "import secrets; print(secrets.token_urlsafe(32))"
  Set-Content -Path $TokenPath -Value $Token
}

$Bridge = $null
$Taskpane = $null

try {
  if (-not $TaskpaneOnly) {
    Write-Host "Starting Word AI bridge on http://${BridgeHost}:${BridgePort} ..."
    $BridgeLog = Join-Path $RunDir "bridge.log"
    $BridgeErr = Join-Path $RunDir "bridge.err.log"
    $BridgeArgs = @("-m", "word_ai_mcp.server_http", "--root", $Root, "--host", $BridgeHost, "--port", $BridgePort, "--token", $Token)
    $BridgeArgs += $AllowedRootArgs
    $Bridge = Start-Process -FilePath $VenvPython -ArgumentList $BridgeArgs -WorkingDirectory $Root -PassThru -NoNewWindow -RedirectStandardOutput $BridgeLog -RedirectStandardError $BridgeErr
    Set-Content -Path (Join-Path $RunDir "bridge.pid") -Value $Bridge.Id
  }

  if (-not $BridgeOnly) {
    if (-not (Test-Path (Join-Path $Root "office-addin\node_modules"))) {
      throw "Missing office-addin\node_modules. Run scripts\install.ps1 first."
    }
    $TaskpaneLog = Join-Path $RunDir "taskpane.log"
    $TaskpaneErr = Join-Path $RunDir "taskpane.err.log"
    $TaskpaneCmd = if ($Http) { "dev:http" } else { "dev" }
    $Scheme = if ($Http) { "http" } else { "https" }
    Write-Host "Starting Office.js taskpane on ${Scheme}://${TaskpaneHost}:${TaskpanePort} ..."
    $env:PORT = $TaskpanePort
    $env:HOST = $TaskpaneHost
    $env:WORD_AI_BRIDGE_URL = "http://${BridgeHost}:${BridgePort}"
    $Npm = if ($IsWindows) { "npm.cmd" } else { "npm" }
    $Taskpane = Start-Process -FilePath $Npm -ArgumentList @("run", $TaskpaneCmd) -WorkingDirectory (Join-Path $Root "office-addin") -PassThru -NoNewWindow -RedirectStandardOutput $TaskpaneLog -RedirectStandardError $TaskpaneErr
    Set-Content -Path (Join-Path $RunDir "taskpane.pid") -Value $Taskpane.Id
  }

  Write-Host ""
  Write-Host "Word AI is running."
  Write-Host "Bridge: http://${BridgeHost}:${BridgePort}"
  if (-not $BridgeOnly) {
    Write-Host "Taskpane: ${Scheme}://${TaskpaneHost}:${TaskpanePort}/taskpane.html"
  }
  Write-Host "Bridge token: $Token"
  Write-Host "Manifest: $(Join-Path $Root 'office-addin\manifest.xml')"
  if ($AllowedRootArgs.Count -gt 0) {
    Write-Host "Additional allowed roots: $($AllowedRootArgs -join ' ')"
  }
  Write-Host "Logs: $RunDir"
  Write-Host "Stop with Ctrl-C or scripts\stop.ps1"

  while ($true) {
    Start-Sleep -Seconds 2
    if ($Bridge) { $Bridge.Refresh(); if ($Bridge.HasExited) { throw "Bridge process exited." } }
    if ($Taskpane) { $Taskpane.Refresh(); if ($Taskpane.HasExited) { throw "Taskpane process exited." } }
  }
}
finally {
  if ($Taskpane -and -not $Taskpane.HasExited) { Stop-Process -Id $Taskpane.Id -ErrorAction SilentlyContinue }
  if ($Bridge -and -not $Bridge.HasExited) { Stop-Process -Id $Bridge.Id -ErrorAction SilentlyContinue }
  Remove-Item (Join-Path $RunDir "taskpane.pid") -Force -ErrorAction SilentlyContinue
  Remove-Item (Join-Path $RunDir "bridge.pid") -Force -ErrorAction SilentlyContinue
}
