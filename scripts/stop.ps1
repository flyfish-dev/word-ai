$ErrorActionPreference = "Stop"
$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
$RunDir = Join-Path $Root ".wordai\run"

function Stop-PidFile($Name) {
  $File = Join-Path $RunDir "$Name.pid"
  if (-not (Test-Path $File)) {
    return
  }
  $PidText = (Get-Content $File -Raw).Trim()
  if ($PidText) {
    $Process = Get-Process -Id ([int]$PidText) -ErrorAction SilentlyContinue
    if ($Process) {
      Write-Host "Stopping $Name process $PidText ..."
      Stop-Process -Id $Process.Id -ErrorAction SilentlyContinue
    }
  }
  Remove-Item $File -Force -ErrorAction SilentlyContinue
}

Stop-PidFile "taskpane"
Stop-PidFile "bridge"

Write-Host "Word AI local processes stopped."
