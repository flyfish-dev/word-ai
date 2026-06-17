param(
  [string]$RuntimeIdentifier = ""
)

$ErrorActionPreference = "Stop"
$Root = Resolve-Path (Join-Path $PSScriptRoot "..")

if (-not $RuntimeIdentifier) {
  $Arch = if ([System.Runtime.InteropServices.RuntimeInformation]::OSArchitecture -eq "Arm64") { "arm64" } else { "x64" }
  if ($IsWindows) {
    $RuntimeIdentifier = "win-$Arch"
  } elseif ($IsMacOS) {
    $RuntimeIdentifier = "osx-$Arch"
  } elseif ($IsLinux) {
    $RuntimeIdentifier = "linux-$Arch"
  } else {
    throw "Unsupported OS for automatic RID detection."
  }
}

$Out = Join-Path $Root "dist\native\$RuntimeIdentifier"
dotnet publish (Join-Path $Root "dotnet\WordAi.OpenXml\WordAi.OpenXml.csproj") `
  -c Release `
  -r $RuntimeIdentifier `
  --self-contained true `
  -p:UseAppHost=true `
  -p:PublishSingleFile=true `
  -p:PublishTrimmed=false `
  -p:EnableCompressionInSingleFile=true `
  -o $Out

Write-Host "Published WordAi.OpenXml native backend to: $Out"
