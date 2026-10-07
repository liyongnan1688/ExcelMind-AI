# ==============================================================================
# ExcelMind AI: Isolated Release ZIP Smoke Verification Runner
# ==============================================================================

param(
    [string]$RunId = ""
)

$ErrorActionPreference = "Stop"

$ScriptDir = $PSScriptRoot
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $ScriptDir)

if ([string]::IsNullOrEmpty($RunId)) {
    $RunId = "smoke_isolated_" + (Get-Date -Format "yyyyMMdd_HHmmss")
}

$ArtifactsDir = Join-Path $ProjectRoot ".artifacts\tests\$RunId"
if (-not (Test-Path $ArtifactsDir)) {
    New-Item -ItemType Directory -Force -Path $ArtifactsDir | Out-Null
}

$ZipPath = Join-Path $ProjectRoot ".artifacts\release\ExcelMindAI-v1.2.0-rc1.zip"
if (-not (Test-Path $ZipPath)) {
    Write-Host "[ERROR] Release ZIP not found: $ZipPath" -ForegroundColor Red
    Write-Host "Please run scripts/package_release.ps1 first!" -ForegroundColor Yellow
    exit 1
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "    ExcelMind AI: Isolated Release Smoke Verification ($RunId)                  " -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "ZIP Path       : $ZipPath"
Write-Host "Artifacts Dir  : $ArtifactsDir"

# 1. Compile standalone smoke test tool
Write-Host "`n[1/3] Compiling standalone VerifyIsolatedReleaseSmoke tool..." -ForegroundColor Yellow
$CscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$WpfLib = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF"
$SourceFile = Join-Path $ProjectRoot "tests\tools\VerifyIsolatedReleaseSmoke.cs"
$OutExe = Join-Path $ArtifactsDir "VerifyIsolatedReleaseSmoke.exe"

$CscArgs = @(
    "/nologo",
    "/target:exe",
    "/out:$OutExe",
    "/lib:$WpfLib",
    "/r:System.dll",
    "/r:System.Core.dll",
    "/r:System.Drawing.dll",
    "/r:System.Windows.Forms.dll",
    "/r:System.IO.Compression.dll",
    "/r:System.IO.Compression.FileSystem.dll",
    "/r:Microsoft.CSharp.dll",
    "/r:UIAutomationClient.dll",
    "/r:UIAutomationTypes.dll",
    "/r:WindowsBase.dll",
    "`"$SourceFile`""
)

$cmd = "& `"$CscPath`" " + ($CscArgs -join " ")
Invoke-Expression $cmd
if ($LASTEXITCODE -ne 0 -or !(Test-Path $OutExe)) {
    Write-Host "[ERROR] Failed to compile VerifyIsolatedReleaseSmoke.exe, code: $LASTEXITCODE" -ForegroundColor Red
    exit 1
}
Write-Host "      Compilation successful: $OutExe" -ForegroundColor Green

# 2. Run isolated smoke check
Write-Host "`n[2/3] Running isolated release smoke check..." -ForegroundColor Yellow
$runnerCmd = "& `"$OutExe`" `"$ArtifactsDir`" `"$ZipPath`""
Invoke-Expression $runnerCmd
$smokeExitCode = $LASTEXITCODE

# 3. Archive report
Write-Host "`n[3/3] Archiving smoke report and artifacts..." -ForegroundColor Yellow
$reportPath = Join-Path $ArtifactsDir "smoke_report.md"
if (Test-Path $reportPath) {
    Write-Host "      Report archived: $reportPath" -ForegroundColor Green
}

if ($smokeExitCode -eq 0) {
    Write-Host "`n[PASS] All isolated smoke checks PASSED!" -ForegroundColor Green
} else {
    Write-Host "`n[FAIL] Smoke checks failed, exit code: $smokeExitCode" -ForegroundColor Red
}

exit $smokeExitCode
