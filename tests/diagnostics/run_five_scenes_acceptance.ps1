# ==============================================================================
# ExcelMind AI: Five Scenes Live Acceptance Verification Suite
# ==============================================================================

param(
    [string]$RunId = "",
    [switch]$CompileOnly
)

$ErrorActionPreference = "Stop"

$ScriptDir = $PSScriptRoot
$ProjectRoot = Split-Path -Parent (Split-Path -Parent $ScriptDir)

if ([string]::IsNullOrWhiteSpace($RunId)) {
    $RunId = "five_scenes_" + (Get-Date -Format "yyyyMMdd_HHmmss")
}
if ($RunId -notmatch '^[A-Za-z0-9._-]+$' -or $RunId -eq '.' -or $RunId -eq '..') {
    throw "RunId contains invalid path characters: $RunId"
}
$ArtifactsDir = Join-Path $ProjectRoot ".artifacts\tests\$RunId"

if (-not (Test-Path $ArtifactsDir)) {
    New-Item -ItemType Directory -Force -Path $ArtifactsDir | Out-Null
}

$ScreenshotsDir = Join-Path $ArtifactsDir "screenshots"
if (-not (Test-Path $ScreenshotsDir)) {
    New-Item -ItemType Directory -Force -Path $ScreenshotsDir | Out-Null
}

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "      ExcelMind AI: Five Scenes Live Acceptance Suite                         " -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "Project Root   : $ProjectRoot"
Write-Host "Artifacts Dir  : $ArtifactsDir"

# 1. Compile C# COM Driver
Write-Host "`n[1/3] Compiling Native Excel COM Driver (FiveScenesExcelDriver.exe)..." -ForegroundColor Yellow
$CscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$WpfLib = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF"
$DriverSource = Join-Path $ProjectRoot "tests\tools\FiveScenesExcelDriver.cs"
$DriverExe = Join-Path $ArtifactsDir "FiveScenesExcelDriver.exe"
$LeeExcelDll = Join-Path $ProjectRoot "bin\LeeExcel.dll"

$CscArgs = @(
    "/nologo",
    "/target:exe",
    "/out:$DriverExe",
    "/lib:$WpfLib",
    "/r:System.dll",
    "/r:System.Core.dll",
    "/r:System.Drawing.dll",
    "/r:System.Windows.Forms.dll",
    "/r:Microsoft.CSharp.dll",
    "/r:WindowsBase.dll",
    "/r:$LeeExcelDll",
    $DriverSource
)

& $CscPath @CscArgs
if ($LASTEXITCODE -ne 0 -or !(Test-Path $DriverExe)) {
    Write-Host "[ERROR] Failed to compile FiveScenesExcelDriver.exe, code: $LASTEXITCODE" -ForegroundColor Red
    exit 1
}
Write-Host "      Compilation successful: $DriverExe" -ForegroundColor Green

# Copy dependencies to ArtifactsDir
Copy-Item $LeeExcelDll (Join-Path $ArtifactsDir "LeeExcel.dll") -Force
$ExcelDnaDll = Join-Path $ProjectRoot "packages\ExcelDna.Integration.1.9.0\lib\net462\ExcelDna.Integration.dll"
if (Test-Path $ExcelDnaDll) {
    Copy-Item $ExcelDnaDll (Join-Path $ArtifactsDir "ExcelDna.Integration.dll") -Force
}

if ($CompileOnly) {
    Write-Host "[PASS] CompileOnly completed; Excel was not started." -ForegroundColor Green
    exit 0
}

# 2. Run Scene 1 and Scene 2 (Native Excel COM)
Write-Host "`n[2/3] Launching real Excel instance and executing Scene 1 & Scene 2..." -ForegroundColor Yellow
& $DriverExe $ArtifactsDir
$driverExitCode = $LASTEXITCODE
if ($driverExitCode -ne 0) {
    Write-Host "[ERROR] Live Excel COM test failed, code: $driverExitCode" -ForegroundColor Red
    exit $driverExitCode
}
Write-Host "      Scene 1 and Scene 2 verified successfully!" -ForegroundColor Green

# 3. Run Scene 3, 4, 5 and Generate Summary Report
Write-Host "`n[3/3] Executing Scene 3, 4, 5 and generating summary report..." -ForegroundColor Yellow
$nodeScript = Join-Path $ProjectRoot "tests\diagnostics\run_scenes_3_4_5.cjs"
$env:LEE_EXCEL_TEST_RUN_ID = $RunId
& node $nodeScript
$nodeExitCode = $LASTEXITCODE
Remove-Item Env:\LEE_EXCEL_TEST_RUN_ID -ErrorAction SilentlyContinue
if ($nodeExitCode -ne 0) {
    Write-Host "[ERROR] Node test execution failed, code: $nodeExitCode" -ForegroundColor Red
    exit $nodeExitCode
}

Write-Host "`n================================================================================" -ForegroundColor Green
Write-Host "      Acceptance run completed. Dynamic results are in the generated reports." -ForegroundColor Green
Write-Host "================================================================================" -ForegroundColor Green
Write-Host "Reports archived:"
Write-Host "  - Scene 1 Report: $ArtifactsDir\scene1_track_a_overview.md" -ForegroundColor Magenta
Write-Host "  - Scene 2 Report: $ArtifactsDir\scene2_track_b_readback.md" -ForegroundColor Magenta
Write-Host "  - Scene 3 Report: $ArtifactsDir\scene3_evidence_ledger_compaction.md" -ForegroundColor Magenta
Write-Host "  - Scene 4 Report: $ArtifactsDir\scene4_macro_save_eligibility_ui.md" -ForegroundColor Magenta
Write-Host "  - Scene 5 Report: $ArtifactsDir\scene5_deployment_compatibility_cache.md" -ForegroundColor Magenta
Write-Host "  - Summary Report: $ArtifactsDir\acceptance_summary_report.md" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Green

exit 0
