# ==============================================================================
# ExcelMind AI: Real Desktop Excel Acceptance Automation Entry
# ==============================================================================

param(
    [string]$RunId = "",
    [switch]$CompileOnly
)

$ErrorActionPreference = "Stop"
$ScriptDir = $PSScriptRoot
$ProjectRoot = Split-Path -Parent $ScriptDir

if ([string]::IsNullOrWhiteSpace($RunId)) {
    $RunId = "desktop_acceptance_" + (Get-Date -Format "yyyyMMdd_HHmmss")
}
if ($RunId -notmatch '^[A-Za-z0-9._-]+$' -or $RunId -eq '.' -or $RunId -eq '..') {
    throw "RunId contains invalid path characters: $RunId"
}

$ArtifactsDir = Join-Path $ProjectRoot ".artifacts\tests\$RunId"
if (-not (Test-Path $ArtifactsDir)) {
    New-Item -ItemType Directory -Force -Path $ArtifactsDir | Out-Null
}

$LogFile = Join-Path $ArtifactsDir "runner_console.log"
$CscPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$WpfLib = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\WPF"

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "    ExcelMind AI: Real Desktop Excel Acceptance Automation ($RunId)" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# 1. Environment check
Write-Host "[1/4] Checking environment and pre-requisites..." -ForegroundColor Yellow
if (-not (Test-Path $CscPath)) {
    Write-Host "[ERROR] .NET Framework 64-bit compiler not found: $CscPath" -ForegroundColor Red
    exit 1
}

$XllPath = Join-Path $ProjectRoot "bin\LeeExcel64.xll"
if (-not (Test-Path $XllPath)) {
    Write-Host "[WARN] bin\LeeExcel64.xll not found, building addin..." -ForegroundColor Yellow
    & "$ProjectRoot\build_addin.ps1"
}

# 2. Compile standalone runner into .artifacts/tests/<run-id>/
Write-Host "[2/4] Compiling standalone DesktopAcceptanceRunner..." -ForegroundColor Yellow
$SourceFile = Join-Path $ProjectRoot "tests\tools\DesktopAcceptanceRunner.cs"
$OutExe = Join-Path $ArtifactsDir "DesktopAcceptanceRunner.exe"

$CscArgs = @(
    "/nologo",
    "/target:exe",
    "/out:$OutExe",
    "/lib:$WpfLib",
    "/r:System.dll",
    "/r:System.Core.dll",
    "/r:System.Drawing.dll",
    "/r:System.Windows.Forms.dll",
    "/r:System.Net.Http.dll",
    "/r:System.Security.dll",
    "/r:System.IO.Compression.dll",
    "/r:System.IO.Compression.FileSystem.dll",
    "/r:$ProjectRoot\bin\LeeExcel.dll",
    "/r:WindowsBase.dll",
    "/r:UIAutomationClient.dll",
    "/r:UIAutomationTypes.dll",
    $SourceFile
)

$compileOutput = & $CscPath $CscArgs 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "[FATAL] Compiling DesktopAcceptanceRunner failed:" -ForegroundColor Red
    $compileOutput | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    exit 1
}
Write-Host "    * Compiler success: $OutExe" -ForegroundColor Green

if ($CompileOnly) {
    Write-Host "[PASS] CompileOnly completed; Excel was not started." -ForegroundColor Green
    exit 0
}

# Copy managed runtime dependencies beside the standalone runner.  The runner
# directly references LeeExcel and must resolve its dependencies from its base dir.
$runtimeDlls = Get-ChildItem -LiteralPath (Join-Path $ProjectRoot "bin") -Filter "*.dll" -File
foreach ($runtimeDll in $runtimeDlls) {
    Copy-Item -LiteralPath $runtimeDll.FullName -Destination (Join-Path $ArtifactsDir $runtimeDll.Name) -Force
}
$nativeLoader = Join-Path $ProjectRoot "bin\WebView2Loader.dll"
if (Test-Path -LiteralPath $nativeLoader) {
    Copy-Item -LiteralPath $nativeLoader -Destination (Join-Path $ArtifactsDir "WebView2Loader.dll") -Force
}
if (-not (Test-Path -LiteralPath (Join-Path $ArtifactsDir "LeeExcel.dll"))) {
    throw "Desktop runner dependency copy failed: LeeExcel.dll is missing."
}

# 3. Offline unit tests
Write-Host "[3/4] Running offline regression gates..." -ForegroundColor Yellow
$unitRes = & node "$ProjectRoot\test_suite_unit.cjs"
$unitExit = $LASTEXITCODE
if ($unitExit -ne 0) {
    Write-Host "[FAIL] Unit test suite failed!" -ForegroundColor Red
    exit 1
}
$unitSummary = ($unitRes | Select-String -Pattern 'Unit Tests Summary:.*' | Select-Object -Last 1).Line
if ([string]::IsNullOrWhiteSpace($unitSummary)) { $unitSummary = "Unit Tests Summary: unknown" }
Write-Host "    * $unitSummary" -ForegroundColor Green

$regexRes = & node "$ProjectRoot\test_regex_counter_example.cjs"
if ($LASTEXITCODE -ne 0) {
    Write-Host "[FAIL] Regex counter-example gate failed!" -ForegroundColor Red
    exit 1
}
Write-Host "    * Regex counter-example gate: PASS" -ForegroundColor Green

# 4. Run Desktop Acceptance Tool
Write-Host "[4/4] Executing Real Desktop Acceptance (UIA + COM + TaskPane)..." -ForegroundColor Yellow
& $OutExe $RunId | Tee-Object -FilePath $LogFile
$runnerExit = $LASTEXITCODE

Write-Host "================================================================================" -ForegroundColor Cyan
if ($runnerExit -eq 0) {
    Write-Host "    Acceptance run completed! Report generated: $ArtifactsDir\acceptance_report.md" -ForegroundColor Green
} else {
    Write-Host "    Acceptance run exited with error code: $runnerExit" -ForegroundColor Red
}
Write-Host "================================================================================" -ForegroundColor Cyan

exit $runnerExit
