$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$RunId = "security_baseline_" + (Get-Date -Format "yyyyMMdd_HHmmss")
$OutDir = Join-Path $ProjectRoot (Join-Path ".artifacts\tests" $RunId)
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outExe = Join-Path $OutDir "VerifySecurityBaseline.exe"
$source = Join-Path $ProjectRoot "tests\tools\VerifySecurityBaseline.cs"
$library = Join-Path $ProjectRoot "bin\LeeExcel.dll"

if (-not (Test-Path -LiteralPath $library)) {
    throw "缺少编译产物: $library，请先运行 build_addin.ps1"
}

$runtimeFiles = @(
    "LeeExcel.dll",
    "ExcelDna.Integration.dll",
    "Microsoft.Web.WebView2.Core.dll",
    "Microsoft.Web.WebView2.WinForms.dll",
    "WebView2Loader.dll"
)
foreach ($fileName in $runtimeFiles) {
    Copy-Item -LiteralPath (Join-Path $ProjectRoot "bin\$fileName") -Destination (Join-Path $OutDir $fileName) -Force
}

& $csc /nologo /target:exe /out:"$outExe" /r:System.dll /r:System.Core.dll /r:"$library" "$source"
if ($LASTEXITCODE -ne 0) {
    throw "安全基线测试编译失败，退出码: $LASTEXITCODE"
}

$env:LEE_EXCEL_PROJECT_ROOT = $ProjectRoot
& $outExe $ProjectRoot
if ($LASTEXITCODE -ne 0) {
    throw "安全基线测试失败，退出码: $LASTEXITCODE"
}

Write-Host "[PASS] 安全基线测试通过，产物: $OutDir" -ForegroundColor Green