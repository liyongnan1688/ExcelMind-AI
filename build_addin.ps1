$ErrorActionPreference = "Stop"

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outDir = "bin"

if (!(Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

$refs = @(
    "System.dll",
    "System.Core.dll",
    "System.Windows.Forms.dll",
    "System.Drawing.dll",
    "Microsoft.CSharp.dll",
    (Resolve-Path "packages\ExcelDna.Integration.1.9.0\lib\net462\ExcelDna.Integration.dll").Path,
    (Resolve-Path "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.Core.dll").Path,
    (Resolve-Path "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.WinForms.dll").Path
)

$refArgs = $refs | ForEach-Object { "/r:`"$_`"" }
$sources = (Get-ChildItem -Path "src\*.cs" | Select-Object -ExpandProperty FullName) | ForEach-Object { "`"$_`"" }

Write-Host "正在使用系统 .NET Framework 4.8 编译 LeeExcel.dll..." -ForegroundColor Cyan

$cmd = "& `"$csc`" /nologo /target:library /out:`"$outDir\LeeExcel.dll`" $refArgs $sources"
Invoke-Expression $cmd
if ($LASTEXITCODE -ne 0) {
    Write-Host "编译失败，退出码: $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host "编译成功: $outDir\LeeExcel.dll" -ForegroundColor Green

function Safe-Copy($src, $dst) {
    if (!(Test-Path $dst)) {
        Copy-Item $src $dst -Force
    }
}

Safe-Copy "packages\ExcelDna.Integration.1.9.0\lib\net462\ExcelDna.Integration.dll" "$outDir\ExcelDna.Integration.dll"
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.Core.dll" "$outDir\Microsoft.Web.WebView2.Core.dll"
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.WinForms.dll" "$outDir\Microsoft.Web.WebView2.WinForms.dll"
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x64\native\WebView2Loader.dll" "$outDir\WebView2Loader.dll"
Safe-Copy "packages\ExcelDna.AddIn.1.9.0\tools\net462\ExcelDna.xll" "$outDir\LeeExcel.xll"
Safe-Copy "packages\ExcelDna.AddIn.1.9.0\tools\net462\ExcelDna64.xll" "$outDir\LeeExcel64.xll"
Safe-Copy "LeeExcel.dna" "$outDir\LeeExcel.dna"
Safe-Copy "LeeExcel64.dna" "$outDir\LeeExcel64.dna"

# 组织 runtimes 双架构原生加载器目录
$x64Dir = "$outDir\runtimes\win-x64\native"
$x86Dir = "$outDir\runtimes\win-x86\native"
if (!(Test-Path $x64Dir)) { New-Item -ItemType Directory -Path $x64Dir -Force | Out-Null }
if (!(Test-Path $x86Dir)) { New-Item -ItemType Directory -Path $x86Dir -Force | Out-Null }
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x64\native\WebView2Loader.dll" "$x64Dir\WebView2Loader.dll"
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x86\native\WebView2Loader.dll" "$x86Dir\WebView2Loader.dll"

Write-Host "原生加载项双架构组织就绪: 32位($outDir\LeeExcel.xll) 与 64位($outDir\LeeExcel64.xll)" -ForegroundColor Green

