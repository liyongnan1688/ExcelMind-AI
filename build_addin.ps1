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

# 拷贝运行时依赖
Copy-Item "packages\ExcelDna.Integration.1.9.0\lib\net462\ExcelDna.Integration.dll" $outDir -Force
Copy-Item "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.Core.dll" $outDir -Force
Copy-Item "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.WinForms.dll" $outDir -Force
Copy-Item "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x64\native\WebView2Loader.dll" $outDir -Force

# 拷贝 Excel-DNA 宿主 xll 并重命名
Copy-Item "packages\ExcelDna.AddIn.1.9.0\tools\net462\ExcelDna64.xll" "$outDir\LeeExcel64.xll" -Force
Copy-Item "LeeExcel.dna" $outDir -Force
Copy-Item "LeeExcel64.dna" $outDir -Force

Write-Host "原生加载项组织就绪: $outDir\LeeExcel64.xll" -ForegroundColor Green
