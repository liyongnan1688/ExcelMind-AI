$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$outDir = Join-Path $projectRoot "bin"

if (!(Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

# 产品版本与提交号由构建脚本注入，避免诊断信息随版本升级产生硬编码漂移。
$packageManifest = Get-Content -LiteralPath (Join-Path $projectRoot "package.json") -Raw | ConvertFrom-Json
$productVersion = [string]$packageManifest.version
if ($productVersion -notmatch '^\d+\.\d+\.\d+$') {
    throw "package.json 版本格式无效: $productVersion"
}
$buildCommit = "Unknown"
$gitCommit = & git -C $projectRoot rev-parse --short HEAD 2>$null
if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($gitCommit)) {
    $buildCommit = $gitCommit.Trim()
}
$buildMetadataDir = Join-Path $projectRoot (".artifacts\tmp\build_metadata_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $buildMetadataDir | Out-Null
$buildMetadataSource = Join-Path $buildMetadataDir "BuildMetadata.cs"
$buildMetadataText = @(
    'using System.Reflection;'
    ('[assembly: AssemblyVersion("{0}")]' -f "$productVersion.0")
    ('[assembly: AssemblyFileVersion("{0}")]' -f "$productVersion.0")
    ('[assembly: AssemblyInformationalVersion("v{0}")]' -f $productVersion)
    ('[assembly: AssemblyMetadata("Commit", "{0}")]' -f $buildCommit)
) -join "`r`n"
[IO.File]::WriteAllText($buildMetadataSource, $buildMetadataText, (New-Object Text.UTF8Encoding($false)))
$refs = @(
    "System.dll",
    "System.Core.dll",
    "System.Windows.Forms.dll",
    "System.Drawing.dll",
    "System.Security.dll",
    "System.Net.Http.dll",
    "System.IO.Compression.dll",
    "System.IO.Compression.FileSystem.dll",
    "Microsoft.CSharp.dll",
    (Resolve-Path (Join-Path $projectRoot "packages\ExcelDna.Integration.1.9.0\lib\net462\ExcelDna.Integration.dll")).Path,
    (Resolve-Path (Join-Path $projectRoot "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.Core.dll")).Path,
    (Resolve-Path (Join-Path $projectRoot "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.WinForms.dll")).Path
)

$refArgs = $refs | ForEach-Object { "/r:`"$_`"" }
$sources = @((Get-ChildItem -Path (Join-Path $projectRoot "src\*.cs") | Select-Object -ExpandProperty FullName) | ForEach-Object { "`"$_`"" }) + @("`"$buildMetadataSource`"")

Write-Host "正在使用系统 .NET Framework 4.8 编译 LeeExcel.dll..." -ForegroundColor Cyan

$cscArgs = @("/nologo", "/target:library", "/out:`"$outDir\LeeExcel.dll`"")
$cscArgs += $refArgs
$cscArgs += $sources
& $csc @cscArgs
if ($LASTEXITCODE -ne 0) {
    Write-Host "编译失败，退出码: $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host "编译成功: $outDir\LeeExcel.dll" -ForegroundColor Green

function Safe-Copy($src, $dst) {
    $sourcePath = if ([IO.Path]::IsPathRooted($src)) { $src } else { Join-Path $projectRoot $src }
    $destPath = if ([IO.Path]::IsPathRooted($dst)) { $dst } else { Join-Path $projectRoot $dst }
    if (!(Test-Path -LiteralPath $destPath)) {
        Copy-Item -LiteralPath $sourcePath -Destination $destPath -Force
    }
}

Safe-Copy "packages\ExcelDna.Integration.1.9.0\lib\net462\ExcelDna.Integration.dll" "$outDir\ExcelDna.Integration.dll"
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.Core.dll" "$outDir\Microsoft.Web.WebView2.Core.dll"
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.WinForms.dll" "$outDir\Microsoft.Web.WebView2.WinForms.dll"
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x64\native\WebView2Loader.dll" "$outDir\WebView2Loader.dll"
Safe-Copy "packages\ExcelDna.AddIn.1.9.0\tools\net462\ExcelDna.xll" "$outDir\LeeExcel.xll"
Safe-Copy "packages\ExcelDna.AddIn.1.9.0\tools\net462\ExcelDna64.xll" "$outDir\LeeExcel64.xll"
Copy-Item -LiteralPath (Join-Path $projectRoot "LeeExcel.dna") -Destination (Join-Path $outDir "LeeExcel.dna") -Force
Copy-Item -LiteralPath (Join-Path $projectRoot "LeeExcel64.dna") -Destination (Join-Path $outDir "LeeExcel64.dna") -Force

# 组织 runtimes 双架构原生加载器目录
$x64Dir = Join-Path $outDir "runtimes\win-x64\native"
$x86Dir = Join-Path $outDir "runtimes\win-x86\native"
if (!(Test-Path $x64Dir)) { New-Item -ItemType Directory -Path $x64Dir -Force | Out-Null }
if (!(Test-Path $x86Dir)) { New-Item -ItemType Directory -Path $x86Dir -Force | Out-Null }
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x64\native\WebView2Loader.dll" "$x64Dir\WebView2Loader.dll"
Safe-Copy "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x86\native\WebView2Loader.dll" "$x86Dir\WebView2Loader.dll"

Write-Host "原生加载项双架构组织就绪: 32位($outDir\LeeExcel.xll) 与 64位($outDir\LeeExcel64.xll)" -ForegroundColor Green

