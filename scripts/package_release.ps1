# ExcelMind AI 跨机器发布包全自动打包流水线
# 职责：编译前端与后端、组织双架构依赖、生成安装卸载套件、打包为独立 ZIP

param(
    [string]$Version = "v1.0"
)

$ErrorActionPreference = "Stop"

# 统一控制台输出编码为 UTF-8，防止在 Windows PowerShell 5.1 下输出中文字符乱码
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$projectRoot = Split-Path -Parent $PSScriptRoot

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "         ExcelMind AI 跨平台/跨版本 发布包自动化打包流水线      " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "版本: $Version"
Write-Host "项目根路径: $projectRoot"

# 1. 确保输出目录安全隔离 (.artifacts/ 规范)
$artifactsDir = Join-Path $projectRoot ".artifacts"
$releaseDir = Join-Path $artifactsDir "release"
$tmpStageDir = Join-Path $artifactsDir "tmp\pack_stage_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

if (!(Test-Path $releaseDir)) {
    New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
}
if (Test-Path $tmpStageDir) {
    Remove-Item -Path $tmpStageDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tmpStageDir -Force | Out-Null

$packRoot = Join-Path $tmpStageDir "ExcelMind_AI_Release"
New-Item -ItemType Directory -Path $packRoot -Force | Out-Null

# 2. 构建前端生产包 (Svelte 5 + Vite 6)
Write-Host "`n[1/5] 正在构建前端生产包 (pnpm run build)..." -ForegroundColor Yellow
$webDir = Join-Path $projectRoot "web"
$pnpmCmd = Get-Command pnpm -ErrorAction SilentlyContinue
if (!$pnpmCmd) {
    Write-Host "      未找到 pnpm 命令，尝试使用 npm..." -ForegroundColor DarkGray
    Push-Location $webDir
    if (!(Test-Path "node_modules")) {
        Write-Host "      正在安装前端依赖 (npm install)..." -ForegroundColor Yellow
        npm install
    }
    npm run build
    Pop-Location
} else {
    Push-Location $webDir
    if (!(Test-Path "node_modules")) {
        Write-Host "      正在安装前端依赖 (pnpm install)..." -ForegroundColor Yellow
        pnpm install
    }
    pnpm run build
    Pop-Location
}

$frontendDist = Join-Path $projectRoot "bin\dist"
if (!(Test-Path (Join-Path $frontendDist "index.html"))) {
    throw "前端打包失败：未在 bin\dist 中找到 index.html"
}
Write-Host "      前端生产包就绪: $frontendDist" -ForegroundColor Green

# 3. 编译 C# 插件主程序集 (LeeExcel.dll - Any CPU)
Write-Host "`n[2/5] 正在使用系统 .NET 编译 C# 核心库 (Any CPU)..." -ForegroundColor Yellow
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (!(Test-Path $csc)) {
    throw "未找到系统 csc.exe 编译器: $csc"
}

$binDir = Join-Path $projectRoot "bin"
$dllOut = Join-Path $binDir "LeeExcel.dll"

$refs = @(
    "System.dll",
    "System.Core.dll",
    "System.Windows.Forms.dll",
    "System.Drawing.dll",
    "Microsoft.CSharp.dll",
    (Resolve-Path (Join-Path $projectRoot "packages\ExcelDna.Integration.1.9.0\lib\net462\ExcelDna.Integration.dll")).Path,
    (Resolve-Path (Join-Path $projectRoot "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.Core.dll")).Path,
    (Resolve-Path (Join-Path $projectRoot "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.WinForms.dll")).Path
)
$refArgs = $refs | ForEach-Object { "/r:`"$_`"" }
$sources = (Get-ChildItem -Path (Join-Path $projectRoot "src\*.cs") | Select-Object -ExpandProperty FullName) | ForEach-Object { "`"$_`"" }

$buildCmd = "& `"$csc`" /nologo /target:library /out:`"$dllOut`" $refArgs $sources"
Invoke-Expression $buildCmd
if ($LASTEXITCODE -ne 0) {
    throw "C# 编译失败，退出码: $LASTEXITCODE"
}
Write-Host "      编译成功: $dllOut" -ForegroundColor Green

# 4. 组装双架构发布包目录树
Write-Host "`n[3/5] 正在组织 32位与64位 双架构依赖文件..." -ForegroundColor Yellow

# 基础文件
Copy-Item $dllOut (Join-Path $packRoot "LeeExcel.dll") -Force
Copy-Item (Join-Path $projectRoot "packages\ExcelDna.Integration.1.9.0\lib\net462\ExcelDna.Integration.dll") (Join-Path $packRoot "ExcelDna.Integration.dll") -Force
Copy-Item (Join-Path $projectRoot "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.Core.dll") (Join-Path $packRoot "Microsoft.Web.WebView2.Core.dll") -Force
Copy-Item (Join-Path $projectRoot "packages\Microsoft.Web.WebView2.1.0.4191.47\lib\net462\Microsoft.Web.WebView2.WinForms.dll") (Join-Path $packRoot "Microsoft.Web.WebView2.WinForms.dll") -Force

# 32位与64位 XLL
Copy-Item (Join-Path $projectRoot "packages\ExcelDna.AddIn.1.9.0\tools\net462\ExcelDna.xll") (Join-Path $packRoot "LeeExcel.xll") -Force
Copy-Item (Join-Path $projectRoot "packages\ExcelDna.AddIn.1.9.0\tools\net462\ExcelDna64.xll") (Join-Path $packRoot "LeeExcel64.xll") -Force

# DNA 配置
Copy-Item (Join-Path $projectRoot "LeeExcel.dna") (Join-Path $packRoot "LeeExcel.dna") -Force
Copy-Item (Join-Path $projectRoot "LeeExcel64.dna") (Join-Path $packRoot "LeeExcel64.dna") -Force

# WebView2Loader 原生运行库 (win-x64 与 win-x86)
$x64LoaderDir = Join-Path $packRoot "runtimes\win-x64\native"
$x86LoaderDir = Join-Path $packRoot "runtimes\win-x86\native"
New-Item -ItemType Directory -Path $x64LoaderDir -Force | Out-Null
New-Item -ItemType Directory -Path $x86LoaderDir -Force | Out-Null

Copy-Item (Join-Path $projectRoot "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x64\native\WebView2Loader.dll") (Join-Path $x64LoaderDir "WebView2Loader.dll") -Force
Copy-Item (Join-Path $projectRoot "packages\Microsoft.Web.WebView2.1.0.4191.47\runtimes\win-x86\native\WebView2Loader.dll") (Join-Path $x86LoaderDir "WebView2Loader.dll") -Force
# 根目录兜底默认放置 64 位 Loader
Copy-Item (Join-Path $x64LoaderDir "WebView2Loader.dll") (Join-Path $packRoot "WebView2Loader.dll") -Force

# 前端 dist
$distTarget = Join-Path $packRoot "dist"
Copy-Item -Path $frontendDist -Destination $distTarget -Recurse -Force

# 引导批处理脚本
Copy-Item -LiteralPath (Join-Path $projectRoot "scripts\core\bat_templates\install.bat") -Destination (Join-Path $packRoot "安装插件(开启常驻).bat") -Force
Copy-Item -LiteralPath (Join-Path $projectRoot "scripts\core\bat_templates\uninstall.bat") -Destination (Join-Path $packRoot "卸载插件.bat") -Force
Copy-Item -LiteralPath (Join-Path $projectRoot "scripts\core\bat_templates\portable.bat") -Destination (Join-Path $packRoot "免安装启动.bat") -Force

# 核心脚本引擎目录 core/
$coreTarget = Join-Path $packRoot "core"
New-Item -ItemType Directory -Path $coreTarget -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $projectRoot "scripts\core\install_addin.ps1") -Destination (Join-Path $coreTarget "install_addin.ps1") -Force
Copy-Item -LiteralPath (Join-Path $projectRoot "scripts\core\uninstall_addin.ps1") -Destination (Join-Path $coreTarget "uninstall_addin.ps1") -Force
Copy-Item -LiteralPath (Join-Path $projectRoot "scripts\core\launch_portable.ps1") -Destination (Join-Path $coreTarget "launch_portable.ps1") -Force

# 使用说明书
Copy-Item -LiteralPath (Join-Path $projectRoot "scripts\core\README_template.txt") -Destination (Join-Path $packRoot "README_使用说明.txt") -Force

Write-Host "      发布目录树组装完成: $packRoot" -ForegroundColor Green

# 5. 压缩为发布 ZIP 包
$zipName = "ExcelMind_AI_Release_${Version}.zip"
$zipPath = Join-Path $releaseDir $zipName

Write-Host "`n[4/5] 正在打包压缩发布文件包: $zipName ..." -ForegroundColor Yellow
if (Test-Path $zipPath) {
    Remove-Item -Path $zipPath -Force
}

Compress-Archive -Path $packRoot -DestinationPath $zipPath -CompressionLevel Optimal

# 6. 清理临时暂存区
Remove-Item -Path $tmpStageDir -Recurse -Force

# 7. 校验产物与输出报告
$zipItem = Get-Item $zipPath
$zipSizeMB = [math]::Round($zipItem.Length / 1MB, 2)

Write-Host "`n[5/5] 打包成功！" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "发布包路径: $zipPath" -ForegroundColor Cyan
Write-Host "文件大小  : $zipSizeMB MB" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "包含内容清单："
Write-Host "  - LeeExcel.xll                   (32位 Excel加载项)"
Write-Host "  - LeeExcel64.xll                 (64位 Excel加载项)"
Write-Host "  - LeeExcel.dll                   (C# 核心功能库)"
Write-Host "  - runtimes/win-x64 & win-x86     (双架构 WebView2 原生加载器)"
Write-Host "  - dist/                          (Svelte 5 前端完整资产)"
Write-Host "  - 安装插件(开启常驻).bat          (全自动检测与注册工具)"
Write-Host "  - 卸载插件.bat                   (一键干净注销工具)"
Write-Host "  - 免安装启动.bat                 (绿色即用启动器)"
Write-Host "  - README_使用说明.txt            (用户使用手册)"
Write-Host ""
