# ExcelMind AI 便携式免安装启动器 (Portable Launcher)
# 自动检测当前 Excel 架构并挂载对应 XLL 启动，不修改注册表

$ErrorActionPreference = "Continue"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "           ExcelMind AI 原生插件 - 便携启动向导           " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. 定位插件发布目录 (即 core/ 的父目录)
$releaseDir = Split-Path -Parent $PSScriptRoot
if (!(Test-Path (Join-Path $releaseDir "LeeExcel64.xll"))) {
    $releaseDir = $PSScriptRoot
}

# 解除安全锁定
try {
    Get-ChildItem -Path $releaseDir -Recurse -Include *.dll,*.xll,*.dna,*.html,*.js,*.css | ForEach-Object {
        Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue
    }
} catch { }

# 2. 定位 Excel 并精确检测其架构 (32 位 vs 64 位)
function Find-ExcelPath {
    $reg = (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue).'(default)'
    if (![string]::IsNullOrEmpty($reg) -and (Test-Path $reg)) { return $reg }
    
    $commonPaths = @(
        "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
        "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
        "C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
        "C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE",
        "C:\Program Files\Microsoft Office\Office15\EXCEL.EXE",
        "C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE",
        "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE",
        "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE"
    )
    foreach ($p in $commonPaths) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

function Get-ExeBitness($filePath) {
    if (![string]::IsNullOrEmpty($filePath) -and (Test-Path $filePath)) {
        try {
            $stream = [System.IO.File]::OpenRead($filePath)
            $reader = New-Object System.IO.BinaryReader($stream)
            $mz = $reader.ReadUInt16()
            if ($mz -eq 0x5A4D) {
                $stream.Position = 0x3C
                $peOffset = $reader.ReadUInt32()
                $stream.Position = $peOffset
                $peSig = $reader.ReadUInt32()
                if ($peSig -eq 0x00004550) {
                    $machine = $reader.ReadUInt16()
                    $stream.Dispose()
                    if ($machine -eq 0x014c) { return 32 }
                    if ($machine -eq 0x8664) { return 64 }
                }
            }
            $stream.Dispose()
        } catch { }
    }
    if ([Environment]::Is64BitOperatingSystem) { return 64 } else { return 32 }
}

$excelPath = Find-ExcelPath
if ([string]::IsNullOrEmpty($excelPath)) {
    Write-Host "[错误] 未能在本机找到可执行的 excel.exe，请确保系统已安装 Microsoft Excel！" -ForegroundColor Red
    exit 1
}

$bitness = Get-ExeBitness $excelPath
$targetXll = if ($bitness -eq 32) {
    Join-Path $releaseDir "LeeExcel.xll"
} else {
    Join-Path $releaseDir "LeeExcel64.xll"
}

if (!(Test-Path $targetXll)) {
    Write-Host "[错误] 找不到对应的加载项文件: $targetXll" -ForegroundColor Red
    exit 1
}

Write-Host "检测到 Excel: $excelPath" -ForegroundColor DarkGray
Write-Host "匹配架构版本: $bitness 位" -ForegroundColor Green
Write-Host "正在启动 Excel 并挂载加载项: $targetXll ..." -ForegroundColor Cyan

Start-Process -FilePath $excelPath -ArgumentList "`"$targetXll`""
Write-Host "已发送启动指令！" -ForegroundColor Green
