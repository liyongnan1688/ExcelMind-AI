# Lee-Excel 自动化安装与环境配置脚本
# 适用：Office 2010 / 2013 / 2016 / 2019 / 2021 / Microsoft 365 (32位与64位)

$ErrorActionPreference = "Continue"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "           Lee-Excel AI 原生插件 - 自动化安装向导           " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. 定位插件发布目录 (即 core/ 的父目录)
$releaseDir = Split-Path -Parent $PSScriptRoot
if (!(Test-Path (Join-Path $releaseDir "LeeExcel64.xll"))) {
    $releaseDir = $PSScriptRoot
}

Write-Host "[1/5] 正在解除外来下载文件的系统安全锁定 (Mark-of-the-Web)..." -ForegroundColor Yellow
try {
    Get-ChildItem -Path $releaseDir -Recurse -Include *.dll,*.xll,*.dna,*.html,*.js,*.css | ForEach-Object {
        Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue
    }
    Write-Host "      文件安全锁定已全部解除。" -ForegroundColor Green
} catch {
    Write-Host "      解锁提示: $($_.Exception.Message)" -ForegroundColor DarkGray
}

# 2. 检测 WebView2 运行时
Write-Host "[2/5] 正在检测 Microsoft Edge WebView2 运行时..." -ForegroundColor Yellow
function Get-WebView2Version {
    $regPaths = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-F55F-4E4E-9A04-E8763F00F763}",
        "HKCU:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-F55F-4E4E-9A04-E8763F00F763}",
        "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-F55F-4E4E-9A04-E8763F00F763}"
    )
    foreach ($p in $regPaths) {
        if (Test-Path $p) {
            $ver = (Get-ItemProperty -Path $p -Name "pv" -ErrorAction SilentlyContinue).pv
            if (![string]::IsNullOrEmpty($ver) -and $ver -ne "0.0.0.0") {
                return $ver
            }
        }
    }
    return $null
}

$wv2Ver = Get-WebView2Version
if (![string]::IsNullOrEmpty($wv2Ver)) {
    Write-Host "      已安装 WebView2 运行时 (版本: $wv2Ver)" -ForegroundColor Green
} else {
    Write-Host "      [提示] 未在当前系统中检测到 Microsoft Edge WebView2 运行时！" -ForegroundColor Red
    Write-Host "      WebView2 是插件右侧 AI 交互窗格的必要渲染引擎。" -ForegroundColor Yellow
    $ans = Read-Host "      是否立即从微软官方自动下载并静默安装 WebView2？(Y/n，按回车默认安装)"
    if ([string]::IsNullOrEmpty($ans) -or $ans.Trim().ToLower() -eq 'y') {
        Write-Host "      正在从微软官方 CDN 下载轻量引导安装器 (约2MB)..." -ForegroundColor Cyan
        $bootstrapperUrl = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
        $tmpInstaller = Join-Path $env:TEMP "MicrosoftEdgeWebview2Setup.exe"
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $bootstrapperUrl -OutFile $tmpInstaller -UseBasicParsing
            Write-Host "      下载完成，正在静默安装 WebView2，请稍候约10~30秒..." -ForegroundColor Cyan
            $p = Start-Process -FilePath $tmpInstaller -ArgumentList "/silent /install" -Wait -PassThru
            Write-Host "      WebView2 安装完成！退出码: $($p.ExitCode)" -ForegroundColor Green
        } catch {
            Write-Host "      自动下载安装失败: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "      请手动访问微软官网下载安装: https://developer.microsoft.com/microsoft-edge/webview2/" -ForegroundColor Yellow
        }
    } else {
        Write-Host "      已跳过安装。若打开插件后任务窗格提示组件缺失，请手动安装 WebView2。" -ForegroundColor Yellow
    }
}

# 3. 定位 Excel 并精确检测其架构 (32 位 vs 64 位)
Write-Host "[3/5] 正在识别本机 Excel 安装版本与架构位数..." -ForegroundColor Yellow

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
$bitness = 64
if (![string]::IsNullOrEmpty($excelPath)) {
    $bitness = Get-ExeBitness $excelPath
    Write-Host "      找到 Excel 主程序: $excelPath" -ForegroundColor DarkGray
    Write-Host "      识别到 Excel 架构: $bitness 位" -ForegroundColor Green
} else {
    Write-Host "      [提示] 未能在常见路径中自动锁定 excel.exe，将根据当前 Windows 架构进行预判 ($bitness 位)。" -ForegroundColor Yellow
}

$targetXll = if ($bitness -eq 32) {
    Join-Path $releaseDir "LeeExcel.xll"
} else {
    Join-Path $releaseDir "LeeExcel64.xll"
}

if (!(Test-Path $targetXll)) {
    Write-Host "      [错误] 找不到对应的加载项文件: $targetXll" -ForegroundColor Red
    exit 1
}

# 4. 注册到 Excel 启动项 (Office 2010 ~ Office 365 多版本扫描注入)
Write-Host "[4/5] 正在向 Excel 注册自启动项..." -ForegroundColor Yellow

$officeVersions = @("16.0", "15.0", "14.0") # 16.0 = 2016/2019/2021/365, 15.0 = 2013, 14.0 = 2010
$regSuccessCount = 0

foreach ($ver in $officeVersions) {
    $optionsPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (Test-Path "HKCU:\Software\Microsoft\Office\$ver\Excel") {
        if (!(Test-Path $optionsPath)) {
            New-Item -Path $optionsPath -Force | Out-Null
        }

        # 检查是否已经注册该插件 (避免重复写入)
        $props = Get-ItemProperty -Path $optionsPath -ErrorAction SilentlyContinue
        $alreadyRegisteredKey = $null

        foreach ($propName in ($props.PSObject.Properties.Name | Where-Object { $_ -match "^OPEN\d*$" })) {
            $val = $props.$propName
            if ($val -like "*LeeExcel*") {
                $alreadyRegisteredKey = $propName
                break
            }
        }

        $regValue = "/R `"$targetXll`""

        if (![string]::IsNullOrEmpty($alreadyRegisteredKey)) {
            # 更新已有项的路径
            Set-ItemProperty -Path $optionsPath -Name $alreadyRegisteredKey -Value $regValue
            Write-Host "      已更新 Office $ver 启动项 [$alreadyRegisteredKey]: $targetXll" -ForegroundColor Green
            $regSuccessCount++
        } else {
            # 寻找下一个可用的 OPEN 键编号
            $targetKey = "OPEN"
            $counter = 1
            while ($props.PSObject.Properties.Name -contains $targetKey) {
                $targetKey = "OPEN$counter"
                $counter++
            }

            New-ItemProperty -Path $optionsPath -Name $targetKey -Value $regValue -PropertyType String -Force | Out-Null
            Write-Host "      已注入 Office $ver 启动项 [$targetKey]: $targetXll" -ForegroundColor Green
            $regSuccessCount++
        }
    }
}

if ($regSuccessCount -eq 0) {
    # 如果用户尚未配置过 Office 注册表项，主动为 Office 16.0 创建
    $fallbackPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
    if (!(Test-Path $fallbackPath)) {
        New-Item -Path $fallbackPath -Force | Out-Null
    }
    New-ItemProperty -Path $fallbackPath -Name "OPEN" -Value "/R `"$targetXll`"" -PropertyType String -Force | Out-Null
    Write-Host "      已为通用 Office 16.0 注册自启动项: $targetXll" -ForegroundColor Green
}

# 5. 完成提示
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "               Lee-Excel AI 插件安装完成！               " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "使用指引：" -ForegroundColor Cyan
Write-Host "1. 打开任意 Excel 工作簿；"
Write-Host "2. 顶部功能区将自动出现【AI 助手】专属选项卡；"
Write-Host "3. 点击【打开 AI 任务窗格】即可享受全功能数据助手体验！"
Write-Host ""
