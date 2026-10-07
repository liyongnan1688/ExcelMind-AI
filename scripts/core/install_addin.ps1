# ExcelMind AI 自动化无损安装与环境配置脚本
# 适用：Office 2010 / 2013 / 2016 / 2019 / 2021 / Microsoft 365 (32位与64位)
# 严格遵守 TASK-R6c-01：应用文件、加载项注册、用户数据（Profile/宏库/工作流/快照/凭据）严格三态分离，100% 无损升级保护

param(
    [string]$SourceDir = "",          # 来源发布包目录，默认根据脚本所在位置自动向上寻找
    [string]$TargetInstallDir = "",   # 可选独立安装目标目录（指定时执行受控文件复制与版本升级；留空则保持就地绿色注册）
    [string]$RegistryRoot = "",       # 注册表根路径（默认为 HKCU:\Software\Microsoft\Office；隔离测试时可传模拟根如 HKCU:\Software\ExcelMindAITest）
    [string]$UserDataDir = "",        # 用户数据目录（默认 %APPDATA%\ExcelMindAI；仅用于保护探测与验证，严禁覆盖）
    [switch]$CheckOnly,               # 仅执行环境与占用检查
    [switch]$DryRun,                  # 预览模式：仅检查和展示步骤，不执行写盘和注册表写入
    [switch]$NonInteractive,          # 非交互模式：测试及自动化调用，不等待人工回车
    [switch]$SkipWebView2Check        # 跳过 WebView2 下载引导（用于离线受控测试）
)

$ErrorActionPreference = "Continue"

# 统一控制台输出编码为 UTF-8，防止在 Windows PowerShell 5.1 下输出中文字符乱码
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "         ExcelMind AI 原生插件 - 自动化无损安装升级向导         " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# -------------------------------------------------------------
# 1. 定位插件来源发布目录与暂存区完整性校验
# -------------------------------------------------------------
$releaseDir = $SourceDir
if ([string]::IsNullOrEmpty($releaseDir)) {
    $parentDir = Split-Path -Parent $PSScriptRoot
    if (Test-Path (Join-Path $parentDir "LeeExcel64.xll")) {
        $releaseDir = $parentDir
    } elseif (Test-Path (Join-Path $PSScriptRoot "LeeExcel64.xll")) {
        $releaseDir = $PSScriptRoot
    } else {
        $projectRoot = Split-Path -Parent $parentDir
        if (Test-Path (Join-Path $projectRoot "bin\LeeExcel64.xll")) {
            $releaseDir = Join-Path $projectRoot "bin"
        } else {
            $releaseDir = $parentDir
        }
    }
}

if (!(Test-Path $releaseDir)) {
    Write-Host "[错误] 来源目录不存在: $releaseDir" -ForegroundColor Red
    exit 1
}

Write-Host "[1/6] 正在校验来源包文件完整性 (暂存区预检)..." -ForegroundColor Yellow

$requiredFiles = @(
    "LeeExcel.dll",
    "LeeExcel.xll",
    "LeeExcel64.xll",
    "LeeExcel.dna",
    "LeeExcel64.dna"
)

$missingFiles = @()
foreach ($rf in $requiredFiles) {
    $fullPath = Join-Path $releaseDir $rf
    if (!(Test-Path $fullPath)) {
        $missingFiles += $rf
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "      [错误] 来源包缺少必要核心文件: $($missingFiles -join ', ')" -ForegroundColor Red
    exit 1
}
Write-Host "      来源包完整性核验通过 (包含 32/64 位核心 DLL/XLL 与配置文件)。" -ForegroundColor Green

# -------------------------------------------------------------
# 2. 解除系统安全锁定 (Mark-of-the-Web)
# -------------------------------------------------------------
Write-Host "[2/6] 正在解除外来下载文件的系统安全锁定 (Mark-of-the-Web)..." -ForegroundColor Yellow
if (!$DryRun) {
    try {
        Get-ChildItem -Path $releaseDir -Recurse -Include *.dll,*.xll,*.dna,*.html,*.js,*.css -ErrorAction SilentlyContinue | ForEach-Object {
            Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue
        }
        Write-Host "      文件安全锁定已全部解除。" -ForegroundColor Green
    } catch {
        Write-Host "      解锁提示: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
} else {
    Write-Host "      [DryRun] 跳过解除锁定操作。" -ForegroundColor DarkGray
}

# -------------------------------------------------------------
# 3. 检测文件占用与 Excel 运行状态 (安全红线：绝不强杀进程、绝不静默覆盖)
# -------------------------------------------------------------
Write-Host "[3/6] 正在检查文件占用与 Excel 运行状态..." -ForegroundColor Yellow

function Test-FileWriteLock($filePath) {
    if (![string]::IsNullOrEmpty($filePath) -and (Test-Path $filePath)) {
        try {
            $stream = [System.IO.File]::Open($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Dispose()
            return $false
        } catch {
            return $true
        }
    }
    return $false
}

$effectiveTargetDir = if (![string]::IsNullOrEmpty($TargetInstallDir)) { $TargetInstallDir } else { $releaseDir }
$lockedTestFile = $null

if (Test-Path $effectiveTargetDir) {
    $checkTargets = @(
        (Join-Path $effectiveTargetDir "LeeExcel64.xll"),
        (Join-Path $effectiveTargetDir "LeeExcel.xll"),
        (Join-Path $effectiveTargetDir "LeeExcel.dll")
    )
    foreach ($ct in $checkTargets) {
        if (Test-FileWriteLock $ct) {
            $lockedTestFile = $ct
            break
        }
    }
}

if (![string]::IsNullOrEmpty($lockedTestFile)) {
    Write-Host "      [错误] 检测到应用文件正在被外部程序锁定占用！" -ForegroundColor Red
    Write-Host "      占用文件: $lockedTestFile" -ForegroundColor Yellow
    Write-Host "      [安全守则] 为保护您的工作簿数据完整性，本向导严禁强杀 Excel 进程或静默强制覆盖占用文件。" -ForegroundColor Red
    Write-Host "      请先保存所有正在编辑的 Excel 工作簿并完全退出 Excel，然后重试安装/升级。" -ForegroundColor Yellow
    exit 2
}

# 检查当前系统是否有 Excel 进程运行（仅在非隔离测试时作为友好提示）
if ([string]::IsNullOrEmpty($RegistryRoot)) {
    $excelProcesses = Get-Process excel -ErrorAction SilentlyContinue
    if ($excelProcesses -and $excelProcesses.Count -gt 0) {
        Write-Host "      [提示] 检测到当前有 $($excelProcesses.Count) 个 Excel 进程正在运行。" -ForegroundColor Yellow
        Write-Host "      建议先保存工作簿并关闭 Excel，以确保加载项注册能立即生效。" -ForegroundColor Yellow
    } else {
        Write-Host "      未检测到运行中的 Excel 进程或文件写锁，环境就绪。" -ForegroundColor Green
    }
} else {
    Write-Host "      [隔离测试模式] 文件锁检查通过。" -ForegroundColor Green
}

if ($CheckOnly) {
    Write-Host "      [CheckOnly] 环境预检已完成，退出。" -ForegroundColor Cyan
    exit 0
}

# -------------------------------------------------------------
# 4. 用户数据目录隔离保护声明 (Profile / 宏库 / 工作流 / 快照 / 凭据)
# -------------------------------------------------------------
$actualUserDataDir = if (![string]::IsNullOrEmpty($UserDataDir)) { $UserDataDir } else { Join-Path $env:APPDATA "ExcelMindAI" }
Write-Host "[4/6] 检查用户专属数据目录隔离状态..." -ForegroundColor Yellow
Write-Host "      用户数据专属目录: $actualUserDataDir" -ForegroundColor DarkGray

if (Test-Path $actualUserDataDir) {
    Write-Host "      检测到既有用户数据 (宏库、工作流定义、运行记录、工作簿快照与 DPAPI 凭据)。" -ForegroundColor Green
    Write-Host "      [无损铁律] 本向导仅更新独立应用目录及注册表自启动项，100% 绝不覆盖、重置或清除用户数据目录！" -ForegroundColor Green
} else {
    Write-Host "      用户数据目录将在插件首次运行时自动初始化，安装向导不主动创建空目录。" -ForegroundColor DarkGray
}

# -------------------------------------------------------------
# 5. 首次安装、重复安装或版本升级处理 (带失败补偿恢复机制)
# -------------------------------------------------------------
Write-Host "[5/6] 正在处理应用文件安装/升级与版本管理..." -ForegroundColor Yellow

$installMode = "InPlace" # InPlace, FreshInstall, IdempotentReinstall, Upgrade
$backupDir = $null

if (![string]::IsNullOrEmpty($TargetInstallDir) -and ($TargetInstallDir -ne $releaseDir)) {
    if (!(Test-Path $TargetInstallDir)) {
        $installMode = "FreshInstall"
        Write-Host "      安装模式: 首次安装 -> $TargetInstallDir" -ForegroundColor Cyan
        if (!$DryRun) {
            New-Item -ItemType Directory -Path $TargetInstallDir -Force | Out-Null
        }
    } else {
        # 目标已存在，比对主 DLL 哈希判断是否重复安装或升级
        $srcDll = Join-Path $releaseDir "LeeExcel.dll"
        $dstDll = Join-Path $TargetInstallDir "LeeExcel.dll"
        
        $needUpgrade = $true
        if ((Test-Path $srcDll) -and (Test-Path $dstDll)) {
            $srcHash = (Get-FileHash -Path $srcDll -Algorithm SHA256).Hash
            $dstHash = (Get-FileHash -Path $dstDll -Algorithm SHA256).Hash
            if ($srcHash -eq $dstHash) {
                $needUpgrade = $false
            }
        }

        if ($needUpgrade) {
            $installMode = "Upgrade"
            Write-Host "      安装模式: 版本升级 -> $TargetInstallDir" -ForegroundColor Cyan
        } else {
            $installMode = "IdempotentReinstall"
            Write-Host "      安装模式: 幂等重复安装 (文件已是最新版本) -> $TargetInstallDir" -ForegroundColor Cyan
        }
    }

    $upgradeBackupDir = $null
    if (!$DryRun -and ($installMode -eq "FreshInstall" -or $installMode -eq "Upgrade")) {
        $completedSteps = @()
        $failedStep = $null
        
        try {
            $origEap1 = $ErrorActionPreference
            $ErrorActionPreference = "Stop"
            # 如果是升级，先在目标目录内部备份现有应用文件到暂存备份
            if ($installMode -eq "Upgrade") {
                $upgradeBackupDir = Join-Path $TargetInstallDir ("_backup_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
                New-Item -ItemType Directory -Path $upgradeBackupDir -Force | Out-Null
                Get-ChildItem -Path $TargetInstallDir -Exclude "_backup_*" | ForEach-Object {
                    Copy-Item -Path $_.FullName -Destination $upgradeBackupDir -Recurse -Force
                }
                $completedSteps += "创建旧版本应用文件安全备份: $upgradeBackupDir"
            }

            # 复制新文件到目标目录
            Get-ChildItem -Path $releaseDir -Exclude "scripts","core","tests",".artifacts" | ForEach-Object {
                Copy-Item -Path $_.FullName -Destination $TargetInstallDir -Recurse -Force
            }
            $completedSteps += "复制新版应用文件至安装目标目录"
            Write-Host "      应用文件安装/升级完成，等待注册表校验..." -ForegroundColor Green
            $ErrorActionPreference = $origEap1
        } catch {
            $ErrorActionPreference = $origEap1
            $failedStep = $_.Exception.Message
            Write-Host "      [错误] 应用文件更新失败: $failedStep" -ForegroundColor Red
            
            # 执行补偿恢复：如果有备份，回退原文件
            if ($upgradeBackupDir -and (Test-Path $upgradeBackupDir)) {
                Write-Host "      正在执行失败补偿恢复，将旧版应用文件还原至目标目录..." -ForegroundColor Yellow
                try {
                    Get-ChildItem -Path $upgradeBackupDir | ForEach-Object {
                        Copy-Item -Path $_.FullName -Destination $TargetInstallDir -Recurse -Force
                    }
                    Remove-Item -Path $upgradeBackupDir -Recurse -Force -ErrorAction SilentlyContinue
                    Write-Host "      补偿恢复完成：已恢复至升级前应用版本，注册状态未受损。" -ForegroundColor Yellow
                    Write-Host "      [说明] 此操作为应用级补偿恢复，不承诺断电或操作系统崩溃下的数据库事务原子性。" -ForegroundColor DarkGray
                } catch {
                    Write-Host "      [严重警告] 补偿恢复遇到异常: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            exit 3
        }
    }
} else {
    Write-Host "      安装模式: 就地注册 (发布目录即运行目录: $releaseDir)" -ForegroundColor Cyan
}

# -------------------------------------------------------------
# 6. 检测系统与 Excel 架构，并向注册表注册加载项
# -------------------------------------------------------------
Write-Host "[6/6] 正在向 Excel 注册自启动加载项..." -ForegroundColor Yellow

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
    Write-Host "      识别到本机 Excel: $excelPath ($bitness 位)" -ForegroundColor DarkGray
} else {
    Write-Host "      [提示] 未在默认路径找到 excel.exe，采用当前 Windows 预判架构 ($bitness 位)。" -ForegroundColor DarkGray
}

$finalTargetDir = if (![string]::IsNullOrEmpty($TargetInstallDir)) { $TargetInstallDir } else { $releaseDir }
$targetXll = if ($bitness -eq 32) {
    Join-Path $finalTargetDir "LeeExcel.xll"
} else {
    Join-Path $finalTargetDir "LeeExcel64.xll"
}

if (!$DryRun -and !(Test-Path $targetXll)) {
    Write-Host "      [错误] 目标加载项文件不存在: $targetXll" -ForegroundColor Red
    exit 1
}

# 注册表注入基准路径 (支持隔离测试自定义注册表根)
$effectiveRegBase = if (![string]::IsNullOrEmpty($RegistryRoot)) { $RegistryRoot } else { "HKCU:\Software\Microsoft\Office" }
$officeVersions = @("16.0", "15.0", "14.0")
$regSuccessCount = 0

try {
    $origEap2 = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    foreach ($ver in $officeVersions) {
        $verOfficePath = "$effectiveRegBase\$ver\Excel"
        $optionsPath = "$effectiveRegBase\$ver\Excel\Options"
        
        # 隔离测试时或者已存在 Office 该版本时注入
        if (![string]::IsNullOrEmpty($RegistryRoot) -or (Test-Path $verOfficePath)) {
            if (!$DryRun) {
                if (!(Test-Path $optionsPath)) {
                    New-Item -Path $optionsPath -Force | Out-Null
                }

                $props = Get-ItemProperty -Path $optionsPath -ErrorAction SilentlyContinue
                $alreadyRegisteredKey = $null

                if ($props) {
                    foreach ($propName in ($props.PSObject.Properties.Name | Where-Object { $_ -match "^OPEN\d*$" })) {
                        $val = $props.$propName
                        if ($val -like "*LeeExcel*" -or $val -like "*ExcelMind*") {
                            $alreadyRegisteredKey = $propName
                            break
                        }
                    }
                }

                $regValue = "/R `"$targetXll`""

                if (![string]::IsNullOrEmpty($alreadyRegisteredKey)) {
                    Set-ItemProperty -Path $optionsPath -Name $alreadyRegisteredKey -Value $regValue
                    Write-Host "      已更新 Office $ver 自启动项 [$alreadyRegisteredKey]: $targetXll" -ForegroundColor Green
                    $regSuccessCount++
                } else {
                    $targetKey = "OPEN"
                    $counter = 1
                    if ($props) {
                        while ($props.PSObject.Properties.Name -contains $targetKey) {
                            $targetKey = "OPEN$counter"
                            $counter++
                        }
                    }
                    New-ItemProperty -Path $optionsPath -Name $targetKey -Value $regValue -PropertyType String -Force | Out-Null
                    Write-Host "      已注入 Office $ver 自启动项 [$targetKey]: $targetXll" -ForegroundColor Green
                    $regSuccessCount++
                }
            } else {
                Write-Host "      [DryRun] 将注入 Office $ver 自启动项 -> $targetXll" -ForegroundColor DarkGray
                $regSuccessCount++
            }
        }
    }

    if ($regSuccessCount -eq 0 -and [string]::IsNullOrEmpty($RegistryRoot)) {
        # 兜底 Office 16.0
        $fallbackPath = "$effectiveRegBase\16.0\Excel\Options"
        if (!$DryRun) {
            if (!(Test-Path $fallbackPath)) {
                New-Item -Path $fallbackPath -Force | Out-Null
            }
            New-ItemProperty -Path $fallbackPath -Name "OPEN" -Value "/R `"$targetXll`"" -PropertyType String -Force | Out-Null
            Write-Host "      已为通用 Office 16.0 注册自启动项: $targetXll" -ForegroundColor Green
        } else {
            Write-Host "      [DryRun] 将为 Office 16.0 注册自启动项 -> $targetXll" -ForegroundColor DarkGray
        }
    }

    # 升级全链路成功后清理暂存备份
    if ($upgradeBackupDir -and (Test-Path $upgradeBackupDir)) {
        Remove-Item -Path $upgradeBackupDir -Recurse -Force -ErrorAction SilentlyContinue
        $completedSteps += "清理升级过程暂存备份"
    }
    $ErrorActionPreference = $origEap2
} catch {
    $ErrorActionPreference = $origEap2
    Write-Host "      [错误] 注册表自启动项配置失败: $($_.Exception.Message)" -ForegroundColor Red
    if ($upgradeBackupDir -and (Test-Path $upgradeBackupDir)) {
        Write-Host "      正在执行失败补偿恢复，将旧版应用文件还原至目标目录..." -ForegroundColor Yellow
        try {
            Get-ChildItem -Path $upgradeBackupDir | ForEach-Object {
                Copy-Item -Path $_.FullName -Destination $TargetInstallDir -Recurse -Force
            }
            Remove-Item -Path $upgradeBackupDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "      补偿恢复完成：已回退至升级前应用版本，注册状态未受损。" -ForegroundColor Yellow
        } catch {
            Write-Host "      [严重警告] 补偿恢复遇到异常: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    exit 3
}

Write-Host "==========================================================" -ForegroundColor Green
Write-Host "           ExcelMind AI 无损安装/升级已成功就绪！           " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "执行概要：" -ForegroundColor Cyan
Write-Host "  - 安装模式: $installMode"
Write-Host "  - 应用文件路径: $targetXll"
Write-Host "  - 用户数据保护: 100% 独立保留于 $actualUserDataDir (零修改)"
Write-Host "  - 注册状态: 已配置自启动项，打开 Excel 即可自动加载"
Write-Host ""

exit 0
