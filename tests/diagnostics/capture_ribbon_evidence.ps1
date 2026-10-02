# ExcelMind AI: 真实 Excel 功能区加载与 UI Automation 取证脚本
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$artifactsDir = Join-Path $projectRoot ".artifacts\tests\macro_verification"
if (!(Test-Path $artifactsDir)) {
    New-Item -ItemType Directory -Path $artifactsDir -Force | Out-Null
}

$screenshotPath = Join-Path $artifactsDir "excel_ribbon_macro_management.png"
$evidenceLogPath = Join-Path $artifactsDir "ribbon_automation_evidence.txt"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  ExcelMind AI: 真实 Excel 功能区 UI Automation 自动化取证 " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. 查找 Excel.exe 与 LeeExcel64.xll
$excelPath = (Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" -ErrorAction SilentlyContinue).'(default)'
if (![string]::IsNullOrEmpty($excelPath) -and !(Test-Path $excelPath)) {
    $excelPath = "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
}
$xllPath = Join-Path $projectRoot "bin\LeeExcel64.xll"

if (!(Test-Path $xllPath)) {
    Write-Host "[ERROR] 未找到插件 bin\LeeExcel64.xll" -ForegroundColor Red
    exit 1
}

Write-Host "Excel 路径: $excelPath"
Write-Host "插件路径: $xllPath"

# 解除文件锁定
Unblock-File -Path $xllPath -ErrorAction SilentlyContinue

# 2. 启动 Excel 并挂载加载项
Write-Host "正在启动 Excel 并挂载加载项..." -ForegroundColor Cyan
$proc = Start-Process -FilePath $excelPath -ArgumentList "`"$xllPath`"" -PassThru

# 等待 Excel 主窗口出现
$timeout = 25
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$mainWindowHandle = [IntPtr]::Zero

while ($sw.Elapsed.TotalSeconds -lt $timeout) {
    $proc.Refresh()
    if ($proc.MainWindowHandle -ne [IntPtr]::Zero) {
        $mainWindowHandle = $proc.MainWindowHandle
        break
    }
    Start-Sleep -Milliseconds 500
}

Write-Host "Excel 主窗口句柄: $mainWindowHandle (耗时 $($sw.Elapsed.TotalSeconds)s)"

# 等待加载项初始化完成
Start-Sleep -Seconds 4

# 3. 使用 UI Automation 查找并激活 ExcelMind AI 选项卡
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$root = [System.Windows.Automation.AutomationElement]::RootElement
$procCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $proc.Id)
$excelWindow = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $procCond)

$evidenceLog = New-Object System.Text.StringBuilder
[void]$evidenceLog.AppendLine("=== ExcelMind AI 功能区 UI Automation 探测证据 ===")
[void]$evidenceLog.AppendLine("时间: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$evidenceLog.AppendLine("进程 PID: $($proc.Id)")

$foundTab = $false
$foundImportBtn = $false
$foundMyMacrosBtn = $false

if ($excelWindow) {
    # 查找并点击 "ExcelMind AI" TabItem
    $tabCond = New-Object System.Windows.Automation.AndCondition(
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::TabItem)),
        (New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::NameProperty, "ExcelMind AI"))
    )
    $tabElement = $excelWindow.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $tabCond)

    if ($tabElement) {
        $foundTab = $true
        [void]$evidenceLog.AppendLine("[CONFIRMED] 成功定位到 ExcelMind AI 选项卡 (TabItem)")
        Write-Host "[OK] 成功定位到 ExcelMind AI 选项卡" -ForegroundColor Green

        try {
            $selectPattern = $tabElement.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
            $selectPattern.Select()
            Write-Host "[OK] 已激活 ExcelMind AI 选项卡" -ForegroundColor Green
        } catch {
            Write-Host "[WARN] 激活选项卡异常: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        Start-Sleep -Milliseconds 1200

        # 枚举选项卡下的所有按钮
        $btnCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
        $buttons = $excelWindow.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)

        [void]$evidenceLog.AppendLine("探测到的按钮清单:")
        foreach ($btn in $buttons) {
            $btnName = $btn.Current.Name
            if (![string]::IsNullOrEmpty($btnName)) {
                [void]$evidenceLog.AppendLine("  - 按钮: $btnName")
                if ($btnName -match "导入宏") { $foundImportBtn = $true }
                if ($btnName -match "我的宏") { $foundMyMacrosBtn = $true }
            }
        }
    } else {
        [void]$evidenceLog.AppendLine("[WARNING] 未直接通过 AutomationElement 定位到 TabItem，尝试键盘快捷方式激活")
    }

    # 4. 截取 Excel 窗口图像
    try {
        $rect = $excelWindow.Current.BoundingRectangle
        if ($rect.Width -gt 200 -and $rect.Height -gt 200) {
            $bmp = New-Object System.Drawing.Bitmap([int]$rect.Width, [int]$rect.Height)
            $gfx = [System.Drawing.Graphics]::FromImage($bmp)
            $gfx.CopyFromScreen([int]$rect.X, [int]$rect.Y, 0, 0, New-Object System.Drawing.Size([int]$rect.Width, [int]$rect.Height))
            $bmp.Save($screenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
            $gfx.Dispose()
            $bmp.Dispose()
            Write-Host "[OK] Excel 功能区截图已保存: $screenshotPath" -ForegroundColor Green
            [void]$evidenceLog.AppendLine("[CONFIRMED] 窗口截图已成功生成: $screenshotPath")
        }
    } catch {
        Write-Host "[WARN] 截图捕获异常: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

[void]$evidenceLog.AppendLine("关键按钮验证结果:")
[void]$evidenceLog.AppendLine("  - 发现 ExcelMind AI 选项卡: $foundTab")
[void]$evidenceLog.AppendLine("  - 发现【导入宏】按钮: $foundImportBtn")
[void]$evidenceLog.AppendLine("  - 发现【我的宏】按钮: $foundMyMacrosBtn")

[System.IO.File]::WriteAllText($evidenceLogPath, $evidenceLog.ToString(), [System.Text.Encoding]::UTF8)
Write-Host "证据日志已保存: $evidenceLogPath"

# 优雅关闭测试 Excel 进程
try {
    Write-Host "正在关闭测试 Excel 进程..." -ForegroundColor Cyan
    $proc.CloseMainWindow() | Out-Null
    Start-Sleep -Seconds 2
    if (!$proc.HasExited) {
        $proc.Kill()
    }
} catch { }

Write-Host "自动化取证完毕！" -ForegroundColor Green
