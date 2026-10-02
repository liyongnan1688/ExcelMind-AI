# ExcelMind AI 卸载与清理脚本
# 适用：Office 2010 / 2013 / 2016 / 2019 / 2021 / Microsoft 365

$ErrorActionPreference = "Continue"

# 统一控制台输出编码为 UTF-8，防止在 Windows PowerShell 5.1 下输出中文字符乱码
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "           ExcelMind AI 原生插件 - 卸载与清理向导           " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$officeVersions = @("16.0", "15.0", "14.0")
$removedTotal = 0

foreach ($ver in $officeVersions) {
    $optionsPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Options"
    if (Test-Path $optionsPath) {
        $props = Get-ItemProperty -Path $optionsPath -ErrorAction SilentlyContinue
        $openKeys = $props.PSObject.Properties | Where-Object { $_.Name -match "^OPEN\d*$" }
        
        $cleanedItems = @()
        $foundTarget = $false

        foreach ($p in $openKeys) {
            $val = [string]$p.Value
            if ($val -like "*LeeExcel*" -or $val -like "*ExcelMind*") {
                Remove-ItemProperty -Path $optionsPath -Name $p.Name -ErrorAction SilentlyContinue
                Write-Host "      已从 Office $ver 移除加载项: $($p.Name) -> $val" -ForegroundColor Yellow
                $removedTotal++
                $foundTarget = $true
            } else {
                $cleanedItems += [PSCustomObject]@{
                    OriginalName = $p.Name
                    Value = $val
                }
            }
        }

        # 如果移除了项，对剩余的 OPEN 项进行紧凑重排，避免序号断号
        if ($foundTarget) {
            # 先清除所有剩余的旧 OPEN 项
            foreach ($item in $cleanedItems) {
                Remove-ItemProperty -Path $optionsPath -Name $item.OriginalName -ErrorAction SilentlyContinue
            }

            # 重新按顺序紧凑写入
            for ($i = 0; $i -lt $cleanedItems.Count; $i++) {
                $newName = if ($i -eq 0) { "OPEN" } else { "OPEN$i" }
                New-ItemProperty -Path $optionsPath -Name $newName -Value $cleanedItems[$i].Value -PropertyType String -Force | Out-Null
            }
            Write-Host "      已重构 Office $ver 剩余加载项序列索引。" -ForegroundColor DarkGray
        }
    }
}

Write-Host "==========================================================" -ForegroundColor Green
if ($removedTotal -gt 0) {
    Write-Host "       ExcelMind AI 加载项已成功从系统注册表中完全注销！     " -ForegroundColor Green
} else {
    Write-Host "       未在系统中检测到残留的 ExcelMind AI 注册项。         " -ForegroundColor Green
}
Write-Host "==========================================================" -ForegroundColor Green
Write-Host "说明：卸载仅注销自启动项，不会影响您已保存的 Excel 工作簿。" -ForegroundColor Cyan
Write-Host ""
