# scratch/test_warehouse_inventory_clean_run.ps1
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$runId = [Guid]::NewGuid().ToString("N").Substring(0, 8)
$testDir = Join-Path $env:TEMP ("LeeExcel_WarehouseClean_" + $runId)
New-Item -ItemType Directory -Path $testDir | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $prompt = "制作一个多仓库库存与警戒看板：新建一张表，录入4个仓库（北京仓、上海仓、广州仓、武汉仓）中5类商品（芯片、主板、电源、内存、硬盘）的当前库存与安全库存基准线；自动计算各仓库总库存量；自动计算各商品的全国总库存，并与安全基准线比对计算全国缺货量；红色或浅红醒目高亮低于安全基准线的缺货单元格；在顶部或右侧制作3张关键指标卡片：总库存件数、全国缺货商品数、最低库存仓库与件数。排版清晰商务，使用标准公式计算。"

    $wbPath = Join-Path $testDir "WarehouseInventory_Success.xlsx"
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($wbPath)

    $sysPrompt = @"
你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。
目标工作簿: "$($wb.Name)"。
【核心执行协议与规范】：
1. 根据用户的自然语言需求，自主决定最合适的高效实现方案（可自由使用循环、数组、公式、格式、图表、筛选及辅助过程等，不受限固定模板与行数）。
2. 代码必须是完整可编译运行的标准 VBA，语法严格遵循 VB6/VBA 规范，包裹在单个 ```vba ... ``` 代码块中，以 End Sub 正常闭合。
3. 【语法细节注意】：条件格式若使用 FormatConditions.Add，其参数名为 Formula1；提前退出使用 Exit Sub；严禁调用 MsgBox、Application.Quit 或弹出阻塞式交互确认框。
4. 【范围规范】：边框与背景色仅设置在实际业务数据区域（如 ws.Range(...)），严禁对全表 ws.Cells 批量设置边框或背景色。
5. 【结构与输出】：直接输出完整可执行的标准 VBA 代码，包裹在 ```vba ... ``` 代码块中。
"@

    $sysFile = Join-Path $testDir "sys.txt"
    $outFile = Join-Path $testDir "out.txt"
    $metaFile = Join-Path $testDir "meta.json"
    [System.IO.File]::WriteAllText($sysFile, $sysPrompt, [System.Text.Encoding]::UTF8)

    $env:THINKING_MODE = "disabled"

    Write-Host "-> 正在请求模型生成多仓库库存看板 VBA..." -ForegroundColor Cyan
    $nodeCmd = "node scratch/call_llm.cjs `"$sysFile`" `"$prompt`" `"$outFile`" `"$metaFile`""
    cmd.exe /c $nodeCmd

    $meta = (Get-Content $metaFile -Raw -Encoding UTF8) | ConvertFrom-Json
    $rawResponse = [System.IO.File]::ReadAllText($outFile, [System.Text.Encoding]::UTF8)

    Write-Host "`n-> 模型返回完成：" -ForegroundColor Green
    Write-Host "    - 是否触发范围失控拦截: $($meta.scopeAudit.hasRisk)"
    if ($meta.scopeAudit.hasRisk) {
        Write-Host "    - 命中高危片段: $($meta.scopeAudit.riskSnippet)" -ForegroundColor Red
        Write-Host "    - 重生后是否已消除: $($meta.scopeAudit.hasRiskAfterRegen -eq $false)"
    }
    Write-Host "    - 最终代码 SHA256: $($meta.versionControl.executedVersionHash)"
    Write-Host "    - Token 账目: 提示词=$($meta.tokenAccounting.promptTokens), 思考=$($meta.tokenAccounting.reasoningTokens), 正文=$($meta.tokenAccounting.contentTokens), 累计=$($meta.tokenAccounting.totalTokens)" -ForegroundColor Cyan

    $vbaCodeToRun = $meta.versionControl.executedVersionCode
    [System.IO.File]::WriteAllText((Join-Path $testDir "executed.vba"), $vbaCodeToRun, [System.Text.Encoding]::UTF8)

    Write-Host "`n-> 注入执行代码到 Excel..." -ForegroundColor Cyan
    $execReq = @{
        action = "execute_vba"
        code = $vbaCodeToRun
        prompt = $prompt
        targetWorkbookName = $wb.Name
        targetWorkbookFullName = $wbPath
        rawModelResponse = $rawResponse
    }
    $resJson = [LeeExcel.NativeBridge]::Dispatch(($execReq | ConvertTo-Json -Compress), $excel)
    $parsed = [LeeExcel.SimpleJson]::ParseFlatObject($resJson)

    Write-Host "`n-> 执行结果：" -ForegroundColor Yellow
    Write-Host "    - ok: $($parsed['ok'])" -ForegroundColor $(if ($parsed['ok'] -eq 'True') { "Green" } else { "Red" })
    Write-Host "    - phase: $($parsed['executionPhase'])"
    Write-Host "    - precheck: $($parsed['precheckStatus'])"
    Write-Host "    - summary: $($parsed['summary'])"
    Write-Host "    - elapsedMs: $($parsed['elapsedMs']) ms"

    $ws = $wb.ActiveSheet
    $usedRange = $ws.UsedRange.Address($false, $false)
    $rows = $ws.UsedRange.Rows.Count
    $cols = $ws.UsedRange.Columns.Count

    Write-Host "`n-> 读回数据矩阵 (表名: $($ws.Name), 使用区域: $usedRange, 行: $rows, 列: $cols):" -ForegroundColor Cyan
    for ($r = 1; $r -le [Math]::Min($rows, 20); $r++) {
        $rowTexts = @()
        for ($c = 1; $c -le [Math]::Min($cols, 10); $c++) {
            $txt = $ws.Cells.Item($r, $c).Text
            $rowTexts += [string]::Format("{0,-12}", $txt)
        }
        Write-Host ("    行 {0,2}: {1}" -f $r, ($rowTexts -join " | "))
    }

    $wb.Save()
    Write-Host "`n工作簿已成功保存: $wbPath" -ForegroundColor Green

} finally {
    try {
        $excel.Workbooks | ForEach-Object { $_.Close($false) }
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch { }
}
