# scratch/test_warehouse_inventory_scope_cycle.ps1
# 真实重跑多仓库库存看板任务：
# 记录第一版代码的风险发现（ws.Cells.Borders） -> 模型重新生成第二版完整VBA -> 真正执行 -> 读回库存与卡片 -> 检验无挂起
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$runId = [Guid]::NewGuid().ToString("N").Substring(0, 8)
$testDir = Join-Path $env:TEMP ("LeeExcel_WarehouseCycle_" + $runId)
New-Item -ItemType Directory -Path $testDir | Out-Null

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " 多仓库库存看板真实重跑：第一版失控拦截 -> 纠偏重生 -> 真实执行 -> 数据读回 " -ForegroundColor Cyan
Write-Host " 测试工作目录: $testDir" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# 启动隔离 Excel 进程
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $prompt = "制作一个多仓库库存与警戒看板：新建一张表，录入4个仓库（北京仓、上海仓、广州仓、武汉仓）中5类商品（芯片、主板、电源、内存、硬盘）的当前库存与安全库存基准线；自动计算各仓库总库存量；自动计算各商品的全国总库存，并与安全基准线比对计算全国缺货量；红色或浅红醒目高亮低于安全基准线的缺货单元格；在顶部或右侧制作3张关键指标卡片：总库存件数、全国缺货商品数、最低库存仓库与件数。排版清晰商务，使用标准公式计算。"

    $wbPath = Join-Path $testDir "WarehouseInventory_RealRun.xlsx"
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($wbPath)

    # 包含产品标准系统提示词（明确规范标准 VBA、禁止阻塞式 MsgBox）
    $sysPrompt = @"
你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。
目标工作簿: "$($wb.Name)"。
【核心执行协议与规范】：
1. 根据用户的自然语言需求，自主决定最合适的高效实现方案（可自由使用循环、数组、公式、格式、图表、筛选及辅助过程等，不受限固定模板与行数）。
2. 代码必须是完整可编译运行的标准 VBA，语法严格遵循 VB6/VBA 规范，包裹在单个 ```vba ... ``` 代码块中，以 End Sub 正常闭合。
3. 【安全约束】：严禁调用 MsgBox、Application.Quit 或弹出阻塞式交互确认框。
4. 【结构与输出】：直接输出完整可执行的标准 VBA 代码，包裹在 ```vba ... ``` 代码块中。
"@

    $sysFile = Join-Path $testDir "sys.txt"
    $outFile = Join-Path $testDir "out.txt"
    $metaFile = Join-Path $testDir "meta.json"
    [System.IO.File]::WriteAllText($sysFile, $sysPrompt, [System.Text.Encoding]::UTF8)

    # 设置思考模式为 disabled 或 budget 避免深度思考打满 16384 token 导致正文截断
    $env:THINKING_MODE = "disabled"

    Write-Host "1. 调用 call_llm.cjs 发起大模型生成（带执行前影响范围检查与自动纠偏重生）..." -ForegroundColor Cyan
    $nodeCmd = "node scratch/call_llm.cjs `"$sysFile`" `"$prompt`" `"$outFile`" `"$metaFile`""
    cmd.exe /c $nodeCmd

    $meta = (Get-Content $metaFile -Raw -Encoding UTF8) | ConvertFrom-Json
    $rawResponse = [System.IO.File]::ReadAllText($outFile, [System.Text.Encoding]::UTF8)

    Write-Host "`n2. 检查大模型版本与影响范围审计结果：" -ForegroundColor Yellow
    Write-Host "    - 初始第一版是否发现失控风险: $($meta.scopeAudit.hasRisk)"
    if ($meta.scopeAudit.hasRisk) {
        Write-Host "    - 命中失控操作: $($meta.scopeAudit.riskSnippet)" -ForegroundColor Red
        Write-Host "    - 诊断纠偏建议: $($meta.scopeAudit.riskAdvice)" -ForegroundColor Yellow
        Write-Host "    - 触发针对性重生: $($meta.scopeAudit.correctionTriggered)"
        Write-Host "    - 第二版重新生成后是否仍有风险: $($meta.scopeAudit.hasRiskAfterRegen)" -ForegroundColor $(if ($meta.scopeAudit.hasRiskAfterRegen) { "Red" } else { "Green" })
        Write-Host "    - 第一版代码 SHA256: $($meta.versionControl.firstVersionHash)"
        Write-Host "    - 第二版代码 SHA256: $($meta.versionControl.regeneratedVersionHash)"
    }
    Write-Host "    - 最终执行版代码 SHA256: $($meta.versionControl.executedVersionHash)"
    Write-Host "    - Token 账目: 提示词=$($meta.tokenAccounting.promptTokens), 思考=$($meta.tokenAccounting.reasoningTokens), 正文=$($meta.tokenAccounting.contentTokens), 累计=$($meta.tokenAccounting.totalTokens)" -ForegroundColor Cyan

    # 保存各版本独立源码文件供人工审计
    if ($meta.versionControl.firstVersionCode) {
        $v1Path = Join-Path $testDir "Version1_Original_RiskFound.vba"
        [System.IO.File]::WriteAllText($v1Path, $meta.versionControl.firstVersionCode, [System.Text.Encoding]::UTF8)
        Write-Host "    -> 第一版源码已保存: $v1Path"
    }
    if ($meta.versionControl.regeneratedVersionCode) {
        $v2Path = Join-Path $testDir "Version2_Regenerated_Safe.vba"
        [System.IO.File]::WriteAllText($v2Path, $meta.versionControl.regeneratedVersionCode, [System.Text.Encoding]::UTF8)
        Write-Host "    -> 第二版源码已保存: $v2Path"
    }

    # 3. 注入执行
    Write-Host "`n3. 向新独立工作簿 [$($wb.Name)] 注入并执行最终版本 VBA 代码..." -ForegroundColor Cyan
    $vbaCodeToRun = $meta.versionControl.executedVersionCode
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

    Write-Host "`n4. Excel 执行结果：" -ForegroundColor Yellow
    Write-Host "    - ok: $($parsed['ok'])" -ForegroundColor $(if ($parsed['ok'] -eq 'True') { "Green" } else { "Red" })
    Write-Host "    - executionPhase: $($parsed['executionPhase'])"
    Write-Host "    - precheckStatus: $($parsed['precheckStatus'])"
    Write-Host "    - summary: $($parsed['summary'])"
    Write-Host "    - elapsedMs: $($parsed['elapsedMs']) ms"
    Write-Host "    - 是否发生 Excel 卡死或无响应挂起: $(if ($parsed['executionPhase'] -like '*hang*') { 'True (发生挂起)' } else { 'False (正常响应)' })" -ForegroundColor $(if ($parsed['executionPhase'] -like '*hang*') { "Red" } else { "Green" })

    # 4. 读回数据并验证业务计算
    $ws = $wb.ActiveSheet
    $usedRange = $ws.UsedRange.Address($false, $false)
    $rows = $ws.UsedRange.Rows.Count
    $cols = $ws.UsedRange.Columns.Count

    Write-Host "`n5. Excel 实际读回矩阵验证：" -ForegroundColor Yellow
    Write-Host "    - 活跃工作表名称: $($ws.Name)"
    Write-Host "    - 实际使用区域: $usedRange (共 $rows 行 $cols 列)"

    Write-Host "`n    【前 20 行数据与格式展示】:"
    for ($r = 1; $r -le [Math]::Min($rows, 20); $r++) {
        $rowTexts = @()
        for ($c = 1; $c -le [Math]::Min($cols, 10); $c++) {
            $txt = $ws.Cells.Item($r, $c).Text
            $color = $ws.Cells.Item($r, $c).Interior.Color
            $isRed = ($color -eq 255 -or $color -eq 13421823 -or $color -eq 13408767 -or $color -eq 8421631) # 常见红/浅红
            if ($isRed -and $txt) {
                $rowTexts += [string]::Format("[红:{0}]", $txt)
            } else {
                $rowTexts += [string]::Format("{0}", $txt)
            }
        }
        Write-Host ("    行 {0,2}: {1}" -f $r, ($rowTexts -join " | "))
    }

    # 保存最终工作簿
    $wb.Save()
    Write-Host "`n测试工作簿已保存: $wbPath" -ForegroundColor Green

} finally {
    try {
        $excel.Workbooks | ForEach-Object { $_.Close($false) }
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch { }
}
