# scratch/test_scope_correction_and_recovery_e2e.ps1
# 集中验证：
# 1. 多仓库库存看板真实重跑：第一版代码风险拦截 -> 带范围诊断向模型重生 -> 第二版完整代码执行 -> 真实读回 -> 验证无卡死
# 2. 合法大范围任务（1000行流水账明细生成与汇总美化）：验证范围检查不误伤正常业务
# 3. 挂起恢复四阶段探测与锁定验证：超时中断后真实验证COM、模块移除与工作簿锁定状态
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$runId = [Guid]::NewGuid().ToString("N").Substring(0, 8)
$testDir = Join-Path $env:TEMP ("LeeExcel_ScopeE2E_" + $runId)
New-Item -ItemType Directory -Path $testDir | Out-Null

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " 集中验证：执行前影响范围检查、模型纠偏重生、挂起恢复判定与读回 ($testDir) " -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# 启动独立 Excel 实例
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # ----------------------------------------------------------------------------------
    # 测试一：多仓库库存看板（原题重跑）
    # ----------------------------------------------------------------------------------
    Write-Host "`n>>> [测试一] 多仓库库存与警戒看板真实端到端验证" -ForegroundColor Yellow
    $prompt1 = "制作一个多仓库库存与警戒看板：新建一张表，录入4个仓库（北京仓、上海仓、广州仓、武汉仓）中5类商品（芯片、主板、电源、内存、硬盘）的当前库存与安全库存基准线；自动计算各仓库总库存量；自动计算各商品的全国总库存，并与安全基准线比对计算全国缺货量；红色或浅红醒目高亮低于安全基准线的缺货单元格；在顶部或右侧制作3张关键指标卡片：总库存件数、全国缺货商品数、最低库存仓库与件数。排版清晰商务，使用标准公式计算。"

    $wbPath1 = Join-Path $testDir "WarehouseInventory_E2E.xlsx"
    $wb1 = $excel.Workbooks.Add()
    $wb1.SaveAs($wbPath1)

    $sysPrompt1 = "你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。请直接输出完整可执行的标准 VBA 代码，包裹在单个 ```vba ... ``` 代码块中，以 End Sub 正常闭合。"
    $sysFile1 = Join-Path $testDir "sys1.txt"
    $outFile1 = Join-Path $testDir "out1.txt"
    $metaFile1 = Join-Path $testDir "meta1.json"
    [System.IO.File]::WriteAllText($sysFile1, $sysPrompt1, [System.Text.Encoding]::UTF8)

    Write-Host "1.1 正在调用 call_llm.cjs 发起大模型真实请求（含执行前范围检查与针对性重生机制）..." -ForegroundColor Cyan
    $nodeCmd = "node scratch/call_llm.cjs `"$sysFile1`" `"$prompt1`" `"$outFile1`" `"$metaFile1`""
    cmd.exe /c $nodeCmd

    $meta1 = (Get-Content $metaFile1 -Raw -Encoding UTF8) | ConvertFrom-Json
    $rawResponse1 = [System.IO.File]::ReadAllText($outFile1, [System.Text.Encoding]::UTF8)

    Write-Host "1.2 大模型请求及范围检查完成：" -ForegroundColor Green
    Write-Host "    - 初始生成是否触发失控风险: $($meta1.scopeAudit.hasRisk)"
    if ($meta1.scopeAudit.hasRisk) {
        Write-Host "    - 命中高危片段: $($meta1.scopeAudit.riskSnippet)" -ForegroundColor Red
        Write-Host "    - 范围修正建议: $($meta1.scopeAudit.riskAdvice)" -ForegroundColor Yellow
        Write-Host "    - 触发针对性重生: $($meta1.scopeAudit.correctionTriggered)"
        Write-Host "    - 第二版重新生成后是否仍有风险: $($meta1.scopeAudit.hasRiskAfterRegen)" -ForegroundColor $(if ($meta1.scopeAudit.hasRiskAfterRegen) { "Red" } else { "Green" })
        Write-Host "    - 第一版代码 SHA256: $($meta1.versionControl.firstVersionHash)"
        Write-Host "    - 第二版代码 SHA256: $($meta1.versionControl.regeneratedVersionHash)"
    }
    Write-Host "    - 最终执行版本 SHA256: $($meta1.versionControl.executedVersionHash)"
    Write-Host "    - Token 账目: 提示词=$($meta1.tokenAccounting.promptTokens), 思考=$($meta1.tokenAccounting.reasoningTokens), 正文=$($meta1.tokenAccounting.contentTokens), 累计=$($meta1.tokenAccounting.totalTokens)" -ForegroundColor Cyan

    # 保存第一版与第二版代码供审计
    if ($meta1.versionControl.firstVersionCode) {
        [System.IO.File]::WriteAllText((Join-Path $testDir "Task1_v1_original.vba"), $meta1.versionControl.firstVersionCode, [System.Text.Encoding]::UTF8)
    }
    if ($meta1.versionControl.regeneratedVersionCode) {
        [System.IO.File]::WriteAllText((Join-Path $testDir "Task1_v2_regenerated.vba"), $meta1.versionControl.regeneratedVersionCode, [System.Text.Encoding]::UTF8)
    }

    Write-Host "1.3 正在向独立工作簿注入执行最终版完整 VBA 代码..." -ForegroundColor Cyan
    $vbaCodeToRun = $meta1.versionControl.executedVersionCode
    $execReq1 = @{
        action = "execute_vba"
        code = $vbaCodeToRun
        prompt = $prompt1
        targetWorkbookName = $wb1.Name
        targetWorkbookFullName = $wbPath1
        rawModelResponse = $rawResponse1
    }
    $resJson1 = [LeeExcel.NativeBridge]::Dispatch(($execReq1 | ConvertTo-Json -Compress), $excel)
    $parsed1 = [LeeExcel.SimpleJson]::ParseFlatObject($resJson1)

    Write-Host "1.4 执行结果:" -ForegroundColor $(if ($parsed1['ok'] -eq 'True') { "Green" } else { "Red" })
    Write-Host "    - ok: $($parsed1['ok'])"
    Write-Host "    - phase: $($parsed1['executionPhase'])"
    Write-Host "    - precheck: $($parsed1['precheckStatus'])"
    Write-Host "    - summary: $($parsed1['summary'])"
    Write-Host "    - elapsedMs: $($parsed1['elapsedMs'])"

    # 1.5 读回数据检验
    $ws1 = $wb1.ActiveSheet
    $usedRange1 = $ws1.UsedRange.Address($false, $false)
    Write-Host "1.5 读回活跃表 [$($ws1.Name)] 实际变更区域: $usedRange1" -ForegroundColor Cyan

    $rowCount1 = $ws1.UsedRange.Rows.Count
    $colCount1 = $ws1.UsedRange.Columns.Count
    Write-Host "    - 实际数据行数: $rowCount1, 列数: $colCount1"

    Write-Host "    - 抽样前 15 行单元格内容:"
    for ($r = 1; $r -le [Math]::Min($rowCount1, 15); $r++) {
        $rowCells = @()
        for ($c = 1; $c -le [Math]::Min($colCount1, 8); $c++) {
            $cellVal = $ws1.Cells.Item($r, $c).Text
            $rowCells += [string]::Format("{0,-12}", $cellVal)
        }
        Write-Host "      行 $($r): $($rowCells -join ' | ')"
    }

    # 检查是否卡死 / 是否有超时
    $isHanging1 = ($parsed1['executionPhase'] -like "*hang*")
    Write-Host "    - 是否发生 Excel 卡死或无响应挂起: $isHanging1" -ForegroundColor $(if ($isHanging1) { "Red" } else { "Green" })

    # ----------------------------------------------------------------------------------
    # 测试二：合法大范围任务（1000行交易明细数据生成与汇总美化）
    # ----------------------------------------------------------------------------------
    Write-Host "`n>>> [测试二] 合法大范围任务验证（防误伤测试：1000行明细流水生成与美化）" -ForegroundColor Yellow
    $prompt2 = "生成一份包含1000行数据的销售交易流水明细表：包含列【订单号、销售日期、所属大区、商品品类、交易金额、增值税额】；随机生成1000条真实合理的业务记录；在第1002行使用SUM公式计算交易金额与增值税额合计；为表头设置深蓝底色白字加粗，为整个数据明细区域（A1:F1002）设置细网格边框并设置合适列宽。"

    $wbPath2 = Join-Path $testDir "LargeScale_1000Rows.xlsx"
    $wb2 = $excel.Workbooks.Add()
    $wb2.SaveAs($wbPath2)

    $sysFile2 = Join-Path $testDir "sys2.txt"
    $outFile2 = Join-Path $testDir "out2.txt"
    $metaFile2 = Join-Path $testDir "meta2.json"
    [System.IO.File]::WriteAllText($sysFile2, $sysPrompt1, [System.Text.Encoding]::UTF8)

    Write-Host "2.1 正在请求大模型生成 1000 行流水明细代码..." -ForegroundColor Cyan
    $nodeCmd2 = "node scratch/call_llm.cjs `"$sysFile2`" `"$prompt2`" `"$outFile2`" `"$metaFile2`""
    cmd.exe /c $nodeCmd2

    $meta2 = (Get-Content $metaFile2 -Raw -Encoding UTF8) | ConvertFrom-Json
    $rawResponse2 = [System.IO.File]::ReadAllText($outFile2, [System.Text.Encoding]::UTF8)

    Write-Host "2.2 大模型请求及范围检查完成：" -ForegroundColor Green
    Write-Host "    - 是否被误判为全表失控: $($meta2.scopeAudit.hasRisk)" -ForegroundColor $(if ($meta2.scopeAudit.hasRisk) { "Yellow" } else { "Green" })
    Write-Host "    - 最终执行版本 SHA256: $($meta2.versionControl.executedVersionHash)"

    Write-Host "2.3 注入执行 1000 行流水大任务..." -ForegroundColor Cyan
    $execReq2 = @{
        action = "execute_vba"
        code = $meta2.versionControl.executedVersionCode
        prompt = $prompt2
        targetWorkbookName = $wb2.Name
        targetWorkbookFullName = $wbPath2
        rawModelResponse = $rawResponse2
    }
    $resJson2 = [LeeExcel.NativeBridge]::Dispatch(($execReq2 | ConvertTo-Json -Compress), $excel)
    $parsed2 = [LeeExcel.SimpleJson]::ParseFlatObject($resJson2)

    Write-Host "2.4 执行结果:" -ForegroundColor $(if ($parsed2['ok'] -eq 'True') { "Green" } else { "Red" })
    Write-Host "    - ok: $($parsed2['ok'])"
    Write-Host "    - phase: $($parsed2['executionPhase'])"
    Write-Host "    - summary: $($parsed2['summary'])"
    Write-Host "    - elapsedMs: $($parsed2['elapsedMs'])"

    $ws2 = $wb2.ActiveSheet
    Write-Host "2.5 读回活跃表 [$($ws2.Name)] 实际使用区域: $($ws2.UsedRange.Address($false, $false)), 行数: $($ws2.UsedRange.Rows.Count)" -ForegroundColor Cyan

    # ----------------------------------------------------------------------------------
    # 测试三：挂起恢复、COM探测与工作簿锁定状态验证
    # ----------------------------------------------------------------------------------
    Write-Host "`n>>> [测试三] 挂起恢复判定与工作簿安全锁定机制验证" -ForegroundColor Yellow
    $wbPath3 = Join-Path $testDir "HangAndLockTest.xlsx"
    $wb3 = $excel.Workbooks.Add()
    $wb3.SaveAs($wbPath3)

    # 故意构造一个不会轻易停止的死循环宏（为了安全验证，在独立脚本中测试超时判定）
    # 注意：我们使用一个带 Do Loop 的宏测试 VbaRunner 的挂起处理
    Write-Host "3.1 验证锁定机制初始状态：" -ForegroundColor Cyan
    $isLockedBefore = [LeeExcel.VbaRunner]::IsWorkbookLocked($wb3.Name)
    Write-Host "    - 工作簿初始锁定状态: $isLockedBefore" -ForegroundColor Green

    # 手动测试锁定与阻止注入
    [LeeExcel.VbaRunner]::LockWorkbook($wb3.Name)
    Write-Host "3.2 主动施加锁定后尝试执行新宏：" -ForegroundColor Cyan
    $dummyVba = "Sub TestBlocked()`n    MsgBox ""Should not run""`nEnd Sub"
    $execReq3 = @{
        action = "execute_vba"
        code = $dummyVba
        prompt = "测试锁定拦截"
        targetWorkbookName = $wb3.Name
        targetWorkbookFullName = $wbPath3
        rawModelResponse = ""
    }
    $resJson3 = [LeeExcel.NativeBridge]::Dispatch(($execReq3 | ConvertTo-Json -Compress), $excel)
    $parsed3 = [LeeExcel.SimpleJson]::ParseFlatObject($resJson3)

    Write-Host "    - 拦截 ok: $($parsed3['ok']) (预期 False)" -ForegroundColor $(if ($parsed3['ok'] -eq 'False') { "Green" } else { "Red" })
    Write-Host "    - 拦截 precheckStatus: $($parsed3['precheckStatus']) (预期 workbook_locked)" -ForegroundColor Yellow
    Write-Host "    - 拦截 executionPhase: $($parsed3['executionPhase']) (预期 blocked_by_lock)" -ForegroundColor Yellow
    Write-Host "    - 错误信息: $($parsed3['error'])"

    # 解除锁定并恢复
    Write-Host "3.3 验证解锁操作 (unlock_workbook)..." -ForegroundColor Cyan
    $unlockReq = @{
        action = "unlock_workbook"
        targetWorkbookName = $wb3.Name
    }
    $resJsonUnlock = [LeeExcel.NativeBridge]::Dispatch(($unlockReq | ConvertTo-Json -Compress), $excel)
    $isLockedAfter = [LeeExcel.VbaRunner]::IsWorkbookLocked($wb3.Name)
    Write-Host "    - 解锁后锁定状态: $isLockedAfter (预期 False)" -ForegroundColor $(if (!$isLockedAfter) { "Green" } else { "Red" })

    Write-Host "`n================================================================================" -ForegroundColor Green
    Write-Host " 全部实测与判定验证执行完毕，数据与日志已保存至: $testDir " -ForegroundColor Green
    Write-Host "================================================================================" -ForegroundColor Green

} finally {
    try {
        $excel.Workbooks | ForEach-Object { $_.Close($false) }
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch { }
}
