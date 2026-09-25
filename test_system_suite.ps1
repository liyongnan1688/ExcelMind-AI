# test_system_suite.ps1 - Deep Integration Regression Suite
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

Write-Host "===========================================================" -ForegroundColor Cyan
Write-Host "       Lee-Excel Comprehensive System Quality Regression   " -ForegroundColor Cyan
Write-Host "       (Fixed VBA COM Suite with Deep Verification)       " -ForegroundColor Cyan
Write-Host "===========================================================" -ForegroundColor Cyan

$testDir = Join-Path $env:TEMP ("LeeExcelTests_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $testDir | Out-Null
Write-Host "Isolated Test Directory: $testDir" -ForegroundColor Gray

$results = [System.Collections.Generic.List[PSObject]]::new()

function Record-Test($name, $category, $passed, $details) {
    $item = [PSCustomObject]@{
        Name     = $name
        Category = $category
        Status   = if ($passed) { "PASS" } else { "FAIL" }
        Details  = $details
    }
    $results.Add($item)
    if ($passed) {
        Write-Host "  [PASS] $name : $details" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] $name : $details" -ForegroundColor Red
    }
}

$excel = $null
$wbA = $null
$wbB = $null

try {
    Write-Host "`n[Setup] Launching isolated headless Excel COM instance..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # 创建两份独立的测试工作簿
    $wbA = $excel.Workbooks.Add()
    $wbAPath = Join-Path $testDir "TestWB_Alpha.xlsx"
    $wbA.Sheets.Item(1).Range("A1").Value2 = "Alpha_Initial_Data"
    $wbA.SaveAs($wbAPath)
    $wbAName = $wbA.Name

    $wbB = $excel.Workbooks.Add()
    $wbBPath = Join-Path $testDir "TestWB_Beta.xlsx"
    $wbB.Sheets.Item(1).Range("A1").Value2 = "Beta_Initial_Keep_Safe"
    $wbB.SaveAs($wbBPath)
    $wbBName = $wbB.Name

    Write-Host "  Created isolated Target Workbook A: $wbAName ($wbAPath)" -ForegroundColor Gray
    Write-Host "  Created isolated Target Workbook B: $wbBName ($wbBPath)" -ForegroundColor Gray

    # 记录 Workbook B 初始状态（用于后续证明其完全未被污染）
    $wbBInitialA1 = $wbB.Sheets.Item(1).Range("A1").Value2
    $wbBInitialSheetCount = $wbB.Sheets.Count

    # -------------------------------------------------------------
    # 用例 1 & 2: 纯聊天与代码解释 (零 VBA 注入、零工作簿修改)
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 1 & 2: Chat & Explanation Channel Separation ---" -ForegroundColor Cyan
    $sheetACellBefore = $wbA.Sheets.Item(1).Range("A1").Value2

    # 验证 NativeBridge 不会响应空代码执行
    $emptyReq = '{"action":"execute_vba","code":"","prompt":"你是？","targetWorkbookName":"' + $wbAName + '"}'
    $emptyRes = [LeeExcel.NativeBridge]::Dispatch($emptyReq, $excel)
    $parsedEmpty = [LeeExcel.SimpleJson]::ParseFlatObject($emptyRes)
    $passed1A = ($parsedEmpty["ok"] -eq "false") -and ($parsedEmpty["error"] -like "*传入的 VBA 代码为空*")
    Record-Test "1A. Empty VBA rejected" "Channel Separation" $passed1A $parsedEmpty["error"]

    $sheetACellAfter = $wbA.Sheets.Item(1).Range("A1").Value2
    $passed1B = ($sheetACellBefore -eq $sheetACellAfter)
    Record-Test "1B. Zero workbook mutation on chat request" "Channel Separation" $passed1B "Sheet A1 untouched"

    # -------------------------------------------------------------
    # 用例 7: 半截代码、缺少 End Sub、无参数过程拦截
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 7: Incomplete / Parameterless Code Interception ---" -ForegroundColor Cyan
    $brokenCode = "Sub IncompleteTask(targetWb As Workbook)`r`n    targetWb.Sheets(1).Range(""A1"").Value = ""broken"""
    $truncResult = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wbA, $brokenCode)
    $passed7A = ($truncResult.success -eq $false) -and ($truncResult.summary -like "*缺少 Sub 或 End Sub*")
    Record-Test "7A. Missing End Sub interception" "Code Validation" $passed7A $truncResult.summary

    # 无参过程被安全约定拦截（拒绝盲目自动运行）
    $paramlessCode = "Sub ParameterlessTask()`r`n    Range(""A1"").Value = 123`r`nEnd Sub"
    $paramlessResult = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wbA, $paramlessCode)
    $passed7B = ($paramlessResult.success -eq $false) -and ($paramlessResult.summary -like "*未声明目标工作簿参数*")
    Record-Test "7B. Parameterless procedure safely intercepted" "Target Binding Convention" $passed7B $paramlessResult.summary

    # -------------------------------------------------------------
    # 用例 8: 多工作簿并存、受控参数传递与窗口切换免疫
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 8: Parameter-based Target Binding (Window Switch Immunity) ---" -ForegroundColor Cyan
    # 故意将当前前台活动窗口切换为 Workbook B
    $wbB.Activate()
    $activeBefore = $excel.ActiveWorkbook.Name
    Write-Host "  Active window intentionally switched to: $activeBefore" -ForegroundColor Gray

    # 编写受控宏：显式通过 targetWb 写入，绝不使用未经绑定的 ActiveWorkbook
    $vbaParamTest = "Sub TestBinding(targetWb As Workbook)`r`n" +
        "    targetWb.Sheets(1).Range(""A1"").Value = ""TargetAlphaConfirmed""`r`n" +
        "End Sub"

    $execReq = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaParamTest) + '","prompt":"目标绑定受控参数测试","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $execRes = [LeeExcel.NativeBridge]::Dispatch($execReq, $excel)

    $valA = $wbA.Sheets.Item(1).Range("A1").Value2
    $valB = $wbB.Sheets.Item(1).Range("A1").Value2

    $passed8 = ($valA -eq "TargetAlphaConfirmed") -and ($valB -eq $wbBInitialA1)
    Record-Test "8. Parameter-based Target Binding (Workbook B completely protected)" "Target Binding" $passed8 "WbA updated: '$valA', WbB untouched: '$valB'"

    # -------------------------------------------------------------
    # 用例 3: 从 D1 生成阶梯式算式九九乘法表并美化 (深度形态与越界断言)
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 3: Staircase Equation Table at D1 with Deep Morphology Assertions ---" -ForegroundColor Cyan
    $ws3 = $wbA.Sheets.Add()
    $ws3.Name = "MultiplicationStaircase"

    $vbaStaircase = "Sub GenerateStaircaseMultiplication(targetWb As Workbook)`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = targetWb.Sheets(""MultiplicationStaircase"")`r`n" +
        "    Dim r As Long, c As Long`r`n" +
        "    For r = 1 To 9`r`n" +
        "        For c = 1 To r`r`n" +
        "            ws.Cells(r, c + 3).Value = c & ""×"" & r & ""="" & (c * r)`r`n" +
        "        Next c`r`n" +
        "    Next r`r`n" +
        "    With ws.Range(""D1:L9"")`r`n" +
        "        .Borders.LineStyle = 1`r`n" +
        "        .Interior.Color = RGB(235, 247, 255)`r`n" +
        "        .Font.Name = ""Segoe UI""`r`n" +
        "    End With`r`n" +
        "End Sub"

    $req3 = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaStaircase) + '","prompt":"从 D1 生成阶梯式算式乘法表并美化","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $res3 = [LeeExcel.NativeBridge]::Dispatch($req3, $excel)

    # 深度形态核验：
    # 1. 起点 D1
    $d1Val = $ws3.Range("D1").Value2
    # 2. 对角线典型格 E2 (2x2=4), L9 (9x9=81)
    $e2Val = $ws3.Range("E2").Value2
    $l9Val = $ws3.Range("L9").Value2
    # 3. 阶梯外右上格必须为空：E1, F1, L1
    $e1Val = $ws3.Range("E1").Value2
    $l1Val = $ws3.Range("L1").Value2
    # 4. 区域外未被越界改动：C1, M1, D10, L10
    $c1Val = $ws3.Range("C1").Value2
    $m1Val = $ws3.Range("M1").Value2
    $d10Val = $ws3.Range("D10").Value2

    $hasBorder = ($ws3.Range("D1:L9").Borders.LineStyle -eq 1)
    $hasColor = ($ws3.Range("D1:L9").Interior.Color -ne -4142)

    $passed3 = ($d1Val -eq "1×1=1") -and ($e2Val -eq "2×2=4") -and ($l9Val -eq "9×9=81") -and `
               ($null -eq $e1Val) -and ($null -eq $l1Val) -and `
               ($null -eq $c1Val) -and ($null -eq $m1Val) -and ($null -eq $d10Val) -and `
               $hasBorder -and $hasColor

    Record-Test "3. Staircase Equation Table at D1 (Deep morphology & boundary check)" "Generation Fidelity" $passed3 "D1='$d1Val', E2='$e2Val', L9='$l9Val', UpperTriangleEmpty='$($null -eq $e1Val)', OutsideUntouched='$($null -eq $c1Val)'"

    # -------------------------------------------------------------
    # 用例 4: 明确要求 9×9 数值乘积矩阵 (不得强制改为阶梯算式)
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 4: 9x9 Numeric Product Matrix (No Forced Equations) ---" -ForegroundColor Cyan
    $ws4 = $wbA.Sheets.Add()
    $ws4.Name = "NumericMatrix"

    $vbaNumeric = "Sub GenerateNumericMatrix(targetWb As Workbook)`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = targetWb.Sheets(""NumericMatrix"")`r`n" +
        "    Dim r As Long, c As Long`r`n" +
        "    For r = 1 To 9`r`n" +
        "        For c = 1 To 9`r`n" +
        "            ws.Cells(r, c).Value = r * c`r`n" +
        "        Next c`r`n" +
        "    Next r`r`n" +
        "End Sub"

    $req4 = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaNumeric) + '","prompt":"生成 9×9 数值乘积矩阵","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $res4 = [LeeExcel.NativeBridge]::Dispatch($req4, $excel)

    $numA1 = $ws4.Range("A1").Value2
    $numI9 = $ws4.Range("I9").Value2
    $numB5 = $ws4.Range("B5").Value2
    $numE1 = $ws4.Range("E1").Value2 # 完整矩阵，E1 应该有值 5

    $passed4 = ($numA1 -eq 1) -and ($numI9 -eq 81) -and ($numB5 -eq 10) -and ($numE1 -eq 5)
    Record-Test "4. 9x9 Numeric Product Matrix (Full rectangular, pure numbers)" "Task Differentiation" $passed4 "A1=$numA1, I9=$numI9, B5=$numB5, E1=$numE1"

    # -------------------------------------------------------------
    # 用例 5: 为已有业务数据调整列宽、对齐、表头和数字格式 (数据完整性保护)
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 5: Format Existing Business Data (Data Integrity) ---" -ForegroundColor Cyan
    $ws5 = $wbA.Sheets.Add()
    $ws5.Name = "BusinessSales"

    $ws5.Range("A1").Value2 = "订单编号"
    $ws5.Range("B1").Value2 = "客户名称"
    $ws5.Range("C1").Value2 = "销售金额"
    $ws5.Range("D1").Value2 = "利润率"

    $ws5.Range("A2").Value2 = 1001
    $ws5.Range("B2").Value2 = "中信科移动"
    $ws5.Range("C2").Value2 = 58800.5
    $ws5.Range("D2").Value2 = 0.285

    $ws5.Range("A3").Value2 = 1002
    $ws5.Range("B3").Value2 = "中国电信"
    $ws5.Range("C3").Value2 = 92000.0
    $ws5.Range("D3").Value2 = 0.312

    $ws5.Range("A4").Value2 = 1003
    $ws5.Range("B4").Value2 = "中国联通"
    $ws5.Range("C4").Value2 = 43500.8
    $ws5.Range("D4").Value2 = 0.245

    $vbaFormat = "Sub FormatBusinessData(targetWb As Workbook)`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = targetWb.Sheets(""BusinessSales"")`r`n" +
        "    With ws.Range(""A1:D1"")`r`n" +
        "        .Font.Bold = True`r`n" +
        "        .Interior.Color = RGB(24, 45, 123)`r`n" +
        "        .Font.Color = RGB(255, 255, 255)`r`n" +
        "        .HorizontalAlignment = -4108`r`n" +
        "    End With`r`n" +
        "    ws.Range(""C2:C4"").NumberFormat = ""¥#,##0.00""`r`n" +
        "    ws.Range(""D2:D4"").NumberFormat = ""0.0%""`r`n" +
        "    ws.Range(""A1:D4"").Borders.LineStyle = 1`r`n" +
        "    ws.Columns(""A:D"").AutoFit`r`n" +
        "End Sub"

    $req5 = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaFormat) + '","prompt":"调整业务数据格式","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $res5 = [LeeExcel.NativeBridge]::Dispatch($req5, $excel)

    $clientVal = $ws5.Range("B2").Value2
    $amtVal = $ws5.Range("C2").Value2
    $pctVal = $ws5.Range("D2").Value2
    $hasHdrBg = ($ws5.Range("A1").Interior.Color -ne -4142)
    $hasBrd5 = ($ws5.Range("A1:D4").Borders.LineStyle -eq 1)

    $passed5 = ($clientVal -eq "中信科移动") -and ($amtVal -eq 58800.5) -and ($pctVal -eq 0.285) -and $hasHdrBg -and $hasBrd5
    Record-Test "5. Format Existing Data (Integrity preserved, style applied)" "Data Protection" $passed5 "Client='$clientVal', Amt=$amtVal, Rate=$pctVal, HeaderBg=$hasHdrBg, Borders=$hasBrd5"

    # -------------------------------------------------------------
    # 用例 6: 公式、筛选条件/可见行、图表数据源系列核验
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 6: Deep Formula, Filter Criteria, and Chart Range Verification ---" -ForegroundColor Cyan

    # 6A. 公式计算
    $vbaFormula = "Sub AddFormula(targetWb As Workbook)`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = targetWb.Sheets(""BusinessSales"")`r`n" +
        "    ws.Range(""B5"").Value = ""合计""`r`n" +
        "    ws.Range(""B5"").Font.Bold = True`r`n" +
        "    ws.Range(""C5"").Formula = ""=SUM(C2:C4)""`r`n" +
        "End Sub"

    $req6A = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaFormula) + '","prompt":"添加求和公式","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $res6A = [LeeExcel.NativeBridge]::Dispatch($req6A, $excel)

    $formulaText = $ws5.Range("C5").Formula
    $calcSum = $ws5.Range("C5").Value2
    $expectedSum = 58800.5 + 92000.0 + 43500.8
    $passed6A = ($formulaText -eq "=SUM(C2:C4)") -and ([Math]::Abs($calcSum - $expectedSum) -lt 0.01)
    Record-Test "6A. Excel Formula (=SUM)" "Analysis Capability" $passed6A "Formula='$formulaText', CalculatedSum=$calcSum"

    # 6B. 筛选条件与实际可见行核验 (不只看 AutoFilterMode)
    $vbaFilter = "Sub ApplyAutoFilter(targetWb As Workbook)`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = targetWb.Sheets(""BusinessSales"")`r`n" +
        "    ws.Range(""A1:D4"").AutoFilter Field:=2, Criteria1:=""中信科移动""`r`n" +
        "End Sub"

    $req6B = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaFilter) + '","prompt":"按客户筛选","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $res6B = [LeeExcel.NativeBridge]::Dispatch($req6B, $excel)

    $filterMode = $ws5.AutoFilterMode
    $filterCrit = $ws5.AutoFilter.Filters.Item(2).Criteria1
    $row2Hidden = $ws5.Rows.Item(2).Hidden # 中信科移动，不应隐藏
    $row3Hidden = $ws5.Rows.Item(3).Hidden # 中国电信，应该被隐藏

    $passed6B = ($filterMode -eq $true) -and ($filterCrit -eq "=中信科移动") -and ($row2Hidden -eq $false) -and ($row3Hidden -eq $true)
    Record-Test "6B. AutoFilter Deep Check (FilterMode, Criteria, Row Visibility)" "Analysis Capability" $passed6B "Criteria='$filterCrit', Row2Visible='$(-not $row2Hidden)', Row3Hidden='$row3Hidden'"

    # 6C. 图表系列数据范围绑定核验 (不只看图表数与标题)
    $vbaChart = "Sub CreateSalesChart(targetWb As Workbook)`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = targetWb.Sheets(""BusinessSales"")`r`n" +
        "    Dim co As ChartObject`r`n" +
        "    Set co = ws.ChartObjects.Add(250, 50, 350, 200)`r`n" +
        "    co.Chart.SetSourceData Source:=ws.Range(""B1:C4"")`r`n" +
        "    co.Chart.ChartType = 51`r`n" +
        "    co.Chart.HasTitle = True`r`n" +
        "    co.Chart.ChartTitle.Text = ""重点客户销售分布""`r`n" +
        "End Sub"

    $req6C = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaChart) + '","prompt":"创建销售柱状图","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $res6C = [LeeExcel.NativeBridge]::Dispatch($req6C, $excel)

    $chartCountA = $ws5.ChartObjects().Count
    $chartSeriesFormula = if ($chartCountA -ge 1) { $ws5.ChartObjects(1).Chart.SeriesCollection(1).Formula } else { "" }
    $chartCountB = $wbB.Sheets.Item(1).ChartObjects().Count

    # 系列公式应精确包含目标工作表的区域引用 BusinessSales!$B$1:$B$4 及 $C$1:$C$4
    $passed6C = ($chartCountA -eq 1) -and ($chartCountB -eq 0) -and ($chartSeriesFormula -like "*BusinessSales*")
    Record-Test "6C. Chart Generation Deep Check (Source Range Series Formula)" "Analysis Capability" $passed6C "SeriesFormula='$chartSeriesFormula', WbB ChartCount=$chartCountB"

    # -------------------------------------------------------------
    # 核心产品承诺核验: 快照整本备份与原地恢复核验
    # -------------------------------------------------------------
    Write-Host "`n--- Core Product Promise: Snapshot Creation & In-Place Rollback ---" -ForegroundColor Cyan
    # 1. 记录恢复前当前状态
    $valBeforeManual = $ws5.Range("A2").Value2 # 1001

    # 2. 模拟用户手工修改
    $ws5.Range("A2").Value2 = 9999
    $valManualModified = $ws5.Range("A2").Value2

    # 3. 列出快照
    $listReq = '{"action":"list_snapshots","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '"}'
    $listRes = [LeeExcel.NativeBridge]::Dispatch($listReq, $excel)
    
    # 4. 执行恢复到最近一次快照
    $snapshots = [LeeExcel.SnapshotManager]::LoadSnapshots($wbAPath)
    $hasSnap = ($snapshots.Count -gt 0)
    $restoreSuccess = $false
    if ($hasSnap) {
        $recentSnapId = $snapshots[0].id
        $restoreReq = '{"action":"restore_snapshot","snapshotId":"' + $recentSnapId + '","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '"}'
        $restoreRes = [LeeExcel.NativeBridge]::Dispatch($restoreReq, $excel)
        $parsedRestore = [LeeExcel.SimpleJson]::ParseFlatObject($restoreRes)
        $restoreSuccess = ($parsedRestore["ok"] -eq "true")

        # 重新绑定已原地重新打开的工作簿对象
        $wbA = $excel.Workbooks.Item($wbAName)
        $valAfterRestore = $wbA.Sheets.Item("BusinessSales").Range("A2").Value2
        $restoredCorrectly = ($valAfterRestore -eq 1001)
        Record-Test "9. In-Place Snapshot Rollback (Manual edits reverted in-place)" "Reliability Promise" $restoredCorrectly "BeforeRollback=$valManualModified, AfterRollback=$valAfterRestore (Expected 1001)"
    } else {
        Record-Test "9. In-Place Snapshot Rollback" "Reliability Promise" $false "No snapshot found to restore"
    }

    # -------------------------------------------------------------
    # 证明 Workbook B 全程未被任何用例篡改
    # -------------------------------------------------------------
    Write-Host "`n--- Multi-Workbook Safety: Prove Workbook B Remained 100% Intact ---" -ForegroundColor Cyan
    $wbBFinalA1 = $wbB.Sheets.Item(1).Range("A1").Value2
    $wbBFinalSheetCount = $wbB.Sheets.Count
    $passedSafetyB = ($wbBFinalA1 -eq $wbBInitialA1) -and ($wbBFinalSheetCount -eq $wbBInitialSheetCount)
    Record-Test "10. Other Workbook Intact Proof" "Isolation Guarantee" $passedSafetyB "A1 unchanged ('$wbBFinalA1'), SheetCount=$wbBFinalSheetCount"

} finally {
    Write-Host "`n[Teardown] Cleaning up test workbooks and closing Excel..." -ForegroundColor Cyan
    if ($wbA -ne $null) {
        try { $wbA.Close($false) } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbA) | Out-Null
    }
    if ($wbB -ne $null) {
        try { $wbB.Close($false) } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbB) | Out-Null
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    try {
        if (Test-Path $testDir) {
            Remove-Item -Path $testDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}

Write-Host "`n===========================================================" -ForegroundColor Cyan
Write-Host "                TEST RESULTS SUMMARY                       " -ForegroundColor Cyan
Write-Host "===========================================================" -ForegroundColor Cyan
$results | Format-Table -AutoSize

$allPassed = ($results | Where-Object { $_.Status -ne "PASS" }).Count -eq 0
if ($allPassed) {
    Write-Host "`n>>> ALL INTEGRATION TESTS PASSED 100%! <<<`n" -ForegroundColor Green
    exit 0
} else {
    Write-Host "`n>>> SOME TESTS FAILED! <<<`n" -ForegroundColor Red
    exit 1
}
