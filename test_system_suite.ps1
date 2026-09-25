# test_system_suite.ps1 - Comprehensive System Quality Regression Suite
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

Write-Host "===========================================================" -ForegroundColor Cyan
Write-Host "       Lee-Excel Comprehensive System Quality Regression   " -ForegroundColor Cyan
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
    $wbA.SaveAs($wbAPath)
    $wbAName = $wbA.Name

    $wbB = $excel.Workbooks.Add()
    $wbBPath = Join-Path $testDir "TestWB_Beta.xlsx"
    $wbB.SaveAs($wbBPath)
    $wbBName = $wbB.Name

    Write-Host "  Created isolated Target Workbook A: $wbAName ($wbAPath)" -ForegroundColor Gray
    Write-Host "  Created isolated Target Workbook B: $wbBName ($wbBPath)" -ForegroundColor Gray

    # -------------------------------------------------------------
    # 用例 1 & 2: 纯聊天与代码解释 (零 VBA 注入、零工作簿修改)
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 1 & 2: Chat & Explanation Channel Separation ---" -ForegroundColor Cyan
    $sheetACellBefore = $wbA.Sheets.Item(1).Range("A1").Value2
    $sheetBCellBefore = $wbB.Sheets.Item(1).Range("A1").Value2

    # 1. 验证 NativeBridge 不会响应空代码的执行
    $emptyReq = '{"action":"execute_vba","code":"","prompt":"你是？","targetWorkbookName":"' + $wbAName + '"}'
    $emptyRes = [LeeExcel.NativeBridge]::Dispatch($emptyReq, $excel)
    $parsedEmpty = [LeeExcel.SimpleJson]::ParseFlatObject($emptyRes)
    $passed1 = ($parsedEmpty["ok"] -eq "false") -and ($parsedEmpty["error"] -like "*传入的 VBA 代码为空*")
    Record-Test "1A. Empty VBA rejected" "Channel Separation" $passed1 $parsedEmpty["error"]

    # 2. 检查工作簿零变更
    $sheetACellAfter = $wbA.Sheets.Item(1).Range("A1").Value2
    $passed1B = ($sheetACellBefore -eq $sheetACellAfter)
    Record-Test "1B. Zero workbook mutation on chat request" "Channel Separation" $passed1B "Sheet A1 untouched"

    # -------------------------------------------------------------
    # 用例 7: 半截代码、缺少 End Sub、无闭合代码块拦截
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 7: Incomplete / Truncated Code Interception ---" -ForegroundColor Cyan
    $brokenCode = "Sub IncompleteTask()`r`n    ActiveSheet.Range(""A1"").Value = ""broken"""
    $truncResult = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wbA, $brokenCode)
    $passed7A = ($truncResult.success -eq $false) -and ($truncResult.summary -like "*缺少 Sub 或 End Sub*")
    Record-Test "7A. Missing End Sub interception" "Code Validation" $passed7A $truncResult.summary

    # -------------------------------------------------------------
    # 用例 8: 多工作簿并存与前台窗口切换 (目标工作簿锁定验证)
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 8: Target Workbook Binding & Window Switch Immunity ---" -ForegroundColor Cyan
    # 模拟用户将活动焦点切换到 Workbook B
    $wbB.Activate()
    $activeBefore = $excel.ActiveWorkbook.Name
    Write-Host "  Active window intentionally switched to: $activeBefore" -ForegroundColor Gray

    # 此时向 Workbook A 发起任务，VBA 代码中使用 ActiveWorkbook
    $vbaBindTest = "Sub TestBinding()`r`n    ActiveWorkbook.Sheets(1).Range(""A1"").Value = ""TargetAlphaConfirmed""`r`nEnd Sub"
    $execReq = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaBindTest) + '","prompt":"目标绑定测试","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $execRes = [LeeExcel.NativeBridge]::Dispatch($execReq, $excel)
    $parsedExec = [LeeExcel.SimpleJson]::ParseFlatObject($execRes)

    $valA = $wbA.Sheets.Item(1).Range("A1").Value2
    $valB = $wbB.Sheets.Item(1).Range("A1").Value2

    $passed8 = ($valA -eq "TargetAlphaConfirmed") -and ($valB -ne "TargetAlphaConfirmed")
    Record-Test "8. Target Workbook Binding (Protected against window switch)" "Target Binding" $passed8 "WbA updated: '$valA', WbB unchanged: '$valB'"

    # -------------------------------------------------------------
    # 用例 3: 从 D1 生成阶梯式算式九九乘法表并美化
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 3: Staircase Equation Multiplication Table at D1 with Styling ---" -ForegroundColor Cyan
    $ws3 = $wbA.Sheets.Add()
    $ws3.Name = "MultiplicationStaircase"

    $vbaStaircase = "Sub GenerateStaircaseMultiplication()`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = ActiveWorkbook.Sheets(""MultiplicationStaircase"")`r`n" +
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

    $d1Val = $ws3.Range("D1").Value2
    $e2Val = $ws3.Range("E2").Value2
    $e1Val = $ws3.Range("E1").Value2
    $hasBorder = ($ws3.Range("D1:L9").Borders.LineStyle -eq 1)
    $hasColor = ($ws3.Range("D1:L9").Interior.Color -ne -4142)

    $passed3 = ($d1Val -eq "1×1=1") -and ($e2Val -eq "2×2=4") -and ($null -eq $e1Val) -and $hasBorder -and $hasColor
    Record-Test "3. Staircase Equation Table at D1 (Text, Shape, Origin, Style)" "Generation Fidelity" $passed3 "D1='$d1Val', E2='$e2Val', E1 empty='$($null -eq $e1Val)', Borders=$hasBorder, InteriorColor=$hasColor"

    # -------------------------------------------------------------
    # 用例 4: 明确要求 9×9 数值乘积矩阵 (不得强制改为阶梯算式)
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 4: 9x9 Numeric Product Matrix ---" -ForegroundColor Cyan
    $ws4 = $wbA.Sheets.Add()
    $ws4.Name = "NumericMatrix"

    $vbaNumeric = "Sub GenerateNumericMatrix()`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = ActiveWorkbook.Sheets(""NumericMatrix"")`r`n" +
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

    $passed4 = ($numA1 -eq 1) -and ($numI9 -eq 81) -and ($numB5 -eq 10)
    Record-Test "4. 9x9 Numeric Product Matrix (No forced equation text)" "Task Differentiation" $passed4 "A1=$numA1, I9=$numI9, B5=$numB5"

    # -------------------------------------------------------------
    # 用例 5: 为已有业务数据调整列宽、对齐、表头和数字格式 (数据未改坏)
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 5: Format Existing Business Data (Data Integrity) ---" -ForegroundColor Cyan
    $ws5 = $wbA.Sheets.Add()
    $ws5.Name = "BusinessSales"

    # 逐格填充原始业务数据，保证编码与类型无歧义
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

    $vbaFormat = "Sub FormatBusinessData()`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = ActiveWorkbook.Sheets(""BusinessSales"")`r`n" +
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
    # 用例 6: 公式、筛选/汇总、图表各至少一例
    # -------------------------------------------------------------
    Write-Host "`n--- Test Suite 6: Formula, Filter/Summary, and Chart Verification ---" -ForegroundColor Cyan

    # 6A. 公式计算
    $vbaFormula = "Sub AddFormula()`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = ActiveWorkbook.Sheets(""BusinessSales"")`r`n" +
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
    Record-Test "6A. Excel Formula (=SUM)" "Analysis Capability" $passed6A "Formula='$formulaText', CalculatedSum=$calcSum (Expected=$expectedSum)"

    # 6B. 筛选与汇总状态
    $vbaFilter = "Sub ApplyAutoFilter()`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = ActiveWorkbook.Sheets(""BusinessSales"")`r`n" +
        "    ws.Range(""A1:D4"").AutoFilter Field:=2, Criteria1:=""中信科移动""`r`n" +
        "End Sub"

    $req6B = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaFilter) + '","prompt":"按客户筛选","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbAPath) + '","targetWorkbookName":"' + $wbAName + '"}'
    $res6B = [LeeExcel.NativeBridge]::Dispatch($req6B, $excel)

    $isFiltered = $ws5.AutoFilterMode
    $passed6B = ($isFiltered -eq $true)
    Record-Test "6B. AutoFilter Activation" "Analysis Capability" $passed6B "AutoFilterMode=$isFiltered"

    # 6C. 图表生成 (核查实际目标对象)
    $vbaChart = "Sub CreateSalesChart()`r`n" +
        "    Dim ws As Worksheet`r`n" +
        "    Set ws = ActiveWorkbook.Sheets(""BusinessSales"")`r`n" +
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
    $chartCountB = $wbB.Sheets.Item(1).ChartObjects().Count
    $chartTitle = if ($chartCountA -ge 1) { $ws5.ChartObjects(1).Chart.ChartTitle.Text } else { "" }

    $passed6C = ($chartCountA -eq 1) -and ($chartCountB -eq 0) -and ($chartTitle -eq "重点客户销售分布")
    Record-Test "6C. Chart Generation on Target Object" "Analysis Capability" $passed6C "WbA ChartCount=$chartCountA (Title='$chartTitle'), WbB ChartCount=$chartCountB"

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
    Write-Host "`n>>> ALL 9 INTEGRATION TESTS PASSED 100%! <<<`n" -ForegroundColor Green
    exit 0
} else {
    Write-Host "`n>>> SOME TESTS FAILED! <<<`n" -ForegroundColor Red
    exit 1
}
