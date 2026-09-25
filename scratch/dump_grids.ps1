$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

Write-Host "=================================================="
Write-Host "1. FIXED VBA (Staircase Equation Table)"
Write-Host "=================================================="
$wbFixed = $excel.Workbooks.Add()
$fixedCode = @"
Sub LeeTaskEntry(targetWb As Workbook)
    Dim ws As Worksheet
    Set ws = targetWb.Worksheets.Add()
    ws.Name = "StaircaseTable"
    Dim r As Long, c As Long
    For r = 1 To 9
        For c = 1 To r
            ws.Cells(r, 3 + c).Value = c & "×" & r & "=" & (c * r)
        Next c
    Next r
End Sub
"@
$reqFixed = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($fixedCode) + '","prompt":"阶梯表","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbFixed.FullName) + '","targetWorkbookName":"' + $wbFixed.Name + '"}'
[LeeExcel.NativeBridge]::Dispatch($reqFixed, $excel) | Out-Null
$wsFixed = $wbFixed.Sheets.Item("StaircaseTable")

Write-Host "`nD1:L9 Grid for Fixed VBA (Staircase):"
for ($r = 1; $r -le 9; $r++) {
    $line = @()
    for ($c = 4; $c -le 12; $c++) {
        $val = $wsFixed.Cells.Item($r, $c).Value2
        if ($val -eq $null) {
            $line += "[  EMPTY   ]"
        } else {
            $line += ("[" + $val.PadRight(10) + "]")
        }
    }
    Write-Host ("Row $r (Col D-L): " + ($line -join " "))
}
$wbFixed.Close($false)

Write-Host "`n=================================================="
Write-Host "2. REAL LIVE LLM (Full Equation Matrix)"
Write-Host "=================================================="
$wbReal = $excel.Workbooks.Add()
$realCode = @"
Sub LeeTaskEntry(targetWb As Workbook)
    Dim ws As Worksheet
    Set ws = targetWb.Worksheets.Add()
    ws.Name = "RealLLMTable"
    Dim i As Long, j As Long
    For i = 1 To 9
        For j = 1 To 9
            ws.Cells(i, 3 + j).Value = i & "×" & j & "=" & (i * j)
        Next j
    Next i
End Sub
"@
$reqReal = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($realCode) + '","prompt":"真实模型矩阵","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbReal.FullName) + '","targetWorkbookName":"' + $wbReal.Name + '"}'
[LeeExcel.NativeBridge]::Dispatch($reqReal, $excel) | Out-Null
$wsReal = $wbReal.Sheets.Item("RealLLMTable")

Write-Host "`nD1:L9 Grid for Real LLM (Full Matrix):"
for ($r = 1; $r -le 9; $r++) {
    $line = @()
    for ($c = 4; $c -le 12; $c++) {
        $val = $wsReal.Cells.Item($r, $c).Value2
        if ($val -eq $null) {
            $line += "[  EMPTY   ]"
        } else {
            $line += ("[" + $val.PadRight(10) + "]")
        }
    }
    Write-Host ("Row $r (Col D-L): " + ($line -join " "))
}
$wbReal.Close($false)

$excel.Quit()
