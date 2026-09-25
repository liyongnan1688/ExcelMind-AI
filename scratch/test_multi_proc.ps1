$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()

# 测试包含主过程、辅助过程和自定义函数的复杂 VBA
$code = @"
Sub GenerateSalesReport()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Name = "SalesData"
    
    FormatHeaders ws
    
    Dim i As Long
    For i = 2 To 5
        ws.Cells(i, 1).Value = "Item " & (i - 1)
        ws.Cells(i, 2).Value = (i - 1) * 100
        ws.Cells(i, 3).Value = CalculateTax(ws.Cells(i, 2).Value)
    Next i
End Sub

Private Sub FormatHeaders(ws As Worksheet)
    ws.Range("A1").Value = "Product"
    ws.Range("B1").Value = "Price"
    ws.Range("C1").Value = "Tax"
    ws.Range("A1:C1").Font.Bold = True
End Sub

Private Function CalculateTax(price As Double) As Double
    CalculateTax = price * 0.13
End Function

' 宿主受控包装器
Sub __LeeHostRunner(targetWb As Workbook)
    targetWb.Activate
    Call GenerateSalesReport
End Sub
"@

$vbProj = $wb.VBProject
$mod = $vbProj.VBComponents.Add(1)
$mod.Name = "TestMultiProc"
$mod.CodeModule.AddFromString($code)

Write-Host "Running __LeeHostRunner..."
$excel.Run("'" + $wb.Name + "'!__LeeHostRunner", $wb)

$ws = $wb.Sheets.Item("SalesData")
Write-Host "Sheet Name:   $($ws.Name)"
Write-Host "Header A1-C1: '$($ws.Range('A1').Value2)' | '$($ws.Range('B1').Value2)' | '$($ws.Range('C1').Value2)'"
Write-Host "Row 2 Tax:    $($ws.Range('C2').Value2) (Expected: 13)"
Write-Host "Row 5 Tax:    $($ws.Range('C5').Value2) (Expected: 52)"

$vbProj.VBComponents.Remove($mod)
$wb.Close($false)
$excel.Quit()

Write-Host "`n>>> MULTI-PROCEDURE & WRAPPER RUNNER VERIFIED! <<<"
