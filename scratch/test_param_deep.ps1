$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$testPath = Join-Path $env:TEMP ("ParamDeepTest_" + [Guid]::NewGuid().ToString("N").Substring(0, 8) + ".xlsx")
$wb = $excel.Workbooks.Add()
$wb.SaveAs($testPath, 51) # .xlsx

Write-Host "Target Workbook FullName: $($wb.FullName)"
Write-Host "Target Workbook Name:     $($wb.Name)"

$vbaCode = @"
Sub LeeTaskEntry(targetWb As Workbook)
    Dim wsLog As Worksheet
    Set wsLog = targetWb.Sheets(1)
    
    ' 记录 VBA 内部实际接收到的真实变量信息
    wsLog.Range("A1").Value = "ParamVerified"
    wsLog.Range("A2").Value = TypeName(targetWb)
    wsLog.Range("A3").Value = targetWb.Name
    wsLog.Range("A4").Value = targetWb.FullName
    wsLog.Range("A5").Value = (targetWb Is ThisWorkbook)
    wsLog.Range("A6").Value = (targetWb.Sheets.Count)
End Sub
"@

$res = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb, $vbaCode)

Write-Host "`nRunVbaCode Result:"
Write-Host "  Success : $($res.success)"
Write-Host "  Summary : $($res.summary)"
Write-Host "  Error   : $($res.error)"

# 从 Excel 单元格中读出 VBA 运行时的真实状态
$ws = $wb.Sheets.Item(1)
$vTag = $ws.Range("A1").Value2
$vType = $ws.Range("A2").Value2
$vName = $ws.Range("A3").Value2
$vFull = $ws.Range("A4").Value2
$vIsThis = $ws.Range("A5").Value2
$vSheets = $ws.Range("A6").Value2

Write-Host "`nVBA Internal Reflection Values written by LeeTaskEntry:"
Write-Host "  Cell A1 (Tag)             : '$vTag'"
Write-Host "  Cell A2 (TypeName)        : '$vType' (Expected: 'Workbook')"
Write-Host "  Cell A3 (targetWb.Name)    : '$vName'"
Write-Host "  Cell A4 (targetWb.FullName): '$vFull'"
Write-Host "  Cell A5 (Is ThisWorkbook) : '$vIsThis'"
Write-Host "  Cell A6 (Sheets.Count)    : '$vSheets'"

$typeMatches = ($vType -eq "Workbook")
$nameMatches = ($vName -eq $wb.Name)
$fullMatches = ($vFull -eq $wb.FullName)

Write-Host "`nParameter Passing Assertions:"
Write-Host "  TypeName is Workbook      : $typeMatches"
Write-Host "  targetWb.Name matches     : $nameMatches"
Write-Host "  targetWb.FullName matches : $fullMatches"

$wb.Close($false)
$excel.Quit()

Remove-Item $testPath -Force -ErrorAction SilentlyContinue

if ($typeMatches -and $nameMatches -and $fullMatches) {
    Write-Host "`n>>> COM OBJECT PARAMETER PASSING VERIFIED 100%! <<<`n" -ForegroundColor Green
    exit 0
} else {
    Write-Host "`n>>> PARAMETER PASSING FAILED! <<<`n" -ForegroundColor Red
    exit 1
}
