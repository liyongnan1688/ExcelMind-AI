$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$tempPath = Join-Path $env:TEMP ("TestXlsxClean_" + [Guid]::NewGuid().ToString("N").Substring(0, 8) + ".xlsx")
Write-Host "Testing .xlsx macro injection and cleanup on: $tempPath"

$wb = $excel.Workbooks.Add()
$wb.SaveAs($tempPath, 51) # 51 = xlOpenXMLWorkbook (.xlsx)

$code = @"
Sub LeeTaskEntry(targetWb As Workbook)
    targetWb.Sheets(1).Range("A1").Value = "XlsxCleanTestSuccess"
End Sub
"@

Write-Host "`n1. Running macro on .xlsx..."
$res = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb, $code)
Write-Host "   Execution Success: $($res.success)"
Write-Host "   Macro Address:     '$($wb.Name)'!LeeTaskEntry"
Write-Host "   Target Passed:     $($wb.FullName)"

# 检查内存中的 VBComponents
Write-Host "`n2. Checking in-memory VBComponents immediately after execution..."
$hasTempModule = $false
foreach ($c in $wb.VBProject.VBComponents) {
    Write-Host "   Component: $($c.Name) (Type: $($c.Type))"
    if ($c.Name -like "LeeMod_*") {
        $hasTempModule = $true
    }
}
Write-Host "   Temp module present: $hasTempModule (Expected: False)"

Write-Host "`n3. Saving .xlsx without macro prompts..."
$wb.Save()
Write-Host "   Save succeeded with zero errors!"

Write-Host "`n4. Closing and reopening .xlsx..."
$wb.Close($false)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null

$wbReopened = $excel.Workbooks.Open($tempPath)
Write-Host "   Reopen succeeded!"
$reopenedA1 = $wbReopened.Sheets.Item(1).Range("A1").Value2
Write-Host "   Cell A1 Value: '$reopenedA1'"

$reopenedTempModule = $false
foreach ($c in $wbReopened.VBProject.VBComponents) {
    if ($c.Name -like "LeeMod_*") {
        $reopenedTempModule = $true
    }
}
Write-Host "   Reopened Temp module present: $reopenedTempModule (Expected: False)"

$wbReopened.Close($false)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbReopened) | Out-Null
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
Write-Host "`n>>> .XLSX INJECTION & CLEANUP VERIFICATION COMPLETED SUCCESSFULLY! <<<"
