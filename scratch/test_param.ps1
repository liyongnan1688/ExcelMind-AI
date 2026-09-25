$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $vbProj = $wb.VBProject
    $mod = $vbProj.VBComponents.Add(1)
    $mod.Name = "TestParamMod"
    $vba = @"
Sub TestWithWbParam(targetWb As Workbook)
    targetWb.Sheets(1).Range("B2").Value = "PassedViaParam"
End Sub
"@
    $mod.CodeModule.AddFromString($vba)
    
    Write-Host "Calling app.Run with Workbook object..."
    $macro = "'" + $wb.Name + "'!TestWithWbParam"
    $excel.Run($macro, $wb)
    
    $res = $wb.Sheets.Item(1).Range("B2").Value2
    Write-Host "Result in cell B2:" $res
} catch {
    Write-Host "Error calling with param:" $_.Exception.Message
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
