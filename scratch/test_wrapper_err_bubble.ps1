$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$mod = $wb.VBProject.VBComponents.Add(1)
$mod.Name = "TestMod"

$code = @"
Sub BrokenLogic()
    Dim a As Long, b As Long
    a = 10
    b = 0
    ActiveSheet.Range("A1").Value = a / b
End Sub

Sub LeeHostRunner_Test(targetWb As Workbook)
    On Error GoTo ErrHandler
    targetWb.Activate
    Call BrokenLogic
    Exit Sub
ErrHandler:
    Err.Raise Err.Number, "LeeExcelHost", Err.Description
End Sub
"@

$mod.CodeModule.AddFromString($code)

Write-Host "Calling LeeHostRunner_Test on broken logic..."
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $excel.Run("TestMod.LeeHostRunner_Test", $wb)
    Write-Host "Returned ok?"
} catch {
    $sw.Stop()
    Write-Host "COM Exception caught in $($sw.ElapsedMilliseconds) ms!"
    Write-Host "Exception message: $($_.Exception.Message)"
} finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
