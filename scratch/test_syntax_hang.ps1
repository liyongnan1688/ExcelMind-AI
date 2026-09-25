# scratch/test_syntax_hang.ps1
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $comp = $wb.VBProject.VBComponents.Add(1)
    
    $code = @"
Sub TestBad()
    Msg = "Hi"
End CleanExit
End Sub
"@
    $comp.CodeModule.AddFromString($code)
    Write-Host "Code added."

    Write-Host "Running app.Run..."
    $excel.Run("'" + $wb.Name + "'!TestBad")
    Write-Host "Run finished!"
} catch {
    Write-Host "Caught error: $($_.Exception.Message)"
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
