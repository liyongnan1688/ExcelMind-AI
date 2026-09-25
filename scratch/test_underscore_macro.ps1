# scratch/test_underscore_macro.ps1
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $comp = $wb.VBProject.VBComponents.Add(1)
    
    $code = @"
Sub NormalName(targetWb As Workbook)
    Range("A1").Value = "NormalNameOK"
End Sub

Sub _SingleUnderscore(targetWb As Workbook)
    Range("A2").Value = "SingleUnderscoreOK"
End Sub

Sub __DoubleUnderscore(targetWb As Workbook)
    Range("A3").Value = "DoubleUnderscoreOK"
End Sub
"@
    $comp.CodeModule.AddFromString($code)

    try {
        $excel.Run("'" + $wb.Name + "'!NormalName", $wb)
        Write-Host "NormalName: SUCCESS!" -ForegroundColor Green
    } catch {
        Write-Host "NormalName: FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }

    try {
        $excel.Run("'" + $wb.Name + "'!_SingleUnderscore", $wb)
        Write-Host "_SingleUnderscore: SUCCESS!" -ForegroundColor Green
    } catch {
        Write-Host "_SingleUnderscore: FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }

    try {
        $excel.Run("'" + $wb.Name + "'!__DoubleUnderscore", $wb)
        Write-Host "__DoubleUnderscore: SUCCESS!" -ForegroundColor Green
    } catch {
        Write-Host "__DoubleUnderscore: FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }

} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
