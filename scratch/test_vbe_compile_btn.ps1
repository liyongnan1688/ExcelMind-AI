$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
$vbProj = $wb.VBProject

Write-Host "Checking VBE CommandBars..."
$compileBtn = $excel.VBE.CommandBars.FindControl([System.Reflection.Missing]::Value, 578)
if ($compileBtn) {
    Write-Host "Found Compile Button! Caption = '$($compileBtn.Caption)', Enabled = $($compileBtn.Enabled)"
} else {
    Write-Host "Compile button not found by ID 578"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
