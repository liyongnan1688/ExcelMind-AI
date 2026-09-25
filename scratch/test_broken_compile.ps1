$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()

$mod = $wb.VBProject.VBComponents.Add(1)
$mod.Name = "ModBroken"
$mod.CodeModule.AddFromString(@"
Sub BrokenSub()
    ws.Columns(1.ColumnWidth) = 10
End Sub
"@)

$btn = $excel.VBE.CommandBars.FindControl([System.Reflection.Missing]::Value, 578)
Write-Host "Broken before compile: Enabled = $($btn.Enabled)"

try {
    Write-Host "Executing compile on broken code..."
    $btn.Execute()
    Write-Host "After broken compile: Enabled = $($btn.Enabled)"
} catch {
    Write-Host "Broken compile threw: $($_.Exception.Message)"
} finally {
    Write-Host "Done test."
}
