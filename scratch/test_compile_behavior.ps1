$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()

$mod = $wb.VBProject.VBComponents.Add(1)
$mod.Name = "ModValid"
$mod.CodeModule.AddFromString(@"
Sub ValidSub()
    Dim x As Long
    x = 1 + 2
End Sub
"@)

$btn = $excel.VBE.CommandBars.FindControl([System.Reflection.Missing]::Value, 578)
Write-Host "Before compile: Enabled = $($btn.Enabled)"

# 执行编译
$btn.Execute()
Write-Host "After compile: Enabled = $($btn.Enabled)"

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
