$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$wb = $excel.Workbooks.Add()
$vbProj = $wb.VBProject
$mod = $vbProj.VBComponents.Add(1)
$mod.Name = "TestBrokenMod"
$mod.CodeModule.AddFromString(@"
Sub BrokenMacro()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Columns(1.ColumnWidth) = 10
End Sub
"@)

# 让 VBE 打开并查看代码
try {
    $excel.VBE.MainWindow.Visible = $true
} catch {}

Write-Host "Excel and VBE are visible. Calling Run in 2 seconds..."
Start-Sleep -Seconds 2

try {
    $excel.Run("TestBrokenMod.BrokenMacro")
} catch {
    Write-Host "Caught: $($_.Exception.Message)"
}
