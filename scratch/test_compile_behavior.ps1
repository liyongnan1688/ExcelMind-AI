$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Add()
    $btn = $excel.VBE.CommandBars.FindControl([Type]::Missing, 578)
    
    # 1. Test code with a variable not defined (Option Explicit)
    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.CodeModule.AddFromString("Option Explicit`r`nSub TestErr()`r`n undeclaredVar = 123`r`nEnd Sub")
    
    Write-Host "Btn.Enabled before: $($btn.Enabled)"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $btn.Execute()
    $sw.Stop()
    Write-Host "Btn.Enabled after: $($btn.Enabled) | Time: $($sw.ElapsedMilliseconds) ms"
    $wb.VBProject.VBComponents.Remove($comp)

    # 2. Test code with a pure syntax error (missing quote)
    $comp2 = $wb.VBProject.VBComponents.Add(1)
    $comp2.CodeModule.AddFromString("Sub TestSyntax()`r`n x = `"abc`r`nEnd Sub")
    Write-Host "`nBtn2.Enabled before: $($btn.Enabled)"
    $sw2 = [System.Diagnostics.Stopwatch]::StartNew()
    $btn.Execute()
    $sw2.Stop()
    Write-Host "Btn2.Enabled after: $($btn.Enabled) | Time: $($sw2.ElapsedMilliseconds) ms"
    $wb.VBProject.VBComponents.Remove($comp2)

    # 3. Test code with VALID code
    $comp3 = $wb.VBProject.VBComponents.Add(1)
    $comp3.CodeModule.AddFromString("Sub TestValid()`r`n Dim x As Long: x = 123`r`nEnd Sub")
    Write-Host "`nBtn3.Enabled before: $($btn.Enabled)"
    $sw3 = [System.Diagnostics.Stopwatch]::StartNew()
    $btn.Execute()
    $sw3.Stop()
    Write-Host "Btn3.Enabled after: $($btn.Enabled) | Time: $($sw3.ElapsedMilliseconds) ms"
    $wb.VBProject.VBComponents.Remove($comp3)

} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
