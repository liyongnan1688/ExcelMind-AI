# scratch/find_compile_error_line.ps1
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $vbaPath = "C:\Users\35651\AppData\Local\Temp\LeeExcel_WarehouseCycle_41cb15fe\Version1_Original_RiskFound.vba"
    $lines = [System.IO.File]::ReadAllLines($vbaPath, [System.Text.Encoding]::UTF8)

    # Test lines incrementally
    for ($count = 10; $count -le $lines.Count; $count += 10) {
        $subLines = $lines[0..($count-1)] + "End Sub"
        $code = $subLines -join "`r`n"
        
        $vbComp = $wb.VBProject.VBComponents.Add(1)
        $vbComp.CodeModule.AddFromString($code)
        try { $vbComp.Activate() } catch {}
        
        $btn = $excel.VBE.CommandBars.FindControl([Type]::Missing, 578)
        if ($btn -and $btn.Enabled) {
            $btn.Execute()
            if ($btn.Enabled) {
                Write-Host "Compile error introduced between line $($count-10) and $count!"
                $wb.VBProject.VBComponents.Remove($vbComp)
                break
            }
        }
        $wb.VBProject.VBComponents.Remove($vbComp)
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
