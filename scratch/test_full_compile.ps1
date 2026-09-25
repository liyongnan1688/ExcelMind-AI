# scratch/test_full_compile.ps1
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $vbaPath = "C:\Users\35651\AppData\Local\Temp\LeeExcel_WarehouseCycle_41cb15fe\Version1_Original_RiskFound.vba"
    $code = [System.IO.File]::ReadAllText($vbaPath, [System.Text.Encoding]::UTF8)

    $vbComp = $wb.VBProject.VBComponents.Add(1)
    $vbComp.Name = "TestDiagMod"
    $vbComp.CodeModule.AddFromString($code)
    try { $vbComp.Activate() } catch {}

    $btn = $excel.VBE.CommandBars.FindControl([Type]::Missing, 578)
    Write-Host "Btn 578 enabled before: $($btn.Enabled)"
    $btn.Execute()
    Write-Host "Btn 578 enabled after: $($btn.Enabled)"

    if ($btn.Enabled) {
        Write-Host "Compile still failed on full file!"
    } else {
        Write-Host "Compile PASSED on full file!" -ForegroundColor Green
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
