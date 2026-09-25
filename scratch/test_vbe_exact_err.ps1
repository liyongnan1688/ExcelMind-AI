# scratch/test_vbe_exact_err.ps1
$ErrorActionPreference = "Stop"
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
    if ($btn) {
        Write-Host "Btn 578 found! Enabled before: $($btn.Enabled)"
        $btn.Execute()
        Write-Host "Btn 578 enabled after: $($btn.Enabled)"
    } else {
        Write-Host "Btn 578 not found"
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
