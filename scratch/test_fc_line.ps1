# scratch/test_fc_line.ps1
$vbaPath = "C:\Users\35651\AppData\Local\Temp\LeeExcel_WarehouseCycle_41cb15fe\Version1_Original_RiskFound.vba"
$code = [System.IO.File]::ReadAllText($vbaPath, [System.Text.Encoding]::UTF8)

# Replace Formula:= with Formula1:=
$fixedCode = $code.Replace("Formula:=", "Formula1:=")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $vbComp = $wb.VBProject.VBComponents.Add(1)
    $vbComp.Name = "TestDiagMod"
    $vbComp.CodeModule.AddFromString($fixedCode)
    try { $vbComp.Activate() } catch {}

    $btn = $excel.VBE.CommandBars.FindControl([Type]::Missing, 578)
    Write-Host "Btn 578 enabled before: $($btn.Enabled)"
    $btn.Execute()
    Write-Host "Btn 578 enabled after: $($btn.Enabled)"
    if (!$btn.Enabled) {
        Write-Host "FIXED! Formula1:= was the exact compile error!" -ForegroundColor Green
    } else {
        Write-Host "Still not fixed."
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
