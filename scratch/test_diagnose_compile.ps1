# scratch/test_diagnose_compile.ps1
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

    Write-Host "Checking compile button..."
    $btn = $excel.VBE.CommandBars.FindControl($null, 578)
    if ($btn) {
        Write-Host "Btn enabled before: $($btn.Enabled)"
        $btn.Execute()
        Write-Host "Btn enabled after: $($btn.Enabled)"
    } else {
        Write-Host "Btn 578 not found"
    }

    # Now let's try running it directly to see if it runs or what runtime error it throws
    Write-Host "Trying to run BuildMultiWarehouseDashboard..."
    try {
        $excel.Run("'" + $wb.Name + "'!BuildMultiWarehouseDashboard")
        Write-Host "Run succeeded!" -ForegroundColor Green
        Write-Host "ActiveSheet: $($wb.ActiveSheet.Name), UsedRange: $($wb.ActiveSheet.UsedRange.Address($false, $false))"
    } catch {
        Write-Host "Run error: $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception.InnerException) {
            Write-Host "Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        }
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
