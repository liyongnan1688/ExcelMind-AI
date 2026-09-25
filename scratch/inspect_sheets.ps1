# scratch/inspect_sheets.ps1
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Open("C:\Users\35651\AppData\Local\Temp\LeeExcel_WarehouseClean_c5706038\WarehouseInventory_Success.xlsx")
    foreach ($ws in $wb.Worksheets) {
        Write-Host "Worksheet: '$($ws.Name)', UsedRange: '$($ws.UsedRange.Address($false, $false))'"
        for ($r = 1; $r -le [Math]::Min($ws.UsedRange.Rows.Count, 15); $r++) {
            $vals = @()
            for ($c = 1; $c -le [Math]::Min($ws.UsedRange.Columns.Count, 8); $c++) {
                $vals += $ws.Cells.Item($r, $c).Text
            }
            Write-Host "  R$r : $($vals -join ' | ')"
        }
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
