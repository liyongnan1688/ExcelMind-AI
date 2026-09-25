# scratch/test_name_fix.ps1
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $content = [System.IO.File]::ReadAllText("C:\Users\35651\AppData\Local\Temp\LeeExcel_GenericE2E_0a65fc71\out_d0662052.txt", [System.Text.Encoding]::UTF8)
    
    $openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
    $start = $openMatch.Index + $openMatch.Length
    $closeIdx = $content.IndexOf('```', $start)
    $vbaCode = $content.Substring($start, $closeIdx - $start).Trim()
    
    # Try with NO leading underscore: LeeHostRunner
    $wrapper = "`r`nSub LeeHostRunner(targetWb As Workbook)`r`n    targetWb.Activate`r`n    Call Main`r`nEnd Sub`r`n"
    $finalCode = $vbaCode + "`r`n" + $wrapper

    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.Name = "LeeMod_Test"
    $comp.CodeModule.AddFromString($finalCode)

    $macroAddress = "'" + $wb.Name + "'!LeeHostRunner"
    Write-Host "Calling app.Run: $macroAddress"
    $excel.Run($macroAddress, $wb)
    Write-Host "LeeHostRunner SUCCESS!" -ForegroundColor Green

    $valA1 = $wb.Sheets.Item(1).Range("A1").Value2
    Write-Host "Cell A1 = $valA1"

} catch {
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
