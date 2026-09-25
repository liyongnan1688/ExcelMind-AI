# scratch/test_exact_task1_failure.ps1
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $testDir = Join-Path $env:TEMP ("LeeExcel_TestTask1_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
    New-Item -ItemType Directory -Path $testDir | Out-Null
    $wbPath = Join-Path $testDir "Task_1_e742ce.xlsx"

    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($wbPath)

    $content = [System.IO.File]::ReadAllText("C:\Users\35651\AppData\Local\Temp\LeeExcel_GenericE2E_0a65fc71\out_d0662052.txt", [System.Text.Encoding]::UTF8)
    
    $openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
    $start = $openMatch.Index + $openMatch.Length
    $closeIdx = $content.IndexOf('```', $start)
    $vbaCode = $content.Substring($start, $closeIdx - $start).Trim()
    
    $wrapper = "`r`nSub __LeeHostRunner_0108(targetWb As Workbook)`r`n    targetWb.Activate`r`n    Call Main`r`nEnd Sub`r`n"
    $finalCode = $vbaCode + "`r`n" + $wrapper

    $moduleName = "LeeMod_" + (Get-Date -Format "yyyyMMdd_HHmmss")
    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.Name = $moduleName
    $comp.CodeModule.AddFromString($finalCode)

    Write-Host "Injected into module: $moduleName"
    
    # Check lines in module
    $lineCount = $comp.CodeModule.CountOfLines
    Write-Host "Total lines in module: $lineCount"

    $macroAddress = "'" + $wb.Name + "'!__LeeHostRunner_0108"
    Write-Host "Calling app.Run with macroAddress: $macroAddress"
    
    $excel.Run($macroAddress, $wb)
    Write-Host "SUCCESS!" -ForegroundColor Green

} catch {
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "HRESULT: $('0x{0:X}' -f $_.Exception.HResult)" -ForegroundColor Red
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
