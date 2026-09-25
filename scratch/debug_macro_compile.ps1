# scratch/debug_macro_compile.ps1
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $content = [System.IO.File]::ReadAllText("C:\Users\35651\AppData\Local\Temp\LeeExcel_GenericE2E_0a65fc71\out_d0662052.txt", [System.Text.Encoding]::UTF8)
    
    # Extract VBA
    $openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
    $start = $openMatch.Index + $openMatch.Length
    $closeIdx = $content.IndexOf('```', $start)
    $vbaCode = $content.Substring($start, $closeIdx - $start).Trim()
    
    $wrapper = @"
' ===== [LeeExcel 自动生成的受控调用入口包装器 - 保持模型正文源码零篡改] =====
Sub __LeeHostRunner_0108(targetWb As Workbook)
    targetWb.Activate
    Call Main
End Sub
"@
    $finalCode = $vbaCode + "`r`n" + $wrapper

    $vbComp = $wb.VBProject.VBComponents.Add(1)
    $vbComp.Name = "TestModule1"
    $vbComp.CodeModule.AddFromString($finalCode)

    Write-Host "Code added to module successfully."
    
    # Try running the macro
    Write-Host "Attempting app.Run..."
    $macroAddress = "'" + $wb.Name + "'!__LeeHostRunner_0108"
    $excel.Run($macroAddress, $wb)
    Write-Host "app.Run succeeded!"
} catch {
    Write-Host "Caught Exception: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "HRESULT: $('0x{0:X}' -f $_.Exception.HResult)" -ForegroundColor Red
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
