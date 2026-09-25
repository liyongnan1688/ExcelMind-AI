# scratch/debug_compile_exact.ps1
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
    
    $vbComp = $wb.VBProject.VBComponents.Add(1)
    $vbComp.Name = "TestModule1"
    $vbComp.CodeModule.AddFromString($vbaCode)

    Write-Host "Attempting direct run of Main..."
    try {
        $excel.Run("'" + $wb.Name + "'!Main")
        Write-Host "Direct Main succeeded!"
    } catch {
        Write-Host "Direct Main error: $($_.Exception.Message)" -ForegroundColor Red
    }
} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
