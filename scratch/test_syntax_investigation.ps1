Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
$wbPath = Join-Path $env:TEMP "TestTask4_Debug.xlsx"
$wb.SaveAs($wbPath)

# 读入 Task 4 大模型实际生成的 VBA 代码
$task4Out = "C:\Users\35651\AppData\Local\Temp\LeeExcel_GenericE2E_615406b5\out_b3ad1627.txt"
$content = [System.IO.File]::ReadAllText($task4Out, [System.Text.Encoding]::UTF8)

# 提取代码
$openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
$start = $openMatch.Index + $openMatch.Length
$closeIdx = $content.IndexOf('```', $start)
$vbaCode = $content.Substring($start, $closeIdx - $start).Trim()

Write-Host "VBA Code Length: $($vbaCode.Length)"
Write-Host "Dispatching..."

$req = @{
    action = "execute_vba"
    code = $vbaCode
    prompt = "Task 4 Debug"
    targetWorkbookName = $wb.Name
    targetWorkbookFullName = $wbPath
    rawModelResponse = $content
}
$reqJson = ($req | ConvertTo-Json -Compress)

try {
    $resJson = [LeeExcel.NativeBridge]::Dispatch($reqJson, $excel)
    Write-Host "Dispatch returned successfully!"
    Write-Host $resJson
} catch {
    Write-Host "Dispatch threw: $($_.Exception.ToString())"
} finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
