$latest = (Get-ChildItem $env:TEMP -Filter "LeeExcel_Task2_NoThink_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
$code = [System.IO.File]::ReadAllText((Join-Path $latest "out.txt"))
$openMatch = [regex]::Match($code, '```(?:vba|vb)?\s*')
$start = $openMatch.Index + $openMatch.Length
$closeIdx = $code.IndexOf('```', $start)
$vba = $code.Substring($start, $closeIdx - $start).Trim()

$vbaFile = Join-Path $latest "pure_vba.bas"
[System.IO.File]::WriteAllText($vbaFile, $vba, [System.Text.Encoding]::UTF8)

# 打印代码前 50 行和后 50 行
Write-Host "VBA Length: $($vba.Length)"
$lines = $vba -split "`r?`n"
Write-Host "Total lines: $($lines.Count)"
Write-Host "First 20 lines:"
$lines[0..20] | ForEach-Object { Write-Host $_ }
Write-Host "... Last 30 lines:"
$lines[($lines.Count-30)..($lines.Count-1)] | ForEach-Object { Write-Host $_ }
