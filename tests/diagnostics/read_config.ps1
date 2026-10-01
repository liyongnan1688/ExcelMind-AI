$logPath = Join-Path $env:LOCALAPPDATA "LeeExcel\WebView2Profile\EBWebView\Default\Local Storage\leveldb\000003.log"
if (!(Test-Path $logPath)) {
    Write-Host "Log file not found."
    exit 0
}

$fs = [System.IO.File]::Open($logPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
$content = $sr.ReadToEnd()
$sr.Close()
$fs.Close()

$match = [regex]::Match($content, '\{[^{}]*provider[^{}]*\}')
if ($match.Success) {
    $raw = $match.Value
    $desensitized = [regex]::Replace($raw, '("apiKey":\s*")[^"]+(")', '$1***MASKED***$2')
    Write-Host "Config found:"
    Write-Host $desensitized
} else {
    Write-Host "No JSON config matching provider pattern."
}
