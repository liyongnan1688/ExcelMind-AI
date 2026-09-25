# scratch/test_task2_thinking_disabled.ps1
# 验证 Task 2 在 thinking: { type: "disabled" } 时的真实模型端到端表现
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host " Task 2 (thinking: disabled) 真实模型全额代码预算对照实测 " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

$testRunDir = Join-Path $env:TEMP ("LeeExcel_Task2_NoThink_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $testRunDir | Out-Null

$prompt2 = "制作一个多仓库库存调拨与警戒看板：新建一张表，录入4个仓库（北京仓、上海仓、广州仓、成都仓）中5类电子元器件的当前库存与安全库存警戒线；自动计算各仓库总库存量及各品类的全网缺货差额；用黄色高亮低于安全库存的缺货单元格；在顶部生成3个大号指标卡片（总库存件数、总缺货品次、最低警戒仓库）"
$wbPath = Join-Path $testRunDir "WarehouseInventory_NoThink.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$wb.SaveAs($wbPath)

$sysPrompt = "你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。请直接输出完整可执行的标准 VBA 代码，包裹在单个 ```vba ... ``` 代码块中，以 End Sub 正常闭合。"
$tempSys = Join-Path $testRunDir "sys.txt"
$tempOut = Join-Path $testRunDir "out.txt"
$tempMeta = Join-Path $testRunDir "meta.json"
[System.IO.File]::WriteAllText($tempSys, $sysPrompt, [System.Text.Encoding]::UTF8)

$customNodeScript = @"
const fs = require('fs');
const https = require('https');

const logPath = process.env.LOCALAPPDATA + '\\LeeExcel\\WebView2Profile\\EBWebView\\Default\\Local Storage\\leveldb\\000003.log';
const logContent = fs.readFileSync(logPath, 'utf8');
const matches = logContent.match(/\{[^{}]*provider[^{}]*\}/g);
const config = JSON.parse(matches[matches.length - 1]);

const payload = JSON.stringify({
  model: config.model || 'deepseek-flash',
  messages: [
    { role: 'system', content: fs.readFileSync('$($tempSys.Replace('\', '/'))', 'utf8') },
    { role: 'user', content: '$prompt2' }
  ],
  max_tokens: 16384,
  thinking: { type: 'disabled' }
});

const req = https.request('https://api.deepseek.com/v1/chat/completions', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer ' + config.apiKey
  },
  timeout: 180000
}, (res) => {
  let body = '';
  res.on('data', chunk => body += chunk);
  res.on('end', () => {
    const json = JSON.parse(body);
    const content = json.choices?.[0]?.message?.content || '';
    const finish = json.choices?.[0]?.finish_reason;
    const usage = json.usage || {};
    fs.writeFileSync('$($tempOut.Replace('\', '/'))', content, 'utf8');
    fs.writeFileSync('$($tempMeta.Replace('\', '/'))', JSON.stringify({
      finishReason: finish,
      usage: usage,
      contentLength: content.length
    }, null, 2), 'utf8');
    console.log('Finish:', finish, 'ContentLen:', content.length, 'ReasoningTokens:', usage.completion_tokens_details?.reasoning_tokens);
  });
});
req.write(payload);
req.end();
"@

$runnerCjs = Join-Path $testRunDir "call_nothink.cjs"
[System.IO.File]::WriteAllText($runnerCjs, $customNodeScript, [System.Text.Encoding]::UTF8)

Write-Host "-> Requesting LLM with thinking: disabled (全额 Token 留给代码)..." -ForegroundColor Cyan
node $runnerCjs

$rawContent = [System.IO.File]::ReadAllText($tempOut, [System.Text.Encoding]::UTF8)
$meta = (Get-Content $tempMeta -Raw -Encoding UTF8) | ConvertFrom-Json

Write-Host "-> API Result: Finish=$($meta.finishReason), ContentLength=$($rawContent.Length), ReasoningTokens=$($meta.usage.completion_tokens_details.reasoning_tokens)" -ForegroundColor Yellow

function Extract-VbaCode($content) {
    $openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
    $start = $openMatch.Index + $openMatch.Length
    $closeIdx = $content.IndexOf('```', $start)
    if ($closeIdx -lt 0) { return $content.Substring($start).Trim() }
    return $content.Substring($start, $closeIdx - $start).Trim()
}

$code = Extract-VbaCode $rawContent
Write-Host "-> Extracted code length: $($code.Length) chars" -ForegroundColor Green

# 执行宏
$req = @{
    action = "execute_vba"
    code = $code
    prompt = $prompt2
    targetWorkbookName = $wb.Name
    targetWorkbookFullName = $wbPath
    rawModelResponse = $rawContent
}

$resJson = [LeeExcel.NativeBridge]::Dispatch(($req | ConvertTo-Json -Compress), $excel)
$parsed = [LeeExcel.SimpleJson]::ParseFlatObject($resJson)

Write-Host "-> Execute OK: $($parsed['ok']), Phase: $($parsed['executionPhase']), Precheck: $($parsed['precheckStatus'])" -ForegroundColor $(if ($parsed['ok'] -eq 'True') { "Green" } else { "Red" })

if ($parsed['ok'] -eq 'True') {
    $ws = $wb.ActiveSheet
    Write-Host "-> Sheet: $($ws.Name), UsedRange: $($ws.UsedRange.Address($false, $false))" -ForegroundColor Cyan
    for ($r = 1; $r -le [Math]::Min($ws.UsedRange.Rows.Count, 15); $r++) {
        $line = @()
        for ($c = 1; $c -le [Math]::Min($ws.UsedRange.Columns.Count, 10); $c++) {
            $val = $ws.Cells.Item($r, $c).Text
            $line += if ($val) { $val } else { "[空]" }
        }
        Write-Host "   R$r : " ($line -join " | ") -ForegroundColor Gray
    }
}

$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
