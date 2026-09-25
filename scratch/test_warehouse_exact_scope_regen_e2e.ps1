# scratch/test_warehouse_exact_scope_regen_e2e.ps1
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$runId = [Guid]::NewGuid().ToString("N").Substring(0, 8)
$testDir = Join-Path $env:TEMP ("LeeExcel_WarehouseRegen_" + $runId)
New-Item -ItemType Directory -Path $testDir | Out-Null

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host " 多仓库库存看板：第一版全表失控拦截 -> 携带诊断请模型重新生成 -> 真实执行与读回 " -ForegroundColor Cyan
Write-Host " 测试工作目录: $testDir" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan

# 1. 加载上一轮真实模型输出的、包含 ws.Cells.Borders.LineStyle = xlContinuous 的第一版原始代码
$v1Path = "C:\Users\35651\AppData\Local\Temp\LeeExcel_Task2_NoThink_0c6f523e\pure_vba.bas"
if (!(Test-Path $v1Path)) {
    # 备用路径
    $v1Path = (Get-ChildItem $env:TEMP -Filter "LeeExcel_Task2_NoThink_*" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName + "\pure_vba.bas"
}
$v1Code = [System.IO.File]::ReadAllText($v1Path, [System.Text.Encoding]::UTF8)
$v1Hash = [LeeExcel.VbaRunner]::ComputeSha256($v1Code)

Write-Host "`n1. 审计第一版真实模型生成代码：" -ForegroundColor Yellow
Write-Host "    - 源码路径: $v1Path"
Write-Host "    - 源码字符数: $($v1Code.Length)"
Write-Host "    - 第一版 SHA256: $v1Hash"

# 2. 执行前影响范围检查
$scopeRegex = '(?:(?:\bws\b|\bActiveSheet\b|\bWorksheets\([^)]+\)|(?<!\w))\s*\.\s*Cells|(?<!\.)\bCells)\s*\.\s*(?:Borders|Interior|FormatConditions|ClearFormats)\b'
$match = [regex]::Match($v1Code, $scopeRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

Write-Host "`n2. 执行前影响范围检查 (checkVbaScopeRisk)：" -ForegroundColor Yellow
if ($match.Success) {
    Write-Host "    - 【高危拦截】命中全表失控语句: '$($match.Value)'" -ForegroundColor Red
    Write-Host "    - 风险说明: Excel 单表包含逾 171 亿单元格，对 Cells 直接设置边框将耗尽内存并引发深度无响应挂起。" -ForegroundColor Yellow
    Write-Host "    - 拦截决策: 阻止第一版源码直接注入运行，绝不静默暗改源码，携带原需求与范围诊断向模型发起重新生成！" -ForegroundColor Cyan
} else {
    Write-Host "    - 未检测到全表失控语句" -ForegroundColor Green
}

# 3. 携带原需求与范围诊断，请模型重新生成一份完整标准 VBA 代码
Write-Host "`n3. 正在请大模型重新生成一份完整 VBA (保持算法自由，纠偏失控范围)..." -ForegroundColor Cyan
$origPrompt = "制作一个多仓库库存调拨与警戒看板：新建一张表，录入4个仓库（北京仓、上海仓、广州仓、成都仓）中5类电子元器件的当前库存与安全库存警戒线；自动计算各仓库总库存量及各品类的全网缺货差额；用黄色高亮低于安全库存的缺货单元格；在顶部生成3个大号指标卡片（总库存件数、总缺货品次、最低警戒仓库）"

$fixPrompt = "$origPrompt`n`n【重要范围规范修正】：`n上一版代码包含【$($match.Value)】。Excel包含171亿个单元格，直接对全表Cells批量设置边框或底色会导致Excel进程长时间无响应挂起。`n请保留您的所有业务逻辑、公式计算、指标卡片与表格设计，但务必将边框、背景色等格式化语句限定在实际业务数据区域（例如使用 ws.Range(...) 或 Range(ws.Cells(...), ws.Cells(...))），请重新输出一份完整可运行的标准 VBA 代码。"

$sysPrompt = "你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。请直接输出完整可执行的标准 VBA 代码，包裹在单个 ```vba ... ``` 代码块中，以 End Sub 正常闭合。严禁调用 MsgBox、Application.Quit。"

$sysFile = Join-Path $testDir "sys.txt"
$outFile = Join-Path $testDir "out.txt"
$metaFile = Join-Path $testDir "meta.json"
[System.IO.File]::WriteAllText($sysFile, $sysPrompt, [System.Text.Encoding]::UTF8)

# 写入调用脚本并执行
$nodeScript = @"
const fs = require('fs');
const https = require('https');
const crypto = require('crypto');

const logPath = process.env.LOCALAPPDATA + '\\LeeExcel\\WebView2Profile\\EBWebView\\Default\\Local Storage\\leveldb\\000003.log';
const logContent = fs.readFileSync(logPath, 'utf8');
const matches = logContent.match(/\{[^{}]*provider[^{}]*\}/g);
const config = JSON.parse(matches[matches.length - 1]);

const payload = JSON.stringify({
  model: config.model || 'deepseek-flash',
  messages: [
    { role: 'system', content: fs.readFileSync('$($sysFile.Replace('\', '/'))', 'utf8') },
    { role: 'assistant', content: '```vba\n' + fs.readFileSync('$($v1Path.Replace('\', '/'))', 'utf8') + '\n```' },
    { role: 'user', content: fs.readFileSync('$((Join-Path $testDir "fix_prompt.txt").Replace('\', '/'))', 'utf8') }
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
    fs.writeFileSync('$($outFile.Replace('\', '/'))', content, 'utf8');
    fs.writeFileSync('$($metaFile.Replace('\', '/'))', JSON.stringify({
      finishReason: finish,
      usage: usage,
      contentLength: content.length
    }, null, 2), 'utf8');
    console.log('Regen finish:', finish, 'Len:', content.length, 'Usage:', JSON.stringify(usage));
  });
});
req.write(payload);
req.end();
"@

[System.IO.File]::WriteAllText((Join-Path $testDir "fix_prompt.txt"), $fixPrompt, [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText((Join-Path $testDir "call_regen.cjs"), $nodeScript, [System.Text.Encoding]::UTF8)

node (Join-Path $testDir "call_regen.cjs")

$regenRaw = [System.IO.File]::ReadAllText($outFile, [System.Text.Encoding]::UTF8)
$meta = (Get-Content $metaFile -Raw -Encoding UTF8) | ConvertFrom-Json

function Extract-VbaCode($content) {
    $openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
    $start = $openMatch.Index + $openMatch.Length
    $closeIdx = $content.IndexOf('```', $start)
    if ($closeIdx -lt 0) { return $content.Substring($start).Trim() }
    return $content.Substring($start, $closeIdx - $start).Trim()
}

$v2Code = Extract-VbaCode $regenRaw
$v2Hash = [LeeExcel.VbaRunner]::ComputeSha256($v2Code)
[System.IO.File]::WriteAllText((Join-Path $testDir "Version2_Regenerated.vba"), $v2Code, [System.Text.Encoding]::UTF8)

Write-Host "`n4. 检验重新生成的第二版代码：" -ForegroundColor Yellow
Write-Host "    - 第二版代码字符数: $($v2Code.Length)"
Write-Host "    - 第二版 SHA256: $v2Hash"
Write-Host "    - 第二版与第一版哈希是否不同: $($v1Hash -ne $v2Hash)" -ForegroundColor $(if ($v1Hash -ne $v2Hash) { "Green" } else { "Red" })

# 再次进行影响范围检查
$match2 = [regex]::Match($v2Code, $scopeRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
Write-Host "    - 第二版是否仍包含全表失控语句: $($match2.Success)" -ForegroundColor $(if (!$match2.Success) { "Green" } else { "Red" })
if ($match2.Success) {
    Write-Host "    - 依然命中: $($match2.Value)" -ForegroundColor Red
} else {
    Write-Host "    - 【范围纠偏成功】格式化语句已成功限定在数据 Range 内！" -ForegroundColor Green
}

# 5. 在新独立工作簿中真正执行第二版代码
Write-Host "`n5. 向全新独立工作簿注入并执行第二版完整 VBA..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wbPath = Join-Path $testDir "Warehouse_Regen_Executed.xlsx"
    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($wbPath)

    $execReq = @{
        action = "execute_vba"
        code = $v2Code
        prompt = $origPrompt
        targetWorkbookName = $wb.Name
        targetWorkbookFullName = $wbPath
        rawModelResponse = $regenRaw
    }

    $resJson = [LeeExcel.NativeBridge]::Dispatch(($execReq | ConvertTo-Json -Compress), $excel)
    $parsed = [LeeExcel.SimpleJson]::ParseFlatObject($resJson)

    Write-Host "`n6. 真实执行结果与响应状态：" -ForegroundColor Yellow
    Write-Host "    - ok: $($parsed['ok'])" -ForegroundColor $(if ($parsed['ok'] -eq 'True') { "Green" } else { "Red" })
    Write-Host "    - executionPhase: $($parsed['executionPhase'])"
    Write-Host "    - precheckStatus: $($parsed['precheckStatus'])"
    Write-Host "    - elapsedMs: $($parsed['elapsedMs']) ms"
    Write-Host "    - 是否发生 Excel 卡死或无响应挂起: $(if ($parsed['executionPhase'] -like '*hang*') { 'True (挂起)' } else { 'False (正常响应)' })" -ForegroundColor $(if ($parsed['executionPhase'] -like '*hang*') { "Red" } else { "Green" })

    # 7. 从 Excel 读回数据并验证
    Write-Host "`n7. Excel 真实写后读回数据核验：" -ForegroundColor Yellow
    foreach ($sheet in $wb.Worksheets) {
        Write-Host "    工作表: '$($sheet.Name)', 使用区域: '$($sheet.UsedRange.Address($false, $false))' (共 $($sheet.UsedRange.Rows.Count) 行 $($sheet.UsedRange.Columns.Count) 列)"
        for ($r = 1; $r -le [Math]::Min($sheet.UsedRange.Rows.Count, 15); $r++) {
            $rowVals = @()
            for ($c = 1; $c -le [Math]::Min($sheet.UsedRange.Columns.Count, 8); $c++) {
                $t = $sheet.Cells.Item($r, $c).Text
                $rowVals += [string]::Format("{0,-12}", $t)
            }
            Write-Host ("      行 {0,2}: {1}" -f $r, ($rowVals -join " | "))
        }
    }

    $wb.Save()
    Write-Host "`n测试工作簿已保存: $wbPath" -ForegroundColor Green

} finally {
    try {
        $excel.Workbooks | ForEach-Object { $_.Close($false) }
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch { }
}
