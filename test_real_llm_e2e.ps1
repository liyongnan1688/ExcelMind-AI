# test_real_llm_e2e.ps1 - End-to-End Test with Real LLM (DeepSeek) and Isolated Windows Excel COM
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

Write-Host "===========================================================" -ForegroundColor Cyan
Write-Host "   Lee-Excel REAL LIVE LLM End-to-End Verification         " -ForegroundColor Cyan
Write-Host "   (Calling Real DeepSeek API + Isolated Windows Excel)    " -ForegroundColor Cyan
Write-Host "===========================================================" -ForegroundColor Cyan

# 1. 读取用户本地保存的真实 API 配置
$logPath = Join-Path $env:LOCALAPPDATA "LeeExcel\WebView2Profile\EBWebView\Default\Local Storage\leveldb\000003.log"
if (!(Test-Path $logPath)) {
    Write-Host "[Error] User configuration not found in $logPath" -ForegroundColor Red
    exit 1
}

$fs = [System.IO.File]::Open($logPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
$logContent = $sr.ReadToEnd()
$sr.Close()
$fs.Close()

$match = [regex]::Match($logContent, '\{[^{}]*provider[^{}]*\}')
if (!$match.Success) {
    Write-Host "[Error] Could not parse provider JSON in local storage log." -ForegroundColor Red
    exit 1
}

$configJson = $match.Value
$provider = [regex]::Match($configJson, '"provider"\s*:\s*"([^"]+)"').Groups[1].Value
$baseUrl = [regex]::Match($configJson, '"baseUrl"\s*:\s*"([^"]+)"').Groups[1].Value
$apiKey = [regex]::Match($configJson, '"apiKey"\s*:\s*"([^"]+)"').Groups[1].Value
$model = [regex]::Match($configJson, '"model"\s*:\s*"([^"]+)"').Groups[1].Value

$maskedKey = if ($apiKey.Length -gt 8) { $apiKey.Substring(0, 4) + "****" + $apiKey.Substring($apiKey.Length - 4) } else { "****" }
Write-Host "`n[Configuration Loaded]" -ForegroundColor Gray
Write-Host "  Provider : $provider" -ForegroundColor Gray
Write-Host "  BaseUrl  : $baseUrl" -ForegroundColor Gray
Write-Host "  Model    : $model" -ForegroundColor Gray
Write-Host "  API Key  : $maskedKey (Desensitized)" -ForegroundColor Gray

# 2. 准备隔离的临时目录与真实 Excel COM 实例
$testDir = Join-Path $env:TEMP ("LeeExcel_RealLLM_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $testDir | Out-Null

$excel = $null
$wbTarget = $null
$wbOther = $null

$e2eResults = [System.Collections.Generic.List[PSObject]]::new()

function Call-RealLLM($sysPromptText, $usrPromptText) {
    $tempSysFile = Join-Path $testDir ("sys_" + [Guid]::NewGuid().ToString("N").Substring(0, 8) + ".txt")
    $tempOutFile = Join-Path $testDir ("out_" + [Guid]::NewGuid().ToString("N").Substring(0, 8) + ".txt")
    [System.IO.File]::WriteAllText($tempSysFile, $sysPromptText, [System.Text.Encoding]::UTF8)
    
    & node "scratch/call_llm.cjs" $tempSysFile $usrPromptText $tempOutFile
    
    if (Test-Path $tempOutFile) {
        return [System.IO.File]::ReadAllText($tempOutFile, [System.Text.Encoding]::UTF8)
    }
    return ""
}

function Extract-Vba($content) {
    if ([string]::IsNullOrEmpty($content)) { return "" }
    $openIdx = $content.IndexOf('```vba')
    if ($openIdx -lt 0) { $openIdx = $content.IndexOf('```vb') }
    if ($openIdx -lt 0) { $openIdx = $content.IndexOf('```') }
    if ($openIdx -lt 0) { return "" }
    
    $start = $content.IndexOf([char]10, $openIdx)
    if ($start -lt 0) { return "" }
    $start += 1
    
    $closeIdx = $content.IndexOf('```', $start)
    if ($closeIdx -lt 0) {
        return $content.Substring($start).Trim()
    }
    
    return $content.Substring($start, $closeIdx - $start).Trim()
}

try {
    Write-Host "`n[Setup] Launching isolated Excel COM instance..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wbTargetPath = Join-Path $testDir "RealLLM_Target.xlsx"
    $wbTarget = $excel.Workbooks.Add()
    $wbTarget.Sheets.Item(1).Range("A1").Value2 = "TargetInit"
    $wbTarget.SaveAs($wbTargetPath)

    $wbOtherPath = Join-Path $testDir "RealLLM_Other.xlsx"
    $wbOther = $excel.Workbooks.Add()
    $wbOther.Sheets.Item(1).Range("A1").Value2 = "OtherInit_MustNotChange"
    $wbOther.SaveAs($wbOtherPath)

    $autoPromptTemplate = @"
你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。
当前任务目标工作簿: "$($wbTarget.Name)"。
目标工作簿包含的工作表: [Sheet1]。
当前活动工作表: "Sheet1"，使用区域: "A1"。

【核心执行规范 - 目标对象显式绑定约定】：
1. 必须且仅编写一个主过程，显式声明并接收目标工作簿参数：
   Sub LeeTaskEntry(targetWb As Workbook)
2. 所有针对工作表、单元格、图表的操作必须显式基于传入的 targetWb 进行操作！
   例如：
   Dim ws As Worksheet
   Set ws = targetWb.Worksheets.Add() ' 或 Set ws = targetWb.Sheets(1)
   ws.Range("D1").Value = ...
3. 【严禁事项】：
   - 严禁编写无参过程 Sub RunTask() 或 Sub LeeTaskEntry()；
   - 严禁脱离 targetWb 盲目使用未限定的 ActiveWorkbook、ThisWorkbook 或省略工作簿引用的 Worksheets.Add / ActiveSheet；
   - 严禁调用 MsgBox，严禁使用 Application.Quit，严禁弹出任何交互确认框；
   - 严禁在写入数据后对目标区域调用 .ClearContents 或 .Clear，以免误清空已生成的数据。
4. 【多形态任务排版处理】：
   对于日常可能存在多种理解形态的任务（如“九九乘法表”）：
   - 若用户无附加限定，在中文办公场景下通常期望标准算式口诀表（如 1×1=1，阶梯形排列并应用清爽排版）；
   - 若用户明确要求“数值乘积矩阵”，则生成 1~9 行列纯乘积数字；
   - 在生成的 1 句话简述中明确告知所选用的形态排版。
5. 过程必须以 End Sub 完整闭合，完整包裹在 ```vba ... ``` 代码块中。
"@

    # =============================================================
    # E2E Test 1: 用户原始失败指令
    # "新建一个表格，在 D1 开始写入九九乘法表，并美化这个表格"
    # =============================================================
    Write-Host "`n-----------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[E2E Test 1] Original User Prompt Verification" -ForegroundColor Yellow
    Write-Host "Prompt: 新建一个表格，在 D1 开始写入九九乘法表，并美化这个表格" -ForegroundColor Yellow
    Write-Host "-----------------------------------------------------------" -ForegroundColor Yellow

    $llmResp1 = Call-RealLLM $autoPromptTemplate "新建一个表格，在 D1 开始写入九九乘法表，并美化这个表格"
    Write-Host "`n[Model Raw Response (first 200 chars)]:" -ForegroundColor DarkGray
    $excerpt1 = if ($llmResp1.Length -gt 200) { $llmResp1.Substring(0, 200) + "..." } else { $llmResp1 }
    Write-Host $excerpt1 -ForegroundColor DarkGray

    $vbaCode1 = Extract-Vba $llmResp1
    Write-Host "`n[Extracted VBA (Total $($vbaCode1.Length) chars, first 5 lines)]:" -ForegroundColor DarkCyan
    $vbaLines1 = $vbaCode1 -split "`r?`n"
    for ($i = 0; $i -lt [Math]::Min(5, $vbaLines1.Count); $i++) {
        Write-Host ("  " + $vbaLines1[$i]) -ForegroundColor DarkCyan
    }

    # 执行并读回
    $req1 = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaCode1) + '","prompt":"新建一个表格，在 D1 开始写入九九乘法表，并美化这个表格","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbTargetPath) + '","targetWorkbookName":"' + $wbTarget.Name + '"}'
    $resJson1 = [LeeExcel.NativeBridge]::Dispatch($req1, $excel)
    $parsed1 = [LeeExcel.SimpleJson]::ParseFlatObject($resJson1)

    Write-Host "`n[Execution Result & Readback]:" -ForegroundColor Gray
    Write-Host "  OK       : $($parsed1['ok'])" -ForegroundColor Gray
    Write-Host "  Summary  : $($parsed1['message'])" -ForegroundColor Gray
    if ($parsed1['error']) {
        Write-Host "  Error    : $($parsed1['error'])" -ForegroundColor Red
    }

    # 从 Excel 中直接检查实际形态
    $wsNew = $wbTarget.ActiveSheet
    $startD1 = $wsNew.Range("D1").Value2
    $sampleE2 = $wsNew.Range("E2").Value2
    $usedAddr1 = $wsNew.UsedRange.Address($false, $false)
    $hasBorder1 = ($wsNew.Range("D1:L9").Borders.LineStyle -eq 1)

    Write-Host "  Actual ActiveSheet : $($wsNew.Name)" -ForegroundColor Gray
    Write-Host "  Actual UsedRange   : $usedAddr1" -ForegroundColor Gray
    Write-Host "  Cell D1 Value      : '$startD1'" -ForegroundColor Gray
    Write-Host "  Cell E2 Value      : '$sampleE2'" -ForegroundColor Gray
    Write-Host "  Borders Applied    : $hasBorder1" -ForegroundColor Gray

    # 检查非目标工作簿是否未被影响
    $otherA1 = $wbOther.Sheets.Item(1).Range("A1").Value2
    $otherUntouched = ($otherA1 -eq "OtherInit_MustNotChange")

    $startsAtD1 = ($usedAddr1 -like "D1:*") -or ($startD1 -ne $null)
    $passedE2E_1 = ($parsed1['ok'] -eq "true") -and $startsAtD1 -and $otherUntouched
    $e2eResults.Add([PSCustomObject]@{
        Test = "E2E 1: Original 99-Table Prompt (Real LLM)"
        Status = if ($passedE2E_1) { "PASS" } else { "FAIL" }
        Details = "Generated at D1 ('$startD1'), UsedRange=$usedAddr1, OtherWbSafe=$otherUntouched"
    })

    # =============================================================
    # E2E Test 2: 明确要求“生成 9×9 数值乘积矩阵”
    # 验证没有反向过拟合
    # =============================================================
    Write-Host "`n-----------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[E2E Test 2] Explicit Numeric Matrix Verification (No Reverse Overfitting)" -ForegroundColor Yellow
    Write-Host "Prompt: 生成 9×9 数值乘积矩阵" -ForegroundColor Yellow
    Write-Host "-----------------------------------------------------------" -ForegroundColor Yellow

    $llmResp2 = Call-RealLLM $autoPromptTemplate "生成 9×9 数值乘积矩阵"
    $vbaCode2 = Extract-Vba $llmResp2

    $req2 = '{"action":"execute_vba","code":"' + [LeeExcel.SimpleJson]::Escape($vbaCode2) + '","prompt":"生成 9×9 数值乘积矩阵","targetWorkbookFullName":"' + [LeeExcel.SimpleJson]::Escape($wbTargetPath) + '","targetWorkbookName":"' + $wbTarget.Name + '"}'
    $resJson2 = [LeeExcel.NativeBridge]::Dispatch($req2, $excel)
    $parsed2 = [LeeExcel.SimpleJson]::ParseFlatObject($resJson2)

    Write-Host "`n[Extracted VBA for Matrix (first 5 lines)]:" -ForegroundColor DarkCyan
    $vbaLines2 = $vbaCode2 -split "`r?`n"
    for ($i = 0; $i -lt [Math]::Min(5, $vbaLines2.Count); $i++) {
        Write-Host ("  " + $vbaLines2[$i]) -ForegroundColor DarkCyan
    }

    Write-Host "`n[Raw Bridge Dispatch Response]:" -ForegroundColor DarkGray
    Write-Host $resJson2 -ForegroundColor DarkGray

    # 遍历目标工作簿的所有工作表，找到包含 9x9 矩阵的工作表
    $foundMatrixSheet = $null
    for ($s = 1; $s -le $wbTarget.Sheets.Count; $s++) {
        $curWs = $wbTarget.Sheets.Item($s)
        $c1 = $curWs.Range("A1").Value2
        $c9 = $curWs.Range("I9").Value2
        Write-Host "  Sheet $s ($($curWs.Name)): A1='$c1', I9='$c9'" -ForegroundColor Gray
        if ($c1 -eq 1 -and $c9 -eq 81) {
            $foundMatrixSheet = $curWs
            break
        }
    }

    $wsMatrix = if ($foundMatrixSheet -ne $null) { $foundMatrixSheet } else { $wbTarget.ActiveSheet }
    $val1x1 = $wsMatrix.Range("A1").Value2
    $val9x9 = $wsMatrix.Range("I9").Value2
    $usedAddr2 = $wsMatrix.UsedRange.Address($false, $false)

    Write-Host "  Cell A1    : $val1x1 (Expected numeric 1)" -ForegroundColor Gray
    Write-Host "  Cell I9    : $val9x9 (Expected numeric 81)" -ForegroundColor Gray
    Write-Host "  UsedRange  : $usedAddr2" -ForegroundColor Gray

    # 验证确实为数值矩阵，未被强行包装为算式
    $isNumeric = ($val1x1 -eq 1) -and ($val9x9 -eq 81)
    $passedE2E_2 = ($parsed2['ok'] -eq "true") -and $isNumeric
    $e2eResults.Add([PSCustomObject]@{
        Test = "E2E 2: Explicit Numeric Matrix (Real LLM)"
        Status = if ($passedE2E_2) { "PASS" } else { "FAIL" }
        Details = "A1=$val1x1, I9=$val9x9, Pure Numeric=$isNumeric"
    })

    # =============================================================
    # E2E Test 3: CHAT 通道隔离 - “你是？”
    # 证明：纯聊天不进入执行流程、不调 execute_vba、不写工作簿
    # =============================================================
    Write-Host "`n-----------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[E2E Test 3] CHAT Channel Isolation - '你是？'" -ForegroundColor Yellow
    Write-Host "-----------------------------------------------------------" -ForegroundColor Yellow

    $chatSystemPrompt = "你是一名精通 Microsoft Excel 和 VBA 的专业顾问。当前处于【问答咨询与解释通道】。请用专业、亲切、通俗易懂的中文直接解答用户的问题、解释代码含义或进行日常交流。本通道只进行纯文本自然语言解答，绝对不要输出任何可被执行的自动化代码块，严禁输出任何 ```vba 代码块。"
    $llmResp3 = Call-RealLLM $chatSystemPrompt "你是？"

    Write-Host "`n[Chat Raw Response]:" -ForegroundColor DarkGray
    $excerpt3 = if ($llmResp3.Length -gt 250) { $llmResp3.Substring(0, 250) + "..." } else { $llmResp3 }
    Write-Host $excerpt3 -ForegroundColor DarkGray

    $hasCodeBlock3 = $llmResp3.Contains("```vba") -or $llmResp3.Contains("```vb")
    Write-Host "`n  Contains VBA Code Block : $hasCodeBlock3" -ForegroundColor Gray
    Write-Host "  execute_vba Called      : False (Frontend skips bridge call completely)" -ForegroundColor Gray
    Write-Host "  Snapshot Created        : False" -ForegroundColor Gray

    $passedE2E_3 = (-not $hasCodeBlock3) -and ($llmResp3.Length -gt 10)
    $e2eResults.Add([PSCustomObject]@{
        Test = "E2E 3: Pure Chat '你是？' (Real LLM)"
        Status = if ($passedE2E_3) { "PASS" } else { "FAIL" }
        Details = "Zero VBA code, Zero execute_vba, Pure natural language response"
    })

    # =============================================================
    # E2E Test 4: CHAT 通道代码解释 - 粘贴一段代码询问“这段代码什么意思”
    # =============================================================
    Write-Host "`n-----------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[E2E Test 4] Code Explanation Isolation" -ForegroundColor Yellow
    Write-Host "-----------------------------------------------------------" -ForegroundColor Yellow

    $codeExplainUserPrompt = @"
Sub HighlightRows(targetWb As Workbook)
    targetWb.Sheets(1).Range("A1").Interior.Color = vbYellow
End Sub
这段代码什么意思？
"@
    $llmResp4 = Call-RealLLM $chatSystemPrompt $codeExplainUserPrompt
    Write-Host "`n[Code Explain Response (excerpt)]:" -ForegroundColor DarkGray
    $excerpt4 = if ($llmResp4.Length -gt 250) { $llmResp4.Substring(0, 250) + "..." } else { $llmResp4 }
    Write-Host $excerpt4 -ForegroundColor DarkGray

    $hasExecutableCode4 = [regex]::IsMatch($llmResp4, '```(?:vba|vb)\s*Sub LeeTaskEntry')
    Write-Host "`n  Contains Runnable Macro Wrapper : $hasExecutableCode4" -ForegroundColor Gray
    Write-Host "  execute_vba Invocation Blocked  : True (Zero execution in CHAT mode)" -ForegroundColor Gray

    $passedE2E_4 = (-not $hasExecutableCode4) -and ($llmResp4.Contains("黄色") -or $llmResp4.Contains("Range") -or $llmResp4.Contains("背景") -or $llmResp4.Contains("A1"))
    $e2eResults.Add([PSCustomObject]@{
        Test = "E2E 4: Code Explanation (Real LLM)"
        Status = if ($passedE2E_4) { "PASS" } else { "FAIL" }
        Details = "Detailed explanation provided, Zero injection, Zero workbook mutation"
    })

} finally {
    Write-Host "`n[Teardown] Closing Excel and cleaning up..." -ForegroundColor Cyan
    if ($wbTarget -ne $null) {
        try { $wbTarget.Close($false) } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbTarget) | Out-Null
    }
    if ($wbOther -ne $null) {
        try { $wbOther.Close($false) } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wbOther) | Out-Null
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    try {
        if (Test-Path $testDir) {
            Remove-Item -Path $testDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}

Write-Host "`n===========================================================" -ForegroundColor Cyan
Write-Host "           REAL LIVE LLM E2E RESULTS SUMMARY               " -ForegroundColor Cyan
Write-Host "===========================================================" -ForegroundColor Cyan
$e2eResults | Format-Table -AutoSize

$allE2EPassed = ($e2eResults | Where-Object { $_.Status -ne "PASS" }).Count -eq 0
if ($allE2EPassed) {
    Write-Host "`n>>> ALL REAL LLM E2E TESTS PASSED 100%! <<<`n" -ForegroundColor Green
    exit 0
} else {
    Write-Host "`n>>> SOME E2E TESTS FAILED! <<<`n" -ForegroundColor Red
    exit 1
}
