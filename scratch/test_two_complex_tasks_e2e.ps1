# scratch/test_two_complex_tasks_e2e.ps1
# 两项真实大模型复杂任务端到端实测脚本
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "       真实大模型双复杂任务端到端实测 (100% 真实模型与代码)      " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

$testRunDir = Join-Path $env:TEMP ("LeeExcel_DualE2E_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $testRunDir | Out-Null
Write-Host "测试根目录: $testRunDir" -ForegroundColor Gray

function Extract-VbaCode($content) {
    if ([string]::IsNullOrEmpty($content)) { 
        return @{ Code = ""; IsTruncated = $false; Error = "Empty content" } 
    }
    $openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
    if (!$openMatch.Success) {
        return @{ Code = ""; IsTruncated = $false; Error = "No code block found" }
    }
    $start = $openMatch.Index + $openMatch.Length
    $closeIdx = $content.IndexOf('```', $start)
    if ($closeIdx -lt 0) {
        return @{ Code = $content.Substring($start).Trim(); IsTruncated = $true; Error = "Truncated code block" }
    }
    $rawCode = $content.Substring($start, $closeIdx - $start).Trim()
    
    $subMatches = [regex]::Matches($rawCode, '(?im)^\s*(?:Public\s+|Private\s+)?Sub\s+[a-zA-Z0-9_\u4e00-\u9fa5]+\s*\(')
    $endSubMatches = [regex]::Matches($rawCode, '(?im)^\s*End\s+Sub\b')
    $fnMatches = [regex]::Matches($rawCode, '(?im)^\s*(?:Public\s+|Private\s+)?Function\s+[a-zA-Z0-9_\u4e00-\u9fa5]+\s*\(')
    $endFnMatches = [regex]::Matches($rawCode, '(?im)^\s*End\s+Function\b')

    if ($subMatches.Count -eq 0) {
        return @{ Code = $rawCode; IsTruncated = $false; Error = "No Sub declaration" }
    }
    if ($subMatches.Count -gt $endSubMatches.Count -or $fnMatches.Count -gt $endFnMatches.Count) {
        return @{ Code = $rawCode; IsTruncated = $true; Error = "Sub/Function count mismatch with End Sub/Function (truncated)" }
    }
    return @{ Code = $rawCode; IsTruncated = $false; Error = $null }
}

function Run-Single-E2E-Task($taskId, $taskTitle, $userPrompt, $wbFileName) {
    Write-Host "`n-----------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "【$taskId】: $taskTitle" -ForegroundColor Yellow
    Write-Host "用户自然语言需求: $userPrompt" -ForegroundColor Gray
    Write-Host "-----------------------------------------------------------------" -ForegroundColor Yellow

    $wbPath = Join-Path $testRunDir $wbFileName
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $wb.SaveAs($wbPath)
    Write-Host "已创建独立空白测试工作簿: $($wb.Name)" -ForegroundColor Gray

    # 1. 组装系统提示词
    $sysLines = @(
        '你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。',
        "当前任务目标工作簿: `"$($wb.Name)`"。",
        "目标工作簿包含的工作表: [Sheet1]。",
        '当前活动工作表: "Sheet1"，使用区域: "空"。',
        '',
        '【核心执行协议与规范】：',
        '1. 根据用户的自然语言需求，自主决定最合适的高效实现方案（可自由使用循环、数组、公式、格式、图表、筛选、数据透视表及辅助过程等，不受限固定模板与行数）。',
        '2. 主过程可以声明接收目标工作簿参数（如 Sub Main(targetWb As Workbook)），也可以编写无参主过程（如 Sub Main()）；允许定义多个辅助过程与函数。',
        '3. 代码必须是完整可编译运行的标准 VBA，语法严格遵循 VB6/VBA 规范，包裹在单个 ```vba ... ``` 代码块中，以 End Sub 正常闭合。',
        '4. 【安全约束】：严禁调用 MsgBox、Application.Quit 或弹出阻塞式交互确认框。',
        '5. 【结构与输出】：直接输出完整可执行的标准 VBA 代码，包裹在单个 ```vba ... ``` 代码块中，在代码块前后仅提供简明扼要的说明。'
    )
    $sysPrompt = ($sysLines -join "`r`n")
    $tempSys = Join-Path $testRunDir "$($taskId)_sys.txt"
    $tempOut = Join-Path $testRunDir "$($taskId)_out.txt"
    $tempMeta = Join-Path $testRunDir "$($taskId)_meta.json"
    [System.IO.File]::WriteAllText($tempSys, $sysPrompt, [System.Text.Encoding]::UTF8)

    Write-Host "-> 正在向大模型发起真实请求 (包含 16384 token 预算与截断重试协议)..." -ForegroundColor Cyan
    $swApi = [System.Diagnostics.Stopwatch]::StartNew()
    & node "scratch/call_llm.cjs" $tempSys $userPrompt $tempOut $tempMeta
    $swApi.Stop()

    $meta = $null
    if (Test-Path $tempMeta) {
        $meta = (Get-Content $tempMeta -Raw -Encoding UTF8) | ConvertFrom-Json
    }

    $rawContent = ""
    if (Test-Path $tempOut) {
        $rawContent = [System.IO.File]::ReadAllText($tempOut, [System.Text.Encoding]::UTF8)
    }

    $extracted = Extract-VbaCode $rawContent

    $report = @{
        TaskId = $taskId
        TaskTitle = $taskTitle
        UserPrompt = $userPrompt
        WorkbookName = $wb.Name
        ApiDurationMs = $swApi.ElapsedMilliseconds
        FinishReason = $meta.finishReason
        RetryCount = $meta.retryCount
        ReasoningTokens = $meta.reasoningTokens
        CompletionTokens = $meta.completionTokens
        TotalTokens = $meta.totalTokens
        RawResponseLength = $rawContent.Length
        ExtractedCodeLength = $extracted.Code.Length
        IsTruncated = $extracted.IsTruncated
        ExtractionError = $extracted.Error
        OriginalHash = ""
        ExecutedHash = ""
        IsSourceIdentical = $false
        PrecheckStatus = ""
        ExecutionPhase = ""
        ExecutionSuccess = $false
        ExecutionMessage = ""
        ExecutionError = ""
        UsedRange = ""
        TargetSheetName = ""
        SampleGrid = @()
        ResidualModules = -1
    }

    if ($extracted.Error -or $extracted.IsTruncated) {
        Write-Host "❌ 代码提取或截断失败: $($extracted.Error)" -ForegroundColor Red
        $wb.Close($false)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $report
    }

    Write-Host "-> 成功提取完整闭合 VBA 代码 ($($extracted.Code.Length) 字符, SHA256 待算)" -ForegroundColor Green

    # 2. 调用 NativeBridge 执行
    $req = @{
        action = "execute_vba"
        code = $extracted.Code
        prompt = $userPrompt
        targetWorkbookName = $wb.Name
        targetWorkbookFullName = $wbPath
        rawModelResponse = $rawContent
    }

    Write-Host "-> 正在注入目标工作簿并受控执行宏..." -ForegroundColor Cyan
    $resJson = [LeeExcel.NativeBridge]::Dispatch(($req | ConvertTo-Json -Compress), $excel)
    $parsed = [LeeExcel.SimpleJson]::ParseFlatObject($resJson)

    $report.ExecutionSuccess = ($parsed['ok'] -eq "True")
    $report.ExecutionMessage = $parsed['message']
    $report.ExecutionError = $parsed['error']
    $report.OriginalHash = $parsed['originalCodeHash']
    $report.ExecutedHash = $parsed['executedCodeHash']
    $report.IsSourceIdentical = ($parsed['isSourceIdentical'] -eq "True")
    $report.PrecheckStatus = $parsed['precheckStatus']
    $report.ExecutionPhase = $parsed['executionPhase']

    Write-Host "-> 执行状态: ok=$($report.ExecutionSuccess), Phase=$($report.ExecutionPhase), Precheck=$($report.PrecheckStatus)" -ForegroundColor $(if ($report.ExecutionSuccess) { "Green" } else { "Red" })

    # 3. 详细读回验证
    if ($report.ExecutionSuccess) {
        $ws = $wb.ActiveSheet
        $report.TargetSheetName = $ws.Name
        $report.UsedRange = $ws.UsedRange.Address($false, $false)
        Write-Host "-> 活动表: $($ws.Name), 实际生效区域: $($report.UsedRange)" -ForegroundColor Cyan

        $maxR = [Math]::Min($ws.UsedRange.Rows.Count, 16)
        $maxC = [Math]::Min($ws.UsedRange.Columns.Count, 10)
        for ($r = 1; $r -le $maxR; $r++) {
            $line = @()
            for ($c = 1; $c -le $maxC; $c++) {
                $val = $ws.Cells.Item($r, $c).Text
                $line += if ($val) { $val } else { "[空]" }
            }
            $rowStr = "R$r : " + ($line -join " | ")
            Write-Host "   $rowStr" -ForegroundColor Gray
            $report.SampleGrid += $rowStr
        }
    }

    # 4. 检查模块清理
    $wb.Save()
    $wb.Close($false)

    $wbCheck = $excel.Workbooks.Open($wbPath)
    $modCount = 0
    try {
        foreach ($comp in $wbCheck.VBProject.VBComponents) {
            if ($comp.Type -eq 1) { $modCount++ }
        }
    } catch {}
    $wbCheck.Close($false)
    $report.ResidualModules = $modCount
    Write-Host "-> 临时模块清理检验: 残留标准模块数 = $modCount (预期 0)" -ForegroundColor $(if ($modCount -eq 0) { "Green" } else { "Red" })

    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

    $resPath = Join-Path $testRunDir "$($taskId)_result.json"
    [System.IO.File]::WriteAllText($resPath, ($report | ConvertTo-Json -Depth 5), [System.Text.Encoding]::UTF8)

    # 保存代码
    $vbaPath = Join-Path $testRunDir "$($taskId)_executed.vba"
    [System.IO.File]::WriteAllText($vbaPath, $extracted.Code, [System.Text.Encoding]::UTF8)

    return $report
}

# ===== 执行两项任务 =====

# 任务 1：原始复杂工时考勤看板
$prompt1 = "制作一个项目考勤与工时统计看板：新建一张表，录入5名员工5天的工作时长（正常为8小时），自动用浅红底色高亮超过9小时的加班工时，计算每人的总工时和出勤天数，并在顶部生成带大号字体的项目关键指标卡片"
$res1 = Run-Single-E2E-Task "Task1_AttendanceDashboard" "原始复杂工时考勤看板任务" $prompt1 "AttendanceDashboard.xlsx"

# 任务 2：不同领域多步骤任务（多仓库库存调拨与警戒看板）
$prompt2 = "制作一个多仓库库存调拨与警戒看板：新建一张表，录入4个仓库（北京仓、上海仓、广州仓、成都仓）中5类电子元器件的当前库存与安全库存警戒线；自动计算各仓库总库存量及各品类的全网缺货差额；用黄色高亮低于安全库存的缺货单元格；在顶部生成3个大号指标卡片（总库存件数、总缺货品次、最低警戒仓库）"
$res2 = Run-Single-E2E-Task "Task2_WarehouseInventory" "不同领域多步骤多仓库库存警戒看板任务" $prompt2 "WarehouseInventory.xlsx"

Write-Host "`n=================================================================" -ForegroundColor Cyan
Write-Host "                   双复杂任务端到端实测汇总                       " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "Task 1 (考勤看板): OK=$($res1.ExecutionSuccess), Phase=$($res1.ExecutionPhase), Range=$($res1.UsedRange), Tokens=$($res1.TotalTokens), Retries=$($res1.RetryCount)"
Write-Host "Task 2 (库存看板): OK=$($res2.ExecutionSuccess), Phase=$($res2.ExecutionPhase), Range=$($res2.UsedRange), Tokens=$($res2.TotalTokens), Retries=$($res2.RetryCount)"
Write-Host "结果已存盘至: $testRunDir"
