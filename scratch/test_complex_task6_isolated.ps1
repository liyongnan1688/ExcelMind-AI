# scratch/test_complex_task6_isolated.ps1
# 单独重试 Task 6 (复杂工时看板多步骤任务) 真实端到端测试
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "   TASK 6 (复杂工时看板) 真实模型全新隔离重试                   " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

# 1. 隔离目录与目标工作簿
$testDir = Join-Path $env:TEMP ("LeeExcel_Task6_Isolated_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $testDir | Out-Null
$wbPath = Join-Path $testDir "ComplexDashboardTest.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$wb.SaveAs($wbPath)

# 2. 系统提示词与用户需求
$sysLines = @(
    '你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。',
    "当前任务目标工作簿: `"$($wb.Name)`"。",
    "目标工作簿包含的工作表: [Sheet1]。",
    '当前活动工作表: "Sheet1"，使用区域: "空"。',
    '',
    '【核心执行协议与规范】：',
    '1. 根据用户的自然语言需求，自主决定最合适的高效实现方案（可自由使用循环、数组、公式、格式、图表、筛选、数据透视表及辅助过程等，不受限固定模板与行数）。',
    '2. 主过程可以声明接收目标工作簿参数（如 Sub Main(targetWb As Workbook)），也可以编写无参主过程（如 Sub Main()）；允许定义多个辅助过程与函数。',
    '3. 代码必须是完整可编译运行的标准 VBA，语法严格遵循 VB6/VBA 规范（仔细检查括号与属性调用的位置如 ws.Columns(1).ColumnWidth，提前退出请使用 Exit Sub/Function，禁止书写非法的自定义 End 标签 如 End CleanExit 等），包裹在单个 ```vba ... ``` 代码块中，以 End Sub 正常闭合。',
    '4. 【安全约束】：严禁调用 MsgBox、Application.Quit 或弹出阻塞式交互确认框。',
    '5. 【结构与输出】：直接输出完整可执行的标准 VBA 代码，包裹在单个 ```vba ... ``` 代码块中，在代码块前后仅提供简明扼要的说明，避免冗长说明以确保代码完整不被截断。'
)
$sysPrompt = ($sysLines -join "`r`n")
$userPrompt = "制作一个项目考勤与工时统计看板：新建一张表，录入5名员工5天的工作时长（正常为8小时），自动用浅红底色高亮超过9小时的加班工时，计算每人的总工时和出勤天数，并在顶部生成带大号字体的项目关键指标卡片"

# 3. 调用真实大模型
$tempSys = Join-Path $testDir "sys.txt"
$tempOut = Join-Path $testDir "out.txt"
[System.IO.File]::WriteAllText($tempSys, $sysPrompt, [System.Text.Encoding]::UTF8)

Write-Host "-> Requesting DeepSeek API for Task 6..." -ForegroundColor Yellow
$swApi = [System.Diagnostics.Stopwatch]::StartNew()
& node "scratch/call_llm.cjs" $tempSys $userPrompt $tempOut
$swApi.Stop()
Write-Host "API call completed in $($swApi.ElapsedMilliseconds) ms" -ForegroundColor Gray

$rawResponse = ""
if (Test-Path $tempOut) {
    $rawResponse = [System.IO.File]::ReadAllText($tempOut, [System.Text.Encoding]::UTF8)
}

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

$extracted = Extract-VbaCode $rawResponse

$task6Result = @{
    Prompt = $userPrompt
    ApiDurationMs = $swApi.ElapsedMilliseconds
    RawResponseLength = $rawResponse.Length
    ExtractionError = $extracted.Error
    IsTruncated = $extracted.IsTruncated
    ExtractedCodeLength = $extracted.Code.Length
    ExecutionSuccess = $false
    ExecutionMessage = ""
    UsedRange = ""
    ResidualModules = -1
    SampleValues = @()
}

if ($extracted.Error -or $extracted.IsTruncated) {
    Write-Host "[Task 6 Extraction Failed]: $($extracted.Error)" -ForegroundColor Red
} else {
    Write-Host "-> Extracted valid VBA ($($extracted.Code.Length) chars)" -ForegroundColor Green
    
    $req = @{
        action = "execute_vba"
        code = $extracted.Code
        prompt = $userPrompt
        targetWorkbookName = $wb.Name
        targetWorkbookFullName = $wbPath
        rawModelResponse = $rawResponse
    }
    $resJson = [LeeExcel.NativeBridge]::Dispatch(($req | ConvertTo-Json -Compress), $excel)
    $parsed = [LeeExcel.SimpleJson]::ParseFlatObject($resJson)

    $task6Result.ExecutionSuccess = ($parsed['ok'] -eq "True")
    $task6Result.ExecutionMessage = $parsed['message']
    $task6Result.OriginalHash = $parsed['originalCodeHash']
    $task6Result.ExecutedHash = $parsed['executedCodeHash']
    $task6Result.IsSourceIdentical = ($parsed['isSourceIdentical'] -eq "True")
    $task6Result.OriginalVba = $extracted.Code
    $task6Result.ExecutedVba = $parsed['executedVbaCode']

    Write-Host "-> Execution Result: OK = $($task6Result.ExecutionSuccess), Message = $($parsed['message'])" -ForegroundColor $(if ($task6Result.ExecutionSuccess) { "Green" } else { "Red" })

    if ($task6Result.ExecutionSuccess) {
        $ws = $wb.ActiveSheet
        $task6Result.UsedRange = $ws.UsedRange.Address($false, $false)
        Write-Host "-> UsedRange: $($task6Result.UsedRange)" -ForegroundColor Cyan
        
        for ($r = 1; $r -le [Math]::Min($ws.UsedRange.Rows.Count, 15); $r++) {
            $line = @()
            for ($c = 1; $c -le [Math]::Min($ws.UsedRange.Columns.Count, 10); $c++) {
                $line += $ws.Cells.Item($r, $c).Text
            }
            $rowStr = "Row $r : " + ($line -join " | ")
            Write-Host $rowStr -ForegroundColor Gray
            $task6Result.SampleValues += $rowStr
        }
    }
}

# 保存工作簿并核验临时模块残留
$wb.Save()
$wb.Close($false)

$wbRe = $excel.Workbooks.Open($wbPath)
$resMod = 0
try {
    foreach ($comp in $wbRe.VBProject.VBComponents) {
        if ($comp.Type -eq 1) { $resMod++ }
    }
} catch {}
$wbRe.Close($false)
$task6Result.ResidualModules = $resMod

$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

$task6ResultPath = Join-Path $testDir "task6_isolated_summary.json"
$summaryJson = ($task6Result | ConvertTo-Json -Depth 6)
[System.IO.File]::WriteAllText($task6ResultPath, $summaryJson, [System.Text.Encoding]::UTF8)

Write-Host "`nTask 6 Isolated Retrial Complete! Summary: $task6ResultPath" -ForegroundColor Cyan
