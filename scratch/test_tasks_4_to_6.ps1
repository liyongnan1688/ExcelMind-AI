# scratch/test_tasks_4_to_6.ps1
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$logPath = Join-Path $env:LOCALAPPDATA "LeeExcel\WebView2Profile\EBWebView\Default\Local Storage\leveldb\000003.log"
$fs = [System.IO.File]::Open($logPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
$logContent = $sr.ReadToEnd()
$sr.Close()
$fs.Close()

$match = [regex]::Match($logContent, '\{[^{}]*provider[^{}]*\}')
$configJson = $match.Value
$apiKey = [regex]::Match($configJson, '"apiKey"\s*:\s*"([^"]+)"').Groups[1].Value
$model = [regex]::Match($configJson, '"model"\s*:\s*"([^"]+)"').Groups[1].Value

$testDir = Join-Path $env:TEMP ("LeeExcel_Tasks4to6_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $testDir | Out-Null

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
    
    $hasSubDecl = [regex]::IsMatch($rawCode, '(?i)(?:Public\s+|Private\s+)?Sub\s+[a-zA-Z0-9_\u4e00-\u9fa5]+\s*\(')
    $hasEndSub = [regex]::IsMatch($rawCode, '(?i)End\s+Sub')
    if (!$hasSubDecl) {
        return @{ Code = $rawCode; IsTruncated = $false; Error = "No Sub declaration" }
    }
    if (!$hasEndSub) {
        return @{ Code = $rawCode; IsTruncated = $true; Error = "Missing End Sub (truncated)" }
    }
    return @{ Code = $rawCode; IsTruncated = $false; Error = $null }
}

function Build-GenericSysPrompt($wbName, $sheets, $activeSheet, $usedRange) {
    $lines = @(
        '你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。',
        "当前任务目标工作簿: `"$wbName`"。",
        "目标工作簿包含的工作表: [$($sheets -join ', ')]。",
        "当前活动工作表: `"$activeSheet`"，使用区域: `"$usedRange`"。",
        '',
        '【核心执行协议与规范】：',
        '1. 根据用户的自然语言需求，自主决定最合适的高效实现方案（可自由使用循环、数组、公式、格式、图表、筛选、数据透视表及辅助过程等，不受限固定模板与行数）。',
        '2. 主过程可以声明接收目标工作簿参数（如 Sub Main(targetWb As Workbook)），也可以编写无参主过程（如 Sub Main()）；允许定义多个辅助过程与函数。',
        '3. 代码必须是完整可编译运行的标准 VBA，语法严格遵循 VB6/VBA 规范（仔细检查括号与属性调用的位置如 ws.Columns(1).ColumnWidth，提前退出请使用 Exit Sub/Function，禁止书写非法的自定义 End 标签 如 End CleanExit 等），包裹在单个 ```vba ... ``` 代码块中，以 End Sub 正常闭合。',
        '4. 【安全约束】：严禁调用 MsgBox、Application.Quit 或弹出阻塞式交互确认框。',
        '5. 在代码块前后提供简明扼要的自然语言说明。若需求存在合理理解差异，可自主选择合理方案并在说明中简要告知用户。'
    )
    return ($lines -join "`r`n")
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Set-GridValues($ws, $rows) {
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $row = $rows[$r]
        for ($c = 0; $c -lt $row.Count; $c++) {
            $ws.Cells.Item($r + 1, $c + 1).Value2 = $row[$c]
        }
    }
}

try {
    # Task 4: 筛选与汇总
    Write-Host "`n=== Running Task 4 (筛选与汇总) ===" -ForegroundColor Yellow
    $wb4Path = Join-Path $testDir "Task4_Target.xlsx"
    $wb4 = $excel.Workbooks.Add()
    $ws4 = $wb4.Sheets.Item(1)
    Set-GridValues $ws4 @(
        @("姓名", "部门", "职位", "薪资"),
        @("张伟", "研发部", "高级工程师", "22000"),
        @("王芳", "市场部", "商务专员", "12000"),
        @("李强", "研发部", "架构师", "35000"),
        @("刘洋", "财务部", "会计师", "15000"),
        @("陈静", "研发部", "测试主管", "18000"),
        @("赵敏", "市场部", "渠道经理", "20000")
    )
    $wb4.SaveAs($wb4Path)

    $sysPrompt4 = Build-GenericSysPrompt $wb4.Name @("Sheet1") "Sheet1" "A1:D7"
    $usrPrompt4 = "在当前表格开启自动筛选，并在表格下方统计各部门的总人数和平均工资"
    Write-Host "Requesting LLM for Task 4..."
    $resp4 = Call-RealLLM $sysPrompt4 $usrPrompt4
    $ext4 = Extract-VbaCode $resp4
    Write-Host "Extracted VBA length: $($ext4.Code.Length)"
    
    $req4 = @{
        action = "execute_vba"
        code = $ext4.Code
        prompt = $usrPrompt4
        targetWorkbookName = $wb4.Name
        targetWorkbookFullName = $wb4Path
        rawModelResponse = $resp4
    }
    $res4Json = [LeeExcel.NativeBridge]::Dispatch(($req4 | ConvertTo-Json -Compress), $excel)
    $parsed4 = [LeeExcel.SimpleJson]::ParseFlatObject($res4Json)
    Write-Host "Task 4 Result: OK = $($parsed4['ok']), Message = $($parsed4['message'])"
    if ($parsed4['ok'] -eq "True") {
        Write-Host "Task 4 UsedRange: $($ws4.UsedRange.Address($false, $false))"
    } else {
        Write-Host "Task 4 Error: $($parsed4['error'])"
    }
    $wb4.Close($false)

} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
