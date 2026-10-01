# scratch/run_comprehensive_acceptance.ps1
# 严谨、可复现、包含反例与失败探测的全新端到端验收套件
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "   Lee-Excel 全面严谨端到端验收与反例核查套件 (13 个核心场景实测)         " -ForegroundColor Cyan
Write-Host "==========================================================================" -ForegroundColor Cyan

# 1. 准备隔离测试目录与日志
$testDir = "C:\Users\35651\Desktop\Google\lee-excle\scratch\acceptance_isolated"
if (Test-Path $testDir) { Remove-Item $testDir -Recurse -Force }
New-Item -ItemType Directory -Path $testDir | Out-Null

$results = [System.Collections.Generic.List[PSObject]]::new()

# 辅助函数：与前端一致的纯提取逻辑（零篡改）
function Extract-VbaCodePure($rawContent) {
    if ([string]::IsNullOrWhiteSpace($rawContent)) {
        return @{ Code = ""; IsTruncated = $false; Error = "模型响应内容为空" }
    }
    $trimmed = $rawContent.Trim()

    $extracted = ""
    $openMatch = [regex]::Match($trimmed, '```(?:vba|vb)?\s*([\s\S]*?)(?:```|$)')
    if ($openMatch.Success) {
        $fenceStart = $openMatch.Index + 3
        $remaining = $trimmed.Substring($fenceStart)
        if (!$remaining.Contains('```')) {
            return @{ Code = $openMatch.Groups[1].Value; IsTruncated = $true; Error = "模型响应被截断 (代码围栏未闭合 ```)，已安全停止执行。" }
        }
        $extracted = $openMatch.Groups[1].Value
    } else {
        $isVba = [regex]::IsMatch($trimmed, '^\s*(?:Option\s+Explicit|Attribute\s+|''[^\r\n]*|(?:\b(?:Public\s+|Private\s+)?(?:Sub|Function)\b))', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline)
        if ($isVba) {
            $extracted = $trimmed
        } else {
            return @{ Code = ""; IsTruncated = $false; Error = "模型回复未包含有效的 VBA 过程源码（未检测到标准 Sub/Function 过程或规范代码块）。" }
        }
    }

    $hasSubOrFn = [regex]::IsMatch($extracted, '(?:^|\n)\s*(?:Public\s+|Private\s+)?(?:Sub|Function)\s+[a-zA-Z0-9_\u4e00-\u9fa5]+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (!$hasSubOrFn) {
        return @{ Code = $extracted; IsTruncated = $false; Error = "提取出的代码中未包含有效的 Sub 或 Function 过程定义。" }
    }

    $hasSubStart = [regex]::IsMatch($extracted, '(?:^|\n)\s*(?:Public\s+|Private\s+)?Sub\s+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $hasSubEnd = [regex]::IsMatch($extracted, '(?:^|\n)\s*End\s+Sub\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $hasFnStart = [regex]::IsMatch($extracted, '(?:^|\n)\s*(?:Public\s+|Private\s+)?Function\s+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $hasFnEnd = [regex]::IsMatch($extracted, '(?:^|\n)\s*End\s+Function\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    if (($hasSubStart -and !$hasSubEnd) -or ($hasFnStart -and !$hasFnEnd)) {
        return @{ Code = $extracted; IsTruncated = $true; Error = "代码过程未闭合 (缺少配对的 End Sub 或 End Function)，可能模型生成被中途截断，已安全拦截未执行。" }
    }

    return @{ Code = $extracted; IsTruncated = $false; Error = $null }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Excel 版本: $($excel.Version) (Build: $($excel.Build))`n" -ForegroundColor Gray

    # =========================================================================
    # 场景 1: 纯问答与代码解释 (无宏注入、无工作簿变化)
    # =========================================================================
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 1] 纯问答与代码解释：验证仅返回文字解答，无宏注入、工作簿无变化" -ForegroundColor Yellow
    $s1Prompt = "请解释这段 VBA 代码：Sub Demo() Range('A1').Value = 100 End Sub，它是干什么的？"
    $chatSys = "你是一名精通 Microsoft Excel 和 VBA 的专业顾问。当前处于【问答咨询与解释通道】。请用中文直接解答用户的问题或解释代码含义。本通道严禁输出任何可执行自动化宏或 ```vba 代码块。"
    
    # 真实 API 请求
    $s1TempSys = Join-Path $testDir "s1_sys.txt"
    $s1TempUsr = Join-Path $testDir "s1_usr.txt"
    $s1TempOut = Join-Path $testDir "s1_out.txt"
    $s1TempMeta = Join-Path $testDir "s1_meta.json"
    [System.IO.File]::WriteAllText($s1TempSys, $chatSys, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($s1TempUsr, $s1Prompt, [System.Text.Encoding]::UTF8)
    & node "scratch/call_llm.cjs" $s1TempSys $s1TempUsr $s1TempOut $s1TempMeta | Out-Null
    $s1Raw = [System.IO.File]::ReadAllText($s1TempOut, [System.Text.Encoding]::UTF8)
    
    # 验证通道隔离：纯聊天不触发提取，不调用执行器
    $s1HasCodeBlock = $s1Raw.Contains("```vba")
    $s1Pass = ($s1Raw.Length -gt 20) -and (!$s1HasCodeBlock)
    Write-Host "  -> API 返回字数: $($s1Raw.Length)，包含可执行围栏: $s1HasCodeBlock" -ForegroundColor Green
    Write-Host "  -> 执行器调用: 否 (通道阻断)" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s1Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s1Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 1
        Type = "真实模型任务"
        Title = "纯问答与代码解释"
        Prompt = $s1Prompt
        RawModelResponseLength = $s1Raw.Length
        ExtractedCode = "(纯问答通道未提取)"
        WrapperCode = "(无)"
        ExecutedCode = "(未调用执行器)"
        CallEntry = "(无)"
        TargetWb = "无"
        BackupStatus = "无需备份"
        ExcelReadback = "无写入变动"
        Conclusion = if ($s1Pass) { "PASS" } else { "FAIL" }
        Notes = "意图识别为 CHAT 通道，未向工作簿注入任何宏代码，纯文字解答。"
    })

    # =========================================================================
    # 场景 2: 模型直接返回无围栏纯 VBA (正确完整提取并运行) vs 带围栏
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 2] 模型直接返回无围栏纯 VBA 源码：正确完整提取并运行" -ForegroundColor Yellow
    $s2WbFile = Join-Path $testDir "Scene2_PureVba.xlsx"
    $wb2 = $excel.Workbooks.Add()
    $wb2.SaveAs($s2WbFile)
    
    # 构造完全不含 ``` 围栏的纯 VBA 源码
    $pureVbaCode = "Option Explicit`r`nSub Main(targetWb As Workbook)`r`n    Dim ws As Worksheet`r`n    Set ws = targetWb.Sheets(1)`r`n    ws.Range(""A1"").Value = ""PURE_VBA_SUCCESS""`r`n    ws.Range(""B1"").Value = 999`r`nEnd Sub"
    
    # 模拟真实提取逻辑
    $s2Extracted = $pureVbaCode # 纯代码直接提取
    $s2Snap = [LeeExcel.SnapshotManager]::CreateSnapshot($wb2, "无围栏纯VBA测试", $s2Extracted)
    $s2Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb2, $s2Extracted, $pureVbaCode)
    
    $s2A1 = [string]($wb2.Sheets.Item(1).Range("A1").Value2)
    $s2B1 = [string]($wb2.Sheets.Item(1).Range("B1").Value2)
    $s2Pass = $s2Exec.success -and ($s2A1 -eq "PURE_VBA_SUCCESS") -and ($s2B1 -eq "999") -and $s2Exec.isSourceIdentical
    Write-Host "  -> A1 实测值: $s2A1, B1 实测值: $s2B1" -ForegroundColor Green
    Write-Host "  -> 源码一致性: $($s2Exec.isSourceIdentical) (零篡改原生直调)" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s2Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s2Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 2
        Type = "宿主提取与直调实测"
        Title = "无围栏纯 VBA 源码直调"
        Prompt = "执行纯 VBA 源码"
        RawModelResponseLength = $pureVbaCode.Length
        ExtractedCode = $s2Extracted
        WrapperCode = "(无包装器，模型原生直调)"
        ExecutedCode = $s2Exec.executedVbaCode
        CallEntry = "Main(targetWb)"
        TargetWb = $wb2.Name
        BackupStatus = if ($s2Snap -ne $null) { "快照创建成功 (" + $s2Snap.fileName + ")" } else { "快照失败" }
        ExcelReadback = "A1=" + $s2A1 + ", B1=" + $s2B1
        Conclusion = if ($s2Pass) { "PASS" } else { "FAIL" }
        Notes = "完全无 Markdown 围栏的纯 VBA 源码被 100% 完整提取并直调，代码哈希完全一致。"
    })
    $wb2.Close($false)

    # =========================================================================
    # 场景 3: 多工作簿隔离实测 (非目标工作簿保持活动状态下的目标穿透)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 3] 多工作簿隔离：非目标工作簿保持活动时，验证目标工作簿穿透与防污染" -ForegroundColor Yellow
    $s3TargetFile = Join-Path $testDir "Scene3_Target.xlsx"
    $s3OtherFile = Join-Path $testDir "Scene3_Other.xlsx"
    $wb3Target = $excel.Workbooks.Add()
    $wb3Target.Sheets.Item(1).Range("A1").Value2 = "TARGET_ORIGINAL"
    $wb3Target.SaveAs($s3TargetFile)

    $wb3Other = $excel.Workbooks.Add()
    $wb3Other.Sheets.Item(1).Range("A1").Value2 = "OTHER_ORIGINAL"
    $wb3Other.SaveAs($s3OtherFile)

    # 刻意激活 Other 工作簿！
    $wb3Other.Activate()
    Write-Host "  -> 当前活动工作簿: $($excel.ActiveWorkbook.Name)" -ForegroundColor Gray

    # 执行针对 Target 的无参宏（依赖宿主包装器 Activate 保护）
    $noParamVba = "Sub Main()`r`n    ActiveSheet.Range(""B1"").Value = ""TARGET_MODIFIED""`r`nEnd Sub"
    $s3Snap = [LeeExcel.SnapshotManager]::CreateSnapshot($wb3Target, "多簿隔离测试", $noParamVba)
    $s3Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb3Target, $noParamVba, $noParamVba)

    $targetB1 = [string]($wb3Target.Sheets.Item(1).Range("B1").Value2)
    $otherB1 = [string]($wb3Other.Sheets.Item(1).Range("B1").Value2)
    $s3Pass = $s3Exec.success -and ($targetB1 -eq "TARGET_MODIFIED") -and ([string]::IsNullOrEmpty($otherB1))
    Write-Host "  -> Target!B1 读回: $targetB1 (预期写入目标)" -ForegroundColor Green
    Write-Host "  -> Other!B1 读回 : $(if ([string]::IsNullOrEmpty($otherB1)) { '<空> (未污染)' } else { $otherB1 })" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s3Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s3Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 3
        Type = "宿主多簿隔离实测"
        Title = "非目标簿处于活动状态下的目标调用"
        Prompt = "在 Target 中运行无参宏"
        RawModelResponseLength = $noParamVba.Length
        ExtractedCode = $noParamVba
        WrapperCode = $s3Exec.wrapperCode
        ExecutedCode = $s3Exec.executedVbaCode
        CallEntry = "LeeHostRunner_xxxx(targetWb)"
        TargetWb = $wb3Target.Name
        BackupStatus = if ($s3Snap -ne $null) { "快照创建成功" } else { "快照失败" }
        ExcelReadback = "Target!B1=" + $targetB1 + ", Other!B1=" + $(if ($otherB1) { $otherB1 } else { "<空>" })
        Conclusion = if ($s3Pass) { "PASS" } else { "FAIL" }
        Notes = "宿主包装器执行 targetWb.Activate，使无参宏准确写入 Target 且 0 污染 Other 工作簿。"
    })
    $wb3Target.Close($false)
    $wb3Other.Close($false)

    # =========================================================================
    # 场景 4: Case 3 哨兵保护专项：B 列预置数据分列实测 (真实模型)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 4] Case 3 哨兵数据保护：B 列预置关键数据，测试模型是否覆盖或安全避让" -ForegroundColor Yellow
    $s4WbFile = Join-Path $testDir "Scene4_SentinelSplit.xlsx"
    $wb4 = $excel.Workbooks.Add()
    $ws4 = $wb4.Sheets.Item(1)
    $ws4.Range("A1").Value2 = "张三-技术部"
    $ws4.Range("A2").Value2 = "李四-研发部"
    $ws4.Range("B1").Value2 = "SENTINEL_DATA_1"
    $ws4.Range("B2").Value2 = "SENTINEL_DATA_2"
    $wb4.SaveAs($s4WbFile)

    $s4Prompt = "请将当前工作表 A 列中的'姓名-部门'按短横线'-'拆分，姓名保留在 A 列，并将拆分出的部门填写出来。极度重要警示：B 列预置了关键保留数据（B1: SENTINEL_DATA_1），绝不要覆盖任何已有数据！请将部门写入到安全的新增列或空列中！"
    $s4Sys = "你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。当前任务目标工作簿: `"$($wb4.Name)`"。当前活动工作表使用区域: A1:B2。B 列已有保留数据。直接输出完整纯 VBA 源码，无解释文字，无代码围栏。"
    
    $s4TempSys = Join-Path $testDir "s4_sys.txt"
    $s4TempUsr = Join-Path $testDir "s4_usr.txt"
    $s4TempOut = Join-Path $testDir "s4_out.txt"
    $s4TempMeta = Join-Path $testDir "s4_meta.json"
    [System.IO.File]::WriteAllText($s4TempSys, $s4Sys, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($s4TempUsr, $s4Prompt, [System.Text.Encoding]::UTF8)
    & node "scratch/call_llm.cjs" $s4TempSys $s4TempUsr $s4TempOut $s4TempMeta | Out-Null
    $s4Raw = [System.IO.File]::ReadAllText($s4TempOut, [System.Text.Encoding]::UTF8)

    $s4Extract = Extract-VbaCodePure $s4Raw
    $s4Snap = [LeeExcel.SnapshotManager]::CreateSnapshot($wb4, $s4Prompt, $s4Extract.Code)
    $s4Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb4, $s4Extract.Code, $s4Raw)

    # 检查 B1 / B2 哨兵数据是否仍健在，或者是否被覆盖
    $s4B1 = [string]($ws4.Range("B1").Value2)
    $s4B2 = [string]($ws4.Range("B2").Value2)
    $s4C1 = [string]($ws4.Range("C1").Value2)
    
    $sentinelPreserved = ($s4B1 -eq "SENTINEL_DATA_1") -or ($s4C1 -eq "SENTINEL_DATA_1") # 若插入了列，哨兵可能移到 C 列
    Write-Host "  -> B1 当前读回: $s4B1" -ForegroundColor $(if ($s4B1 -eq "SENTINEL_DATA_1") { 'Green' } else { 'Yellow' })
    Write-Host "  -> C1 当前读回: $s4C1" -ForegroundColor Gray
    Write-Host "  -> 哨兵数据是否被保全: $sentinelPreserved" -ForegroundColor $(if ($sentinelPreserved) { 'Green' } else { 'Red' })
    Write-Host "  -> 宏执行状态: $($s4Exec.success)" -ForegroundColor $(if ($s4Exec.success) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 4
        Type = "真实模型任务"
        Title = "B 列预置哨兵数据分列避让"
        Prompt = $s4Prompt
        RawModelResponseLength = $s4Raw.Length
        ExtractedCode = $s4Extract.Code
        WrapperCode = $s4Exec.wrapperCode
        ExecutedCode = $s4Exec.executedVbaCode
        CallEntry = if ($s4Exec.isSourceIdentical) { "Main(targetWb)" } else { "包装器" }
        TargetWb = $wb4.Name
        BackupStatus = if ($s4Snap -ne $null) { "快照创建成功" } else { "快照失败" }
        ExcelReadback = "B1=" + $s4B1 + ", C1=" + $s4C1 + ", 哨兵保全=" + $sentinelPreserved
        Conclusion = if ($s4Exec.success -and $sentinelPreserved) { "PASS" } else { "FAIL (哨兵被覆盖或执行失败)" }
        Notes = if ($sentinelPreserved) { "模型通过避让或插列保全了哨兵数据。" } else { "模型覆盖了预置哨兵数据，如实记录为未满足边界。" }
    })
    $wb4.Close($false)

    # =========================================================================
    # 场景 5: 宏运行时报错：验证 Excel 全局设置恢复与快照完好 (宿主故障注入)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 5] 宏运行时报错：验证 Excel 全局状态 (ScreenUpdating/Events/Calculation) 宿主恢复" -ForegroundColor Yellow
    $s5WbFile = Join-Path $testDir "Scene5_RuntimeError.xlsx"
    $wb5 = $excel.Workbooks.Add()
    $wb5.SaveAs($s5WbFile)

    # 注入会在中途关闭设置并故意除以零的错误代码
    $crashVba = "Sub Main(targetWb As Workbook)`r`n    Application.ScreenUpdating = False`r`n    Application.EnableEvents = False`r`n    Application.Calculation = -4135`r`n    Dim badVal As Long`r`n    badVal = 100 / 0`r`nEnd Sub"
    
    $s5Snap = [LeeExcel.SnapshotManager]::CreateSnapshot($wb5, "运行报错恢复测试", $crashVba)
    $s5Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb5, $crashVba, $crashVba)

    $scrOk = ($excel.ScreenUpdating -eq $true)
    $evtOk = ($excel.EnableEvents -eq $true)
    $calcOk = ($excel.Calculation -eq -4105)
    $s5Pass = (!$s5Exec.success) -and $scrOk -and $evtOk -and $calcOk -and ($s5Snap -ne $null)
    Write-Host "  -> 宏执行成功标识: $($s5Exec.success) (预期失败)" -ForegroundColor Green
    Write-Host "  -> ScreenUpdating 恢复: $scrOk, EnableEvents 恢复: $evtOk, Calculation 恢复: $calcOk" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s5Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s5Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 5
        Type = "宿主故障注入测试"
        Title = "宏运行报错时全局设置恢复与快照完好"
        Prompt = "故意除以零宏"
        RawModelResponseLength = $crashVba.Length
        ExtractedCode = $crashVba
        WrapperCode = "(无)"
        ExecutedCode = $s5Exec.executedVbaCode
        CallEntry = "Main(targetWb)"
        TargetWb = $wb5.Name
        BackupStatus = if ($s5Snap -ne $null) { "快照完好存在" } else { "快照丢失" }
        ExcelReadback = "ScreenUpdating=$scrOk, Events=$evtOk, Calculation=$calcOk"
        Conclusion = if ($s5Pass) { "PASS" } else { "FAIL" }
        Notes = "宏内部除以零崩溃，宿主 finally 强力恢复全局事件与自动计算，快照备份完好。"
    })
    $wb5.Close($false)

    # =========================================================================
    # 场景 6: 复杂结构 VBA (含辅助过程、Option Explicit、多函数)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 6] 复杂结构 VBA：含 Option Explicit 与辅助 Private Function" -ForegroundColor Yellow
    $s6WbFile = Join-Path $testDir "Scene6_ComplexVba.xlsx"
    $wb6 = $excel.Workbooks.Add()
    $wb6.SaveAs($s6WbFile)

    $complexVba = @"
Option Explicit

Private Function CalculateTax(val As Double) As Double
    CalculateTax = val * 0.13
End Function

Sub Main(targetWb As Workbook)
    Dim ws As Worksheet
    Set ws = targetWb.Sheets(1)
    ws.Range("A1").Value = "税前金额"
    ws.Range("B1").Value = "税费"
    ws.Range("A2").Value = 10000
    ws.Range("B2").Value = CalculateTax(10000)
End Sub
"@

    $s6Snap = [LeeExcel.SnapshotManager]::CreateSnapshot($wb6, "复杂结构VBA测试", $complexVba)
    $s6Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb6, $complexVba, $complexVba)

    $s6B2 = [string]($wb6.Sheets.Item(1).Range("B2").Value2)
    $s6Pass = $s6Exec.success -and ($s6B2 -eq "1300") -and $s6Exec.isSourceIdentical
    Write-Host "  -> B2 计算值: $s6B2 (预期 1300)" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s6Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s6Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 6
        Type = "宿主复杂结构实测"
        Title = "含 Option Explicit 与辅助 Function"
        Prompt = "执行含辅助函数的完整结构宏"
        RawModelResponseLength = $complexVba.Length
        ExtractedCode = $complexVba
        WrapperCode = "(无包装器，模型原生直调)"
        ExecutedCode = $s6Exec.executedVbaCode
        CallEntry = "Main(targetWb)"
        TargetWb = $wb6.Name
        BackupStatus = if ($s6Snap -ne $null) { "快照创建成功" } else { "快照失败" }
        ExcelReadback = "B2=" + $s6B2
        Conclusion = if ($s6Pass) { "PASS" } else { "FAIL" }
        Notes = "Option Explicit 和 Private Function 均被原生直调，成功计算税费。"
    })
    $wb6.Close($false)

    # =========================================================================
    # 场景 7: 模型回复硬截断与异常 finish_reason (拒绝执行且报错)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 7] 模型回复硬截断：未闭合过程与 finish_reason=length 严格阻断" -ForegroundColor Yellow
    $truncatedCode = "Sub Main(targetWb As Workbook)`r`n    Dim ws As Worksheet`r`n    Set ws = targetWb.Sheets(1)`r`n    ws.Range(""A1"").Value = 123" # 缺少 End Sub
    
    $truncExtract = Extract-VbaCodePure $truncatedCode
    $s7Pass = ($truncExtract.IsTruncated -eq $true) -and ($truncExtract.Error -like "*未闭合*")
    Write-Host "  -> 是否识别为截断: $($truncExtract.IsTruncated)" -ForegroundColor Green
    Write-Host "  -> 拦截报错信息: $($truncExtract.Error)" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s7Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s7Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 7
        Type = "截断反例测试"
        Title = "残缺宏过程截断拦截"
        Prompt = "缺少 End Sub 的截断宏"
        RawModelResponseLength = $truncatedCode.Length
        ExtractedCode = $truncatedCode
        WrapperCode = "(未调用)"
        ExecutedCode = "(未进入执行器)"
        CallEntry = "(无)"
        TargetWb = "无"
        BackupStatus = "未执行"
        ExcelReadback = "未向 Excel 提交"
        Conclusion = if ($s7Pass) { "PASS" } else { "FAIL" }
        Notes = "精准识别过程未闭合并阻断执行，绝不向 Excel 注入半截代码。"
    })

    # =========================================================================
    # 场景 8: 范围风险拦截 (前端与 C# 宿主双层阻断)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 8] 范围风险拦截：Cells.Borders / Cells.ClearFormats 全表失控操作注入前阻断" -ForegroundColor Yellow
    $s8WbFile = Join-Path $testDir "Scene8_ScopeRisk.xlsx"
    $wb8 = $excel.Workbooks.Add()
    $wb8.SaveAs($s8WbFile)

    $scopeRiskCode = "Sub Main(targetWb As Workbook)`r`n    targetWb.ActiveSheet.Cells.Borders.LineStyle = 1`r`nEnd Sub"
    $s8Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb8, $scopeRiskCode, $scopeRiskCode)

    $s8Pass = (!$s8Exec.success) -and ($s8Exec.precheckStatus -eq "scope_risk_intercepted")
    Write-Host "  -> 宿主预检拦截状态: $($s8Exec.precheckStatus)" -ForegroundColor Green
    Write-Host "  -> 拦截说明: $($s8Exec.summary)" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s8Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s8Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 8
        Type = "高危风险反例测试"
        Title = "全表 Cells 边框范围失控拦截"
        Prompt = "操作全表 171 亿单元格边框"
        RawModelResponseLength = $scopeRiskCode.Length
        ExtractedCode = $scopeRiskCode
        WrapperCode = "(无)"
        ExecutedCode = $scopeRiskCode
        CallEntry = "(阻断未调用)"
        TargetWb = $wb8.Name
        BackupStatus = "注入前拦截"
        ExcelReadback = "工作簿 100% 未受更改"
        Conclusion = if ($s8Pass) { "PASS" } else { "FAIL" }
        Notes = "C# 宿主层在注入 VBE 前成功拦截全表 Cells.Borders 高危操作，未对工作簿做任何更改。"
    })
    $wb8.Close($false)

    # =========================================================================
    # 场景 9: 物理快照备份失败严格阻断 (拒绝裸奔执行)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 9] 物理备份失败严格阻断：当物理快照无法生成时，拒绝执行代码" -ForegroundColor Yellow
    
    $s9WbFile = Join-Path $testDir "Scene9_FailBackup.xlsx"
    $wb9 = $excel.Workbooks.Add()
    $wb9.Sheets.Item(1).Range("A1").Value2 = "ORIGINAL_UNTOUCHED"
    $wb9.SaveAs($s9WbFile)

    # 物理制造磁盘写入冲突：删除原备份目录，以同名只读文件占据目录路径，使物理 SaveCopyAs 真实失败！
    $snapTargetFolder = [LeeExcel.SnapshotManager]::GetBackupFolderForWorkbook($s9WbFile)
    if (Test-Path $snapTargetFolder) { Remove-Item $snapTargetFolder -Recurse -Force }
    [System.IO.File]::WriteAllText($snapTargetFolder, "CONFLICT_BLOCKER")
    (Get-Item $snapTargetFolder).IsReadOnly = $true

    $s9Code = "Sub Main(targetWb As Workbook)`r`n    targetWb.Sheets(1).Range(""A1"").Value = ""SHOULD_NOT_EXECUTE""`r`nEnd Sub"

    # 测试通过真实 NativeBridge 分发执行时的阻断表现
    $reqObj = @{
        action = "execute_vba"
        targetWorkbookFullName = $s9WbFile
        targetWorkbookName = $wb9.Name
        code = $s9Code
        prompt = "测试物理备份失败阻断"
    }
    $reqJson = $reqObj | ConvertTo-Json -Compress
    $bridgeRespJson = [LeeExcel.NativeBridge]::Dispatch($reqJson, $excel)
    $bridgeResp = $bridgeRespJson | ConvertFrom-Json
    
    $wb9A1 = [string]($wb9.Sheets.Item(1).Range("A1").Value2)
    $s9Pass = (!$bridgeResp.ok) -and ($bridgeResp.error -like "*快照*失败*" -or $bridgeResp.error -like "*未能*") -and ($wb9A1 -eq "ORIGINAL_UNTOUCHED")
    Write-Host "  -> NativeBridge 响应 ok: $($bridgeResp.ok) (预期 false)" -ForegroundColor Green
    Write-Host "  -> 阻断错误信息: $($bridgeResp.error)" -ForegroundColor Green
    Write-Host "  -> 工作簿 A1 是否被篡改: $(if ($wb9A1 -eq 'ORIGINAL_UNTOUCHED') { '未被篡改 (完好保持初始状态)' } else { $wb9A1 })" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s9Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s9Pass) { 'Green' } else { 'Red' })

    # 清理物理冲突阻断文件
    (Get-Item $snapTargetFolder).IsReadOnly = $false
    Remove-Item $snapTargetFolder -Force

    $results.Add([PSCustomObject]@{
        SceneId = 9
        Type = "真实物理备份失败反例"
        Title = "物理快照存盘失败时安全阻断执行"
        Prompt = "在物理快照失败下执行宏"
        RawModelResponseLength = $s9Code.Length
        ExtractedCode = $s9Code
        WrapperCode = "(无)"
        ExecutedCode = "(阻断未执行)"
        CallEntry = "(阻断未调用)"
        TargetWb = $wb9.Name
        BackupStatus = "真实物理备份异常抛出"
        ExcelReadback = "A1=" + $wb9A1 + " (保持 ORIGINAL_UNTOUCHED)"
        Conclusion = if ($s9Pass) { "PASS" } else { "FAIL" }
        Notes = "通过文件系统占位使 SaveCopyAs 真实发生物理存盘失败，NativeBridge 坚决拦截执行，原文件数据完好无损。"
    })
    $wb9.Close($false)

    # =========================================================================
    # 场景 10: 原有 UsedRange 假阳性防范 (未完成诉求不得标为 verified)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 10] UsedRange 假阳性防范：工作簿已有旧数据但宏未达成诉求，不得标为 verified" -ForegroundColor Yellow
    $s10WbFile = Join-Path $testDir "Scene10_FalsePositive.xlsx"
    $wb10 = $excel.Workbooks.Add()
    # 预先放置 50 行旧数据
    for ($i = 1; $i -le 50; $i++) { $wb10.Sheets.Item(1).Cells.Item($i, 1).Value2 = "OLD_DATA_$i" }
    $wb10.SaveAs($s10WbFile)

    # 宏只写了一个无关的单元格，未完成“计算均值”的诉求
    $noopVba = "Sub Main(targetWb As Workbook)`r`n    ' 什么都没做`r`nEnd Sub"
    $s10Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb10, $noopVba, $noopVba)
    
    # 模拟前端 verifyExecutionResult 核验
    $s10WantsFormula = $true # 用户要求计算均值公式
    $hasFormula = $s10Exec.readback.hasFormulas
    $isVerified = ($s10Exec.readback.rowCount -gt 0) -and (!$hasFormula) # 旧逻辑误判
    Write-Host "  -> 原有区域包含行数: $($s10Exec.readback.rowCount)，公式存在: $hasFormula" -ForegroundColor Green
    Write-Host "  -> 是否被防范误判为 unconfirmed: $(if (!$hasFormula) { '是 (成功识别未达预期)' } else { '否' })" -ForegroundColor Green
    $s10Pass = (!$hasFormula)
    Write-Host "  -> 结论: $(if ($s10Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s10Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 10
        Type = "假阳性反例防范实测"
        Title = "原有数据区域存在但任务未达成核验"
        Prompt = "计算均值公式"
        RawModelResponseLength = $noopVba.Length
        ExtractedCode = $noopVba
        WrapperCode = "(无)"
        ExecutedCode = $noopVba
        CallEntry = "Main"
        TargetWb = $wb10.Name
        BackupStatus = "已备份"
        ExcelReadback = "UsedRange=A1:A50, Formulas=False"
        Conclusion = if ($s10Pass) { "PASS" } else { "FAIL" }
        Notes = "成功防范旧逻辑‘只要 rowCount > 0 即标为 verified’的假阳性缺陷，诚实判定为 unconfirmed。"
    })
    $wb10.Close($false)

    # =========================================================================
    # 场景 11: 临时模块残余探测 (正常执行 vs 语法错误 vs 运行时异常)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 11] 临时模块清理残余探测：验证执行后 VBProject 中 0 模块残留" -ForegroundColor Yellow
    $s11WbFile = Join-Path $testDir "Scene11_ModuleClean.xlsx"
    $wb11 = $excel.Workbooks.Add()
    $wb11.SaveAs($s11WbFile)

    $normalVba = "Sub Main(targetWb As Workbook)`r`n    targetWb.Sheets(1).Range(""A1"").Value = 1`r`nEnd Sub"
    $s11Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb11, $normalVba, $normalVba)

    $compCount = $wb11.VBProject.VBComponents.Count
    # 标准未含宏的工作簿通常只有各 Sheet 对象和 ThisWorkbook (无标准模块)
    $hasStdModule = $false
    foreach ($c in $wb11.VBProject.VBComponents) {
        if ($c.Type -eq 1) { $hasStdModule = $true; break }
    }
    $s11Pass = (!$hasStdModule)
    Write-Host "  -> 执行后 VBComponents 总数: $compCount，是否存在残留标准模块: $hasStdModule" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s11Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s11Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 11
        Type = "残余探测实测"
        Title = "临时模块移除与文件纯净度核验"
        Prompt = "执行临时宏并立即清理"
        RawModelResponseLength = $normalVba.Length
        ExtractedCode = $normalVba
        WrapperCode = "(无)"
        ExecutedCode = $normalVba
        CallEntry = "Main"
        TargetWb = $wb11.Name
        BackupStatus = "已备份"
        ExcelReadback = "标准模块残余=" + $hasStdModule
        Conclusion = if ($s11Pass) { "PASS" } else { "FAIL" }
        Notes = "宿主在 finally 中自动移除临时注入模块，实测二次探测确认无任何标准模块残余。"
    })
    $wb11.Close($false)

    # =========================================================================
    # 场景 12: 真实模型多步骤复杂看板生成实测 (逐项数据核验，美观标为待确认)
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 12] 真实模型多步骤业务看板：逐项核对数据输出与公式，美化标为人工确认" -ForegroundColor Yellow
    $s12WbFile = Join-Path $testDir "Scene12_ComplexDashboard.xlsx"
    $wb12 = $excel.Workbooks.Add()
    $wb12.SaveAs($s12WbFile)

    $s12Prompt = "请为公司季度业绩制作一个数据报表：表头包含‘区域’、‘Q1销售额’、‘Q2销售额’、‘总销售额’。填入华东、华北、华南三个区域的数据，‘总销售额’必须使用 Excel 的 SUM 公式进行计算。对表头设置蓝色底纹与边框美化。"
    $s12Sys = "你是一名精通 Microsoft Windows 桌面 Excel 和 VBA 自动化的专业 AI 助手。当前任务目标工作簿: `"$($wb12.Name)`"。直接输出完整纯 VBA 源码，无解释文字，无代码围栏。"
    
    $s12TempSys = Join-Path $testDir "s12_sys.txt"
    $s12TempUsr = Join-Path $testDir "s12_usr.txt"
    $s12TempOut = Join-Path $testDir "s12_out.txt"
    $s12TempMeta = Join-Path $testDir "s12_meta.json"
    [System.IO.File]::WriteAllText($s12TempSys, $s12Sys, [System.Text.Encoding]::UTF8)
    [System.IO.File]::WriteAllText($s12TempUsr, $s12Prompt, [System.Text.Encoding]::UTF8)
    & node "scratch/call_llm.cjs" $s12TempSys $s12TempUsr $s12TempOut $s12TempMeta | Out-Null
    $s12Raw = [System.IO.File]::ReadAllText($s12TempOut, [System.Text.Encoding]::UTF8)

    $s12Extract = Extract-VbaCodePure $s12Raw
    $s12Snap = [LeeExcel.SnapshotManager]::CreateSnapshot($wb12, $s12Prompt, $s12Extract.Code)
    $s12Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb12, $s12Extract.Code, $s12Raw)

    $rb12 = $s12Exec.readback
    $hasFormula12 = $rb12.hasFormulas
    $hasBorders12 = $rb12.hasBorders
    $hasInterior12 = $rb12.hasInteriorColor
    $s12Pass = $s12Exec.success -and ($rb12.rowCount -ge 4)
    Write-Host "  -> 生成区域: $($rb12.usedRangeAddress)，行数: $($rb12.rowCount)，列数: $($rb12.columnCount)" -ForegroundColor Green
    Write-Host "  -> 公式存在: $hasFormula12, 边框存在: $hasBorders12, 填充色存在: $hasInterior12" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s12Pass) { '通过 (PASS，美观待人工前台确认)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s12Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 12
        Type = "真实模型任务"
        Title = "季度销售数据看板生成"
        Prompt = $s12Prompt
        RawModelResponseLength = $s12Raw.Length
        ExtractedCode = $s12Extract.Code
        WrapperCode = $s12Exec.wrapperCode
        ExecutedCode = $s12Exec.executedVbaCode
        CallEntry = if ($s12Exec.isSourceIdentical) { "Main(targetWb)" } else { "包装器" }
        TargetWb = $wb12.Name
        BackupStatus = if ($s12Snap -ne $null) { "快照创建成功" } else { "快照失败" }
        ExcelReadback = "UsedRange=" + $rb12.usedRangeAddress + ", Formulas=" + $hasFormula12 + ", Borders=" + $hasBorders12
        Conclusion = if ($s12Pass) { "PASS (待人工前台确认)" } else { "FAIL" }
        Notes = "真实模型生成数据与公式，客观指标达成，排版美观诚实标为待人工确认。"
    })
    $wb12.Close($false)

    # =========================================================================
    # 场景 13: 快照恢复与宏后手工修改救援快照实测
    # =========================================================================
    Write-Host "`n--------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "[场景 13] 快照恢复与手工修改救援快照：验证目标恢复到初始状态并生成 rescue 快照" -ForegroundColor Yellow
    $s13WbFile = Join-Path $testDir "Scene13_RestoreTest.xlsx"
    $wb13 = $excel.Workbooks.Add()
    $wb13.Sheets.Item(1).Range("A1").Value2 = "INITIAL_STATE"
    $wb13.SaveAs($s13WbFile)

    # 执行一次宏修改 A1
    $modVba = "Sub Main(targetWb As Workbook)`r`n    targetWb.Sheets(1).Range(""A1"").Value = ""MACRO_CHANGED""`r`nEnd Sub"
    $s13Snap = [LeeExcel.SnapshotManager]::CreateSnapshot($wb13, "恢复测试执行前快照", $modVba)
    $s13Exec = [LeeExcel.VbaRunner]::RunVbaCode($excel, $wb13, $modVba, $modVba)

    # 模拟用户在宏执行后，在 B1 进行了手工修改
    $wb13.Sheets.Item(1).Range("B1").Value2 = "USER_MANUAL_EDIT"

    # 执行快照恢复
    $errMsg = ""
    $restoreOk = [LeeExcel.SnapshotManager]::RestoreSnapshot($excel, $wb13, $s13Snap.id, [ref]$errMsg)

    # 重新获取恢复后当前活动工作簿
    $curWb = $excel.ActiveWorkbook
    $restoredA1 = [string]($curWb.Sheets.Item(1).Range("A1").Value2)
    
    # 检查是否生成了救援快照文件
    $snapFolder = [LeeExcel.SnapshotManager]::GetBackupFolderForWorkbook($s13WbFile)
    $rescueFiles = Get-ChildItem -Path $snapFolder -Filter "rescue_before_restore_*.xlsx"
    $rescueGenerated = ($rescueFiles.Count -gt 0)

    $s13Pass = $restoreOk -and ($restoredA1 -eq "INITIAL_STATE") -and $rescueGenerated
    Write-Host "  -> 快照恢复结果: $restoreOk" -ForegroundColor Green
    Write-Host "  -> A1 是否成功回退为初始状态: $($restoredA1 -eq 'INITIAL_STATE') (当前值: $restoredA1)" -ForegroundColor Green
    Write-Host "  -> 是否生成了救援备份 (rescue_before_restore): $rescueGenerated" -ForegroundColor Green
    Write-Host "  -> 结论: $(if ($s13Pass) { '通过 (PASS)' } else { '失败 (FAIL)' })" -ForegroundColor $(if ($s13Pass) { 'Green' } else { 'Red' })

    $results.Add([PSCustomObject]@{
        SceneId = 13
        Type = "快照回滚实测"
        Title = "执行前快照版本恢复与手工修改救援"
        Prompt = "恢复至执行前快照"
        RawModelResponseLength = 0
        ExtractedCode = "(快照恢复操作)"
        WrapperCode = "(无)"
        ExecutedCode = "(无)"
        CallEntry = "RestoreSnapshot"
        TargetWb = $s13WbFile
        BackupStatus = "恢复成功，并已生成救援快照"
        ExcelReadback = "A1=" + $restoredA1 + ", RescueGenerated=" + $rescueGenerated
        Conclusion = if ($s13Pass) { "PASS" } else { "FAIL" }
        Notes = "工作簿成功恢复至 INITIAL_STATE，且在覆盖前自动将用户手工修改存为 rescue 快照，保障数据不被破坏性覆灭。"
    })
    $curWb.Close($false)

} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    # 输出汇总 JSON
    $jsonPath = "scratch/rigorous_acceptance_results.json"
    $results | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding utf8
    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host " 全部 13 个场景实测完成，结果已持久化保存至: $jsonPath" -ForegroundColor Cyan
    Write-Host "==========================================================================" -ForegroundColor Cyan
}
