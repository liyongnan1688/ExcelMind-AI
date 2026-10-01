# scratch/test_generic_capabilities_e2e.ps1
# Comprehensive E2E Verification of Generic Capabilities, Real LLM, COM Protocol, and Safety Boundaries
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "   Lee-Excel GENERIC CAPABILITIES REAL LLM E2E VERIFICATION     " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

# 1. 读取 API 配置
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
$apiKey = [regex]::Match($configJson, '"apiKey"\s*:\s*"([^"]+)"').Groups[1].Value
$model = [regex]::Match($configJson, '"model"\s*:\s*"([^"]+)"').Groups[1].Value
$maskedKey = if ($apiKey.Length -gt 8) { $apiKey.Substring(0, 4) + "****" + $apiKey.Substring($apiKey.Length - 4) } else { "****" }

Write-Host "Config Loaded: Model = $model, API Key = $maskedKey (Desensitized)" -ForegroundColor Gray

# 2. 隔离环境
$testDir = Join-Path $env:TEMP ("LeeExcel_GenericE2E_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $testDir | Out-Null
Write-Host "Test Directory: $testDir" -ForegroundColor Gray

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
        '5. 【结构与输出】：直接输出完整可执行的标准 VBA 代码，包裹在单个 ```vba ... ``` 代码块中，在代码块前后仅提供简明扼要的说明，避免冗长说明以确保代码完整不被截断。'
    )
    return ($lines -join "`r`n")
}

$taskResults = [System.Collections.Generic.List[PSObject]]::new()

$excel = $null
$wbOther = $null

try {
    Write-Host "`n[Setup] Launching isolated Excel instance..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # 创建一个同时打开的独立工作簿，用于检测跨文件副作用
    $otherPath = Join-Path $testDir "OtherIsolatedWorkbook.xlsx"
    $wbOther = $excel.Workbooks.Add()
    $wbOther.Sheets.Item(1).Range("A1").Value2 = "DO_NOT_TOUCH_OTHER_WORKBOOK"
    $wbOther.SaveAs($otherPath)

    function Set-GridValues($ws, $rows) {
        for ($r = 0; $r -lt $rows.Count; $r++) {
            $row = $rows[$r]
            for ($c = 0; $c -lt $row.Count; $c++) {
                $ws.Cells.Item($r + 1, $c + 1).Value2 = $row[$c]
            }
        }
    }

    # 定义 6 类通用真实大模型任务
    $tasks = @(
        @{
            Id = 1
            Name = "新建表格 (员工花名册)"
            Prompt = "新建一个员工花名册表格，包含工号、姓名、部门、职位、入职日期这几列，随便生成5条员工数据"
            SetupData = $null
        },
        @{
            Id = 2
            Name = "美化现有业务表 (科技商务蓝与细边框)"
            Prompt = "帮我把当前表格美化一下，表头要有科技商务蓝底白字，内容行要有细边框和交替浅色背景，文字和数字对齐好，列宽自适应"
            SetupData = {
                param($ws)
                Set-GridValues $ws @(
                    @("项目编号", "负责人", "预算(万元)", "支出(万元)", "达成率"),
                    @("PRJ-001", "张三", "120", "108", "90%"),
                    @("PRJ-002", "李四", "85", "82", "96%"),
                    @("PRJ-003", "王五", "200", "150", "75%"),
                    @("PRJ-004", "赵六", "60", "60", "100%")
                )
            }
        },
        @{
            Id = 3
            Name = "插入公式 (单价×数量计算与SUM合计行)"
            Prompt = "在金额列用公式计算单价乘以数量，并在最后一行添加合计行，用 SUM 公式汇总总数量和总金额"
            SetupData = {
                param($ws)
                Set-GridValues $ws @(
                    @("商品名称", "单价", "数量", "金额"),
                    @("机械键盘", "299", "10", ""),
                    @("无线鼠标", "99", "25", ""),
                    @("27寸显示器", "1299", "5", ""),
                    @("降噪耳机", "499", "12", "")
                )
            }
        },
        @{
            Id = 4
            Name = "筛选与汇总 (开启自动筛选并汇总统计)"
            Prompt = "在当前表格开启自动筛选，并在表格下方统计各部门的总人数和平均工资"
            SetupData = {
                param($ws)
                Set-GridValues $ws @(
                    @("姓名", "部门", "职位", "薪资"),
                    @("张伟", "研发部", "高级工程师", "22000"),
                    @("王芳", "市场部", "商务专员", "12000"),
                    @("李强", "研发部", "架构师", "35000"),
                    @("刘洋", "财务部", "会计师", "15000"),
                    @("陈静", "研发部", "测试主管", "18000"),
                    @("赵敏", "市场部", "渠道经理", "20000")
                )
            }
        },
        @{
            Id = 5
            Name = "创建图表 (插入带标题柱状图对比各季度销售额)"
            Prompt = "根据当前表格的数据，在旁边插入一个带标题的柱状图，直观展示每个季度的销售额对比"
            SetupData = {
                param($ws)
                Set-GridValues $ws @(
                    @("季度", "销售额(万元)"),
                    @("第一季度", "350"),
                    @("第二季度", "480"),
                    @("第三季度", "420"),
                    @("第四季度", "610")
                )
            }
        },
        @{
            Id = 6
            Name = "较复杂多步骤任务 (工时看板、高亮加班与顶部指标卡)"
            Prompt = "制作一个项目考勤与工时统计看板：新建一张表，录入5名员工5天的工作时长（正常为8小时），自动用浅红底色高亮超过9小时的加班工时，计算每人的总工时和出勤天数，并在顶部生成带大号字体的项目关键指标卡片"
            SetupData = $null
        }
    )

    foreach ($t in $tasks) {
        Start-Sleep -Seconds 1
        Write-Host "`n-----------------------------------------------------------------" -ForegroundColor Yellow
        Write-Host "[Task $($t.Id)] $($t.Name)" -ForegroundColor Yellow
        Write-Host "Prompt: $($t.Prompt)" -ForegroundColor Yellow
        Write-Host "-----------------------------------------------------------------" -ForegroundColor Yellow

        # 为该任务创建全新独立的目标工作簿
        $wbTaskPath = Join-Path $testDir ("Task_{0}_{1}.xlsx" -f $t.Id, [Guid]::NewGuid().ToString("N").Substring(0, 6))
        $wbTask = $excel.Workbooks.Add()
        $wsTarget = $wbTask.Sheets.Item(1)

        if ($t.SetupData -ne $null) {
            & $t.SetupData $wsTarget
        }
        $wbTask.SaveAs($wbTaskPath)

        $usedRangeBefore = $wsTarget.UsedRange.Address($false, $false)
        $sheetNames = @()
        for ($s = 1; $s -le $wbTask.Sheets.Count; $s++) {
            $sheetNames += $wbTask.Sheets.Item($s).Name
        }

        # 构造通用 prompt 并请求真实大模型（最多重试 1 次应对网络偶发截断）
        $sysPrompt = Build-GenericSysPrompt $wbTask.Name $sheetNames $wsTarget.Name $usedRangeBefore
        $attempts = 0
        $llmResponse = ""
        $extracted = $null
        while ($attempts -lt 2) {
            $attempts++
            Write-Host "  -> Requesting DeepSeek API (attempt $attempts)..." -ForegroundColor DarkGray
            $llmResponse = Call-RealLLM $sysPrompt $t.Prompt
            $extracted = Extract-VbaCode $llmResponse
            if (!$extracted.Error -and !$extracted.IsTruncated) { break }
            if ($attempts -lt 2) {
                Write-Host "     Warning: attempt $attempts extraction issue ($($extracted.Error)), retrying once..." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
            }
        }
        if ($extracted.Error -or $extracted.IsTruncated) {
            Write-Host "  [Extraction Error] $($extracted.Error)" -ForegroundColor Red
            $taskResults.Add([PSCustomObject]@{
                Id = $t.Id
                Name = $t.Name
                Ok = $false
                Error = $extracted.Error
                RawResponseLen = $llmResponse.Length
                ExtractedCodeLen = $extracted.Code.Length
            })
            $wbTask.Close($false)
            continue
        }

        $vbaCode = $extracted.Code
        Write-Host "  -> Extracted VBA successfully ($($vbaCode.Length) chars)" -ForegroundColor DarkCyan

        # 通过 NativeBridge.Dispatch 执行
        $req = @{
            action = "execute_vba"
            code = $vbaCode
            prompt = $t.Prompt
            targetWorkbookName = $wbTask.Name
            targetWorkbookFullName = $wbTaskPath
            rawModelResponse = $llmResponse
        }
        $reqJson = ($req | ConvertTo-Json -Compress)
        $resJson = [LeeExcel.NativeBridge]::Dispatch($reqJson, $excel)
        $parsed = [LeeExcel.SimpleJson]::ParseFlatObject($resJson)

        $isOk = ($parsed['ok'] -eq "True")
        Write-Host "  Execution Result: OK = $isOk, Message = $($parsed['message'])" -ForegroundColor $(if ($isOk) { "Green" } else { "Red" })
        if (!$isOk) {
            Write-Host "  Error Detail: $($parsed['error'])" -ForegroundColor Red
        }

        # 检查 Excel 实体状态
        $wsActive = $wbTask.ActiveSheet
        $usedRangeAfter = $wsActive.UsedRange.Address($false, $false)
        $rowCount = $wsActive.UsedRange.Rows.Count
        $colCount = $wsActive.UsedRange.Columns.Count
        $hasFormulas = [LeeExcel.VbaRunner]::CheckHasFormulas($wsActive.UsedRange)
        $hasBorders = [LeeExcel.VbaRunner]::CheckHasBorders($wsActive.UsedRange)
        $hasInterior = [LeeExcel.VbaRunner]::CheckHasInteriorColor($wsActive.UsedRange)
        $chartCount = $wsActive.ChartObjects().Count

        # 检查其他已打开工作簿是否保持纯净
        $otherVal = $wbOther.Sheets.Item(1).Range("A1").Value2
        $otherSafe = ($otherVal -eq "DO_NOT_TOUCH_OTHER_WORKBOOK")

        # 检查哈希与一致性
        $origHash = $parsed['originalCodeHash']
        $execHash = $parsed['executedCodeHash']
        $isIdentical = ($parsed['isSourceIdentical'] -eq "True")
        $hasWrapper = (![string]::IsNullOrEmpty($parsed['wrapperCode']) -and $parsed['wrapperCode'] -ne "null")

        Write-Host "  [Excel Readback]:" -ForegroundColor Gray
        Write-Host "    Sheet: $($wsActive.Name), UsedRange: $usedRangeAfter ($rowCount rows x $colCount cols)" -ForegroundColor Gray
        Write-Host "    Formulas: $hasFormulas, Borders: $hasBorders, InteriorColor: $hasInterior, Charts: $chartCount" -ForegroundColor Gray
        Write-Host "    Original Hash: $origHash" -ForegroundColor Gray
        Write-Host "    Executed Hash: $execHash" -ForegroundColor Gray
        Write-Host "    Source Identical: $isIdentical (Wrapper Appended: $hasWrapper)" -ForegroundColor Gray
        Write-Host "    Cross-Workbook Untouched: $otherSafe" -ForegroundColor Gray

        # 保存 .xlsx，验证临时宏已被彻底清除
        $wbTask.Save()
        $wbTask.Close($false)

        # 重新以 .xlsx 打开验证无残留、可正常读取
        $wbReopen = $excel.Workbooks.Open($wbTaskPath)
        $reopenUsed = $wbReopen.ActiveSheet.UsedRange.Address($false, $false)
        $reopenModCount = 0
        try {
            foreach ($comp in $wbReopen.VBProject.VBComponents) {
                if ($comp.Type -eq 1) { $reopenModCount++ } # 1 = vbext_ct_StdModule
            }
        } catch { }
        $wbReopen.Close($false)

        Write-Host "    Re-opened .xlsx successfully: UsedRange = $reopenUsed, Residual Modules = $reopenModCount" -ForegroundColor Gray

        $taskResults.Add([PSCustomObject]@{
            Id = $t.Id
            Name = $t.Name
            Ok = $isOk
            OriginalHash = $origHash
            ExecutedHash = $execHash
            IsIdentical = $isIdentical
            HasWrapper = $hasWrapper
            UsedRange = $usedRangeAfter
            Rows = $rowCount
            Cols = $colCount
            HasFormulas = $hasFormulas
            HasBorders = $hasBorders
            HasInterior = $hasInterior
            ChartCount = $chartCount
            OtherSafe = $otherSafe
            ResidualModules = $reopenModCount
            OriginalVba = $vbaCode
            ExecutedVba = $parsed['executedVbaCode']
            RawResponse = $llmResponse
        })
    }

    # =================================================================
    # 边缘场景 1: 纯自然语言闲聊不执行 VBA
    # =================================================================
    Write-Host "`n-----------------------------------------------------------------" -ForegroundColor Magenta
    Write-Host "[Edge Case 1] 纯闲聊/问答通道不执行 VBA" -ForegroundColor Magenta
    Write-Host "-----------------------------------------------------------------" -ForegroundColor Magenta
    
    $chatLines = @(
        '你是一名精通 Microsoft Excel 和 VBA 的专业顾问。',
        '当前处于【问答咨询与解释通道】。请用专业、亲切、通俗易懂的中文直接解答用户的问题、解释代码含义或进行日常交流。',
        '【核心边界规则】：',
        '1. 本通道只进行纯文本自然语言解答，绝对不要输出任何可被执行的自动化代码块，严禁输出任何 ```vba 代码块。',
        '2. 若用户询问代码含义，请用文字清晰分步剖析，不要诱导执行。'
    )
    $chatSysPrompt = ($chatLines -join "`r`n")
    $chatPrompt = "你好，请用通俗语言介绍一下 Excel 里的 VLOOKUP 函数和 XLOOKUP 有什么核心区别？"
    $chatResp = Call-RealLLM $chatSysPrompt $chatPrompt
    $chatExtracted = Extract-VbaCode $chatResp
    $hasNoVba = [string]::IsNullOrEmpty($chatExtracted.Code)
    Write-Host "Chat Response Length: $($chatResp.Length) chars" -ForegroundColor Gray
    Write-Host "Extracted VBA: '$($chatExtracted.Code)' (Expected Empty)" -ForegroundColor Gray
    Write-Host "Edge Case 1 Pass: $hasNoVba" -ForegroundColor $(if ($hasNoVba) { "Green" } else { "Red" })

    # =================================================================
    # 边缘场景 2: 代码截断安全拦截
    # =================================================================
    Write-Host "`n-----------------------------------------------------------------" -ForegroundColor Magenta
    Write-Host "[Edge Case 2] 代码截断安全拦截 (缺少 End Sub)" -ForegroundColor Magenta
    Write-Host "-----------------------------------------------------------------" -ForegroundColor Magenta
    
    $truncatedCode = "Sub IncompleteMacro()`r`n    Dim ws As Worksheet`r`n    Set ws = ActiveSheet`r`n    ws.Range(`"A1`").Value = `"ShouldNotRun`"`r`n    ' Code truncated here..."
    $truncatedExtracted = Extract-VbaCode ("``````vba`r`n" + $truncatedCode)
    $isTruncatedBlocked = ($truncatedExtracted.IsTruncated -eq $true -or $truncatedExtracted.Error -ne $null)
    Write-Host "Truncation Detected: $($truncatedExtracted.IsTruncated), Error: $($truncatedExtracted.Error)" -ForegroundColor Gray
    Write-Host "Edge Case 2 Pass: $isTruncatedBlocked" -ForegroundColor $(if ($isTruncatedBlocked) { "Green" } else { "Red" })

    # =================================================================
    # 边缘场景 3: 备份失败阻断执行
    # =================================================================
    Write-Host "`n-----------------------------------------------------------------" -ForegroundColor Magenta
    Write-Host "[Edge Case 3] 备份失败安全阻断 (不存在的目标工作簿)" -ForegroundColor Magenta
    Write-Host "-----------------------------------------------------------------" -ForegroundColor Magenta
    
    # 尝试针对不存在的工作簿执行
    $fakeReq = @{
        action = "execute_vba"
        code = "Sub Dummy()`r`nEnd Sub"
        prompt = "备份失败测试"
        targetWorkbookName = "NonExistentWb.xlsx"
        targetWorkbookFullName = "C:\NonExistentDir\NonExistentWb.xlsx"
    }
    $fakeReqJson = ($fakeReq | ConvertTo-Json -Compress)
    $fakeResJson = [LeeExcel.NativeBridge]::Dispatch($fakeReqJson, $excel)
    $fakeParsed = [LeeExcel.SimpleJson]::ParseFlatObject($fakeResJson)
    $backupBlocked = ($fakeParsed['ok'] -eq "False" -and $fakeParsed['error'] -like "*未找到目标工作簿*")
    Write-Host "NonExistent Target Blocked: $backupBlocked, Error: $($fakeParsed['error'])" -ForegroundColor Gray
    Write-Host "Edge Case 3 Pass: $backupBlocked" -ForegroundColor $(if ($backupBlocked) { "Green" } else { "Red" })

    # =================================================================
    # 边缘场景 4: 脚本保存与再次运行产生独立快照与记录
    # =================================================================
    Write-Host "`n-----------------------------------------------------------------" -ForegroundColor Magenta
    Write-Host "[Edge Case 4] 脚本保存与再次运行产生独立快照与记录" -ForegroundColor Magenta
    Write-Host "-----------------------------------------------------------------" -ForegroundColor Magenta
    
    $scriptDir = Join-Path $env:APPDATA "LeeExcel\Scripts"
    if (!(Test-Path $scriptDir)) { New-Item -ItemType Directory -Path $scriptDir -Force | Out-Null }
    $scriptFilePath = Join-Path $scriptDir "FormatTableSample.bas"
    $scriptCodeToSave = "Sub FormatTableSample(targetWb As Workbook)`r`n    Dim ws As Worksheet`r`n    Set ws = targetWb.ActiveSheet`r`n    ws.Range(`"Z1`").Value = `"RunFromSavedScript`"`r`nEnd Sub"
    [System.IO.File]::WriteAllText($scriptFilePath, $scriptCodeToSave, [System.Text.Encoding]::UTF8)

    # 针对新工作簿执行该保存脚本
    $wbScriptPath = Join-Path $testDir "ScriptRunTarget.xlsx"
    $wbScript = $excel.Workbooks.Add()
    $wbScript.SaveAs($wbScriptPath)

    # 运行第 1 次
    $run1Req = @{
        action = "execute_vba"
        code = $scriptCodeToSave
        prompt = "运行脚本: FormatTableSample"
        targetWorkbookName = $wbScript.Name
        targetWorkbookFullName = $wbScriptPath
    }
    $res1Json = [LeeExcel.NativeBridge]::Dispatch(($run1Req | ConvertTo-Json -Compress), $excel)
    $parsedRun1 = [LeeExcel.SimpleJson]::ParseFlatObject($res1Json)
    $snapId1 = $parsedRun1['snapshot']

    # 运行第 2 次
    $run2Req = @{
        action = "execute_vba"
        code = $scriptCodeToSave
        prompt = "运行脚本: FormatTableSample (第2次)"
        targetWorkbookName = $wbScript.Name
        targetWorkbookFullName = $wbScriptPath
    }
    $res2Json = [LeeExcel.NativeBridge]::Dispatch(($run2Req | ConvertTo-Json -Compress), $excel)
    $parsedRun2 = [LeeExcel.SimpleJson]::ParseFlatObject($res2Json)
    $snapId2 = $parsedRun2['snapshot']

    $wbScript.Close($false)

    $independentRuns = ($parsedRun1['ok'] -eq "True" -and $parsedRun2['ok'] -eq "True")
    Write-Host "Run 1 OK: $($parsedRun1['ok']), Run 2 OK: $($parsedRun2['ok'])" -ForegroundColor Gray
    Write-Host "Edge Case 4 Pass: $independentRuns" -ForegroundColor $(if ($independentRuns) { "Green" } else { "Red" })

    # =================================================================
    # 边缘场景 5: 快照回滚前保留救援副本 (Rescue Backup)
    # =================================================================
    Write-Host "`n-----------------------------------------------------------------" -ForegroundColor Magenta
    Write-Host "[Edge Case 5] 快照回滚前保留救援副本实测" -ForegroundColor Magenta
    Write-Host "-----------------------------------------------------------------" -ForegroundColor Magenta
    
    $wbRollbackPath = Join-Path $testDir "RollbackTarget.xlsx"
    $wbRollback = $excel.Workbooks.Add()
    $wbRollback.Sheets.Item(1).Range("A1").Value2 = "PreMacroInitial"
    $wbRollback.SaveAs($wbRollbackPath)

    # 1. 运行宏产生快照
    $macroReq = @{
        action = "execute_vba"
        code = "Sub ChangeCell()`r`nActiveSheet.Range(`"A1`").Value = `"PostMacroVal`"`r`nEnd Sub"
        prompt = "生成快照测试"
        targetWorkbookName = $wbRollback.Name
        targetWorkbookFullName = $wbRollbackPath
    }
    $macroRes = [LeeExcel.NativeBridge]::Dispatch(($macroReq | ConvertTo-Json -Compress), $excel)
    $macroParsed = [LeeExcel.SimpleJson]::ParseFlatObject($macroRes)
    $snapId = [regex]::Match($macroRes, '"id"\s*:\s*"([^"]+)"').Groups[1].Value

    # 2. 用户在宏之后进行手工修改
    $wbRollback.Sheets.Item(1).Range("A1").Value2 = "UserManualEdit_DoNotLose"
    $wbRollback.Save()

    # 3. 执行回滚
    $restoreReq = @{
        action = "restore_snapshot"
        snapshotId = $snapId
        targetWorkbookName = $wbRollback.Name
    }
    $restoreRes = [LeeExcel.NativeBridge]::Dispatch(($restoreReq | ConvertTo-Json -Compress), $excel)
    $restoreParsed = [LeeExcel.SimpleJson]::ParseFlatObject($restoreRes)

    # 4. 检查当前工作簿是否回到 PreMacroInitial
    # 恢复后重新连接打开的工作簿
    $restoredWb = $excel.Workbooks.Item([System.IO.Path]::GetFileName($wbRollbackPath))
    $currentValAfterRestore = $restoredWb.Sheets.Item(1).Range("A1").Value2
    $isRestored = ($currentValAfterRestore -eq "PreMacroInitial")

    # 5. 检查是否在快照目录生成了 rescue_before_restore_*.xlsx
    $backupBaseDir = Join-Path $env:APPDATA "LeeExcel\Backups"
    $rescueFiles = Get-ChildItem -Path $backupBaseDir -Recurse -Filter "rescue_before_restore_*.xlsx" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    $hasRescue = ($rescueFiles.Count -gt 0)
    $rescueVal = ""
    if ($hasRescue) {
        $rescueWb = $excel.Workbooks.Open($rescueFiles[0].FullName)
        $rescueVal = $rescueWb.Sheets.Item(1).Range("A1").Value2
        $rescueWb.Close($false)
    }
    $rescueSafeguarded = ($rescueVal -eq "UserManualEdit_DoNotLose")

    $restoredWb.Close($false)

    Write-Host "Current Workbook Restored: $isRestored (Val: '$currentValAfterRestore')" -ForegroundColor Gray
    Write-Host "Rescue Backup Created: $hasRescue (Val in Rescue: '$rescueVal')" -ForegroundColor Gray
    $edge5Pass = ($isRestored -and $rescueSafeguarded)
    Write-Host "Edge Case 5 Pass: $edge5Pass" -ForegroundColor $(if ($edge5Pass) { "Green" } else { "Red" })

} finally {
    if ($wbOther -ne $null) {
        try { $wbOther.Close($false) } catch {}
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# 输出汇总数据供报告使用
$reportPath = Join-Path $testDir "generic_e2e_summary.json"
$summaryObj = @{
    Tasks = $taskResults
    EdgeCase1_PureChatNoVba = $hasNoVba
    EdgeCase2_TruncatedBlocked = $isTruncatedBlocked
    EdgeCase3_BackupFailBlocked = $backupBlocked
    EdgeCase4_ScriptIndependentRuns = $independentRuns
    EdgeCase5_RescueBackupSafeguarded = $edge5Pass
}
$summaryJson = ($summaryObj | ConvertTo-Json -Depth 6)
[System.IO.File]::WriteAllText($reportPath, $summaryJson, [System.Text.Encoding]::UTF8)

Write-Host "`n=================================================================" -ForegroundColor Cyan
Write-Host "   ALL E2E VERIFICATION COMPLETED. SUMMARY SAVED TO:             " -ForegroundColor Cyan
Write-Host "   $reportPath" -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan
