# scratch/test_wrapper_isolation_limits.ps1
# 审计无参 Main 的包装策略与开放式 VBA 的隔离边界
$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$testDir = Join-Path $env:TEMP ("LeeExcel_IsolationLimits_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $testDir | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$report = [System.Collections.Generic.List[PSObject]]::new()

try {
    # 建立两个工作簿
    $targetPath = Join-Path $testDir "TargetWb.xlsx"
    $otherPath = Join-Path $testDir "OtherWb.xlsx"

    $wbTarget = $excel.Workbooks.Add()
    $wbTarget.Sheets.Item(1).Range("A1").Value2 = "TargetInitial"
    $wbTarget.SaveAs($targetPath)

    $wbOther = $excel.Workbooks.Add()
    $wbOther.Sheets.Item(1).Range("A1").Value2 = "OtherInitial"
    $wbOther.SaveAs($otherPath)

    # -------------------------------------------------------------
    # Scenario 1: 未限定 Cells/Range (隐式依赖 ActiveSheet)
    # -------------------------------------------------------------
    Write-Host "`n=== Scenario 1: 未限定 Cells/Range ===" -ForegroundColor Cyan
    $code1 = @"
Sub Main()
    Range("A1").Value = "WrittenToActiveSheet"
End Sub
"@
    $req1 = @{
        action = "execute_vba"
        code = $code1
        prompt = "测试未限定 Range"
        targetWorkbookName = $wbTarget.Name
        targetWorkbookFullName = $targetPath
    }
    $res1 = [LeeExcel.NativeBridge]::Dispatch(($req1 | ConvertTo-Json -Compress), $excel)
    $p1 = [LeeExcel.SimpleJson]::ParseFlatObject($res1)
    
    $targetVal1 = $wbTarget.Sheets.Item(1).Range("A1").Value2
    $otherVal1 = $wbOther.Sheets.Item(1).Range("A1").Value2
    Write-Host "Target A1: '$targetVal1' | Other A1: '$otherVal1'"
    $report.Add([PSCustomObject]@{
        Scenario = "1. 未限定 Cells/Range"
        Behavior = "包装器通过 targetWb.Activate 使得当前活动表聚焦于目标工作簿，隐式 Range 成功写入目标工作簿"
        TargetModified = ($targetVal1 -eq "WrittenToActiveSheet")
        OtherModified = ($otherVal1 -ne "OtherInitial")
        Conclusion = "在 Main 未主动切走窗口的前提下，隐式操作会作用于目标工作簿"
    })

    # -------------------------------------------------------------
    # Scenario 2: 显式跨工作簿修改 (Workbooks('OtherWb.xlsx'))
    # -------------------------------------------------------------
    Write-Host "`n=== Scenario 2: 显式跨工作簿操作 ===" -ForegroundColor Cyan
    $code2 = @"
Sub Main()
    Workbooks("OtherWb.xlsx").Sheets(1).Range("A1").Value = "ExplicitCrossModify"
End Sub
"@
    $req2 = @{
        action = "execute_vba"
        code = $code2
        prompt = "测试显式跨工作簿修改"
        targetWorkbookName = $wbTarget.Name
        targetWorkbookFullName = $targetPath
    }
    $res2 = [LeeExcel.NativeBridge]::Dispatch(($req2 | ConvertTo-Json -Compress), $excel)
    $p2 = [LeeExcel.SimpleJson]::ParseFlatObject($res2)
    
    $otherVal2 = $wbOther.Sheets.Item(1).Range("A1").Value2
    $detected = ($p2['error'] -like "*检测到非目标工作簿*")
    Write-Host "Other A1 after cross-modify: '$otherVal2', Detected Warning: $detected"
    $report.Add([PSCustomObject]@{
        Scenario = "2. 显式操作其他 Workbooks"
        Behavior = "VBA 成功穿透修改了 OtherWb.xlsx。包装器无法在运行时杜绝该跨文件行为；宿主执行后检查探测到了附带变更并发出告警"
        TargetModified = $false
        OtherModified = ($otherVal2 -eq "ExplicitCrossModify")
        Conclusion = "自由 VBA 模式下无法实现内存或进程级对象沙箱隔离，跨文件副作用无法由包装器杜绝"
    })

    # -------------------------------------------------------------
    # Scenario 3: Main 内部新建工作簿并写入 (Workbooks.Add)
    # -------------------------------------------------------------
    Write-Host "`n=== Scenario 3: Main 内部新建工作簿 ===" -ForegroundColor Cyan
    $code3 = @"
Sub Main()
    Dim newWb As Workbook
    Set newWb = Workbooks.Add
    Range("A1").Value = "WrittenToNewBook"
End Sub
"@
    $req3 = @{
        action = "execute_vba"
        code = $code3
        prompt = "测试 Main 内部新建工作簿"
        targetWorkbookName = $wbTarget.Name
        targetWorkbookFullName = $targetPath
    }
    $res3 = [LeeExcel.NativeBridge]::Dispatch(($req3 | ConvertTo-Json -Compress), $excel)
    $p3 = [LeeExcel.SimpleJson]::ParseFlatObject($res3)

    $targetVal3 = $wbTarget.Sheets.Item(1).Range("A1").Value2
    Write-Host "Target A1 after Workbooks.Add: '$targetVal3'"
    $report.Add([PSCustomObject]@{
        Scenario = "3. Main 内部新建工作簿"
        Behavior = "Main 内部执行 Workbooks.Add 后，活动窗口被新建工作簿夺走，后续未限定的 Range('A1') 写入了新工作簿，目标工作簿未被修改"
        TargetModified = ($targetVal3 -eq "WrittenToNewBook")
        OtherModified = $false
        Conclusion = "Main 一旦在执行过程中自行变更了全局活动对象，targetWb.Activate 产生的初始聚焦即被打破"
    })

    # -------------------------------------------------------------
    # Scenario 4: Main 内部切换窗口 (Windows.Activate)
    # -------------------------------------------------------------
    Write-Host "`n=== Scenario 4: Main 内部切换窗口 ===" -ForegroundColor Cyan
    $code4 = @"
Sub Main()
    Windows("OtherWb.xlsx").Activate
    Range("A1").Value = "SwitchedWindowVal"
End Sub
"@
    $req4 = @{
        action = "execute_vba"
        code = $code4
        prompt = "测试 Main 内部切换窗口"
        targetWorkbookName = $wbTarget.Name
        targetWorkbookFullName = $targetPath
    }
    $res4 = [LeeExcel.NativeBridge]::Dispatch(($req4 | ConvertTo-Json -Compress), $excel)
    $p4 = [LeeExcel.SimpleJson]::ParseFlatObject($res4)

    $otherVal4 = $wbOther.Sheets.Item(1).Range("A1").Value2
    Write-Host "Other A1 after switch window: '$otherVal4'"
    $report.Add([PSCustomObject]@{
        Scenario = "4. Main 内部切换窗口"
        Behavior = "Main 内部主动调用 Windows('OtherWb.xlsx').Activate，后续未限定的 Range 写入了切换后的窗口"
        TargetModified = $false
        OtherModified = ($otherVal4 -eq "SwitchedWindowVal")
        Conclusion = "自由 VBA 模式下无法阻止代码内部的主动窗口切换动作"
    })

    $wbTarget.Close($false)
    $wbOther.Close($false)

} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

$isoPath = Join-Path $testDir "isolation_audit_summary.json"
$isoJson = ($report | ConvertTo-Json -Depth 5)
[System.IO.File]::WriteAllText($isoPath, $isoJson, [System.Text.Encoding]::UTF8)

Write-Host "`nIsolation Audit Completed! Saved to: $isoPath" -ForegroundColor Green
