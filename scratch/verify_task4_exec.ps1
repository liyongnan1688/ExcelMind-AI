Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$task4Text = Get-Content -Raw "C:\Users\35651\.gemini\antigravity-ide\brain\d0dc65bc-8cb8-4b59-b2a4-21058974d6eb\.system_generated\tasks\task-1718.log"

# 提取代码
$openMatch = [regex]::Match($task4Text, '```(?:vba|vb)?\s*')
$start = $openMatch.Index + $openMatch.Length
$closeIdx = $task4Text.IndexOf('```', $start)
$vbaCode = $task4Text.Substring($start, $closeIdx - $start).Trim()

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$testDir = "C:\Users\35651\AppData\Local\Temp\LeeExcel_Task4_Verify"
if (!(Test-Path $testDir)) { New-Item -ItemType Directory -Path $testDir | Out-Null }
$wbPath = Join-Path $testDir "Task4_Target.xlsx"

$wb = $excel.Workbooks.Add()
$ws = $wb.Sheets.Item(1)

function Set-GridValues($ws, $rows) {
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $row = $rows[$r]
        for ($c = 0; $c -lt $row.Count; $c++) {
            $ws.Cells.Item($r + 1, $c + 1).Value2 = $row[$c]
        }
    }
}

Set-GridValues $ws @(
    @("姓名", "部门", "职位", "薪资"),
    @("张伟", "研发部", "高级工程师", "22000"),
    @("王芳", "市场部", "商务专员", "12000"),
    @("李强", "研发部", "架构师", "35000"),
    @("刘洋", "财务部", "会计师", "15000"),
    @("陈静", "研发部", "测试主管", "18000"),
    @("赵敏", "市场部", "渠道经理", "20000")
)
$wb.SaveAs($wbPath)

Write-Host "Calling NativeBridge.Dispatch for Task 4..."
$req = @{
    action = "execute_vba"
    code = $vbaCode
    prompt = "在当前表格开启自动筛选，并在表格下方统计各部门的总人数和平均工资"
    targetWorkbookName = $wb.Name
    targetWorkbookFullName = $wbPath
    rawModelResponse = $task4Text
}
$resJson = [LeeExcel.NativeBridge]::Dispatch(($req | ConvertTo-Json -Compress), $excel)
$parsed = [LeeExcel.SimpleJson]::ParseFlatObject($resJson)

Write-Host "Result OK: $($parsed['ok']), Message: $($parsed['message'])"
if ($parsed['ok'] -eq "True") {
    $wsActive = $wb.ActiveSheet
    Write-Host "UsedRange: $($wsActive.UsedRange.Address($false, $false))"
    for ($r = 1; $r -le $wsActive.UsedRange.Rows.Count; $r++) {
        $line = @()
        for ($c = 1; $c -le $wsActive.UsedRange.Columns.Count; $c++) {
            $line += $wsActive.Cells.Item($r, $c).Text
        }
        Write-Host "Row $r : $($line -join ' | ')"
    }
} else {
    Write-Host "Error: $($parsed['error'])"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
