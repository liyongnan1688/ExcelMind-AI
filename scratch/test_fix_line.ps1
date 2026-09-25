$task4Out = "C:\Users\35651\AppData\Local\Temp\LeeExcel_GenericE2E_615406b5\out_b3ad1627.txt"
$content = [System.IO.File]::ReadAllText($task4Out, [System.Text.Encoding]::UTF8)

$openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
$start = $openMatch.Index + $openMatch.Length
$closeIdx = $content.IndexOf('```', $start)
$vbaCode = $content.Substring($start, $closeIdx - $start).Trim()

# 替换笔误
$fixedCode = $vbaCode.Replace("ws.Columns(1.ColumnWidth)", "ws.Columns(1).ColumnWidth")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$wb = $excel.Workbooks.Add()
$ws = $wb.Sheets.Item(1)

# 写入测试数据
$rows = @(
    @("姓名", "部门", "职位", "薪资"),
    @("张伟", "研发部", "高级工程师", "22000"),
    @("王芳", "市场部", "商务专员", "12000"),
    @("李强", "研发部", "架构师", "35000"),
    @("刘洋", "财务部", "会计师", "15000"),
    @("陈静", "研发部", "测试主管", "18000"),
    @("赵敏", "市场部", "渠道经理", "20000")
)
for ($r = 0; $r -lt $rows.Count; $r++) {
    for ($c = 0; $c -lt $rows[$r].Count; $c++) {
        $ws.Cells.Item($r + 1, $c + 1).Value2 = $rows[$r][$c]
    }
}

$vbProj = $wb.VBProject
$mod = $vbProj.VBComponents.Add(1)
$mod.Name = "TestModule"
$mod.CodeModule.AddFromString($fixedCode)

Write-Host "Attempting to call Main with fixed code..."
$sw = [System.Diagnostics.Stopwatch]::StartNew()
try {
    $excel.Run("TestModule.Main")
    $sw.Stop()
    Write-Host "Main succeeded in $($sw.ElapsedMilliseconds) ms!"
    Write-Host "ActiveSheet UsedRange: $($ws.UsedRange.Address($false, $false))"
    for ($r = 1; $r -le $ws.UsedRange.Rows.Count; $r++) {
        $line = @()
        for ($c = 1; $c -le $ws.UsedRange.Columns.Count; $c++) {
            $line += $ws.Cells.Item($r, $c).Text
        }
        Write-Host "Row $r : $($line -join ' | ')"
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
