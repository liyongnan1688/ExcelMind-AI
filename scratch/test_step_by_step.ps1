$task4Out = "C:\Users\35651\AppData\Local\Temp\LeeExcel_GenericE2E_615406b5\out_b3ad1627.txt"
$content = [System.IO.File]::ReadAllText($task4Out, [System.Text.Encoding]::UTF8)

$openMatch = [regex]::Match($content, '```(?:vba|vb)?\s*')
$start = $openMatch.Index + $openMatch.Length
$closeIdx = $content.IndexOf('```', $start)
$vbaCode = $content.Substring($start, $closeIdx - $start).Trim()

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true   # 设为 true 看看前台到底发生了什么！

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

Write-Host "Step 1: Adding module..."
$vbProj = $wb.VBProject
$mod = $vbProj.VBComponents.Add(1)
$mod.Name = "TestModule"

Write-Host "Step 2: Adding code..."
$mod.CodeModule.AddFromString($vbaCode)

Write-Host "Step 3: Checking macro in code..."
Write-Host "Step 4: Attempting to call Main..."
try {
    # 尝试在 COM 中调用
    $excel.Run("TestModule.Main")
    Write-Host "Step 5: Run returned successfully!"
} catch {
    Write-Host "Step 5: Run threw: $($_.Exception.ToString())"
} finally {
    Write-Host "Finished step 5."
}
