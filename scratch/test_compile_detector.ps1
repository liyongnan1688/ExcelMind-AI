$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()

function Test-CodeCompilation($name, $code) {
    $mod = $wb.VBProject.VBComponents.Add(1)
    $mod.Name = $name
    $mod.CodeModule.AddFromString($code)

    $btn = $excel.VBE.CommandBars.FindControl([System.Reflection.Missing]::Value, 578)
    $wasEnabled = $btn.Enabled
    if ($btn.Enabled) {
        $btn.Execute()
    }
    $isStillEnabled = $btn.Enabled
    $compiledOk = ($wasEnabled -and !$isStillEnabled)

    # 清理该测试模块
    $wb.VBProject.VBComponents.Remove($mod)

    return @{
        Name = $name
        CompiledOk = $compiledOk
        Before = $wasEnabled
        After = $isStillEnabled
    }
}

$cases = @(
    @{
        Name = "ValidBasic"
        Code = "Sub Test1()`r`n    Dim i As Long`r`n    i = 10`r`nEnd Sub"
        Expected = $true
    },
    @{
        Name = "ValidWithHelper"
        Code = "Sub Main()`r`n    Call Helper(10)`r`nEnd Sub`r`nSub Helper(x As Long)`r`n    Dim y As Long`r`n    y = x * 2`r`nEnd Sub"
        Expected = $true
    },
    @{
        Name = "BrokenSyntax"
        Code = "Sub Broken()`r`n    ws.Columns(1.ColumnWidth) = 10`r`nEnd Sub"
        Expected = $false
    },
    @{
        Name = "MissingEndIf"
        Code = "Sub BrokenIf()`r`n    If 1 = 1 Then`r`n    Dim x As Long`r`nEnd Sub"
        Expected = $false
    },
    @{
        Name = "Task4Fixed"
        Code = "Sub Task4Demo()`r`n    Dim d As Object`r`n    Set d = CreateObject(`"Scripting.Dictionary`")`r`n    d.Add `"A`", 1`r`nEnd Sub"
        Expected = $true
    }
)

foreach ($c in $cases) {
    $res = Test-CodeCompilation $c.Name $c.Code
    $match = ($res.CompiledOk -eq $c.Expected)
    Write-Host "Case [$($c.Name)]: CompiledOk = $($res.CompiledOk), Expected = $($c.Expected), Match = $match"
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
