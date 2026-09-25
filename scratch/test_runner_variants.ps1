# scratch/test_runner_variants.ps1
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    
    # Test 1: Sub WithDoubleUnderscoreAndParam
    $code1 = @"
Sub TargetMacro()
    ActiveSheet.Range("B1").Value = "TargetRun"
End Sub

Sub __LeeHostRunner_0108(targetWb As Workbook)
    targetWb.Activate
    Call TargetMacro
End Sub
"@
    $comp1 = $wb.VBProject.VBComponents.Add(1)
    $comp1.Name = "Mod1"
    $comp1.CodeModule.AddFromString($code1)
    
    try {
        $excel.Run("'" + $wb.Name + "'!__LeeHostRunner_0108", $wb)
        Write-Host "Test 1 (__LeeHostRunner_0108 with param) succeeded!" -ForegroundColor Green
    } catch {
        Write-Host "Test 1 failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Test 2: Sub WithoutParam
    $code2 = @"
Sub LeeHostRunnerNoParam()
    Call TargetMacro
End Sub
"@
    $comp1.CodeModule.AddFromString($code2)
    try {
        $excel.Run("'" + $wb.Name + "'!LeeHostRunnerNoParam")
        Write-Host "Test 2 (LeeHostRunnerNoParam) succeeded!" -ForegroundColor Green
    } catch {
        Write-Host "Test 2 failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Test 3: Sub WithParam NormalName
    $code3 = @"
Sub LeeRunnerParam(targetWb As Workbook)
    targetWb.Activate
    Call TargetMacro
End Sub
"@
    $comp1.CodeModule.AddFromString($code3)
    try {
        $excel.Run("'" + $wb.Name + "'!LeeRunnerParam", $wb)
        Write-Host "Test 3 (LeeRunnerParam with param) succeeded!" -ForegroundColor Green
    } catch {
        Write-Host "Test 3 failed: $($_.Exception.Message)" -ForegroundColor Red
    }

} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
