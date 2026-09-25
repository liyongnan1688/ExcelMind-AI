# scratch/test_app_run_args.ps1
$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    
    # 1. Test String
    $codeString = @"
Sub TestString(arg As String)
    Range("A1").Value = "String:" & arg
End Sub
"@
    $comp = $wb.VBProject.VBComponents.Add(1)
    $comp.Name = "ModTest"
    $comp.CodeModule.AddFromString($codeString)

    try {
        $excel.Run("'" + $wb.Name + "'!TestString", "Hello")
        Write-Host "Test String: SUCCESS -> Cell A1 = $($wb.Sheets.Item(1).Range("A1").Value2)" -ForegroundColor Green
    } catch {
        Write-Host "Test String: FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }

    # 2. Test Workbook as Object
    $codeObj = @"
Sub TestObj(wb As Object)
    Range("A2").Value = "ObjTypeName:" & TypeName(wb)
End Sub
"@
    $comp.CodeModule.AddFromString($codeObj)
    try {
        $excel.Run("'" + $wb.Name + "'!TestObj", $wb)
        Write-Host "Test Obj (As Object): SUCCESS -> Cell A2 = $($wb.Sheets.Item(1).Range("A2").Value2)" -ForegroundColor Green
    } catch {
        Write-Host "Test Obj (As Object): FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }

    # 3. Test Workbook as Variant
    $codeVar = @"
Sub TestVar(wb As Variant)
    Range("A3").Value = "VarTypeName:" & TypeName(wb)
End Sub
"@
    $comp.CodeModule.AddFromString($codeVar)
    try {
        $excel.Run("'" + $wb.Name + "'!TestVar", $wb)
        Write-Host "Test Var (As Variant): SUCCESS -> Cell A3 = $($wb.Sheets.Item(1).Range("A3").Value2)" -ForegroundColor Green
    } catch {
        Write-Host "Test Var (As Variant): FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }

    # 4. Test Workbook as Workbook
    $codeWb = @"
Sub TestWb(wb As Workbook)
    Range("A4").Value = "WbTypeName:" & TypeName(wb)
End Sub
"@
    $comp.CodeModule.AddFromString($codeWb)
    try {
        $excel.Run("'" + $wb.Name + "'!TestWb", $wb)
        Write-Host "Test Wb (As Workbook): SUCCESS -> Cell A4 = $($wb.Sheets.Item(1).Range("A4").Value2)" -ForegroundColor Green
    } catch {
        Write-Host "Test Wb (As Workbook): FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }

    # 5. Test No Arguments
    $codeNoArg = @"
Sub TestNoArg()
    Range("A5").Value = "NoArgRan"
End Sub
"@
    $comp.CodeModule.AddFromString($codeNoArg)
    try {
        $excel.Run("'" + $wb.Name + "'!TestNoArg")
        Write-Host "Test NoArg: SUCCESS -> Cell A5 = $($wb.Sheets.Item(1).Range("A5").Value2)" -ForegroundColor Green
    } catch {
        Write-Host "Test NoArg: FAILED -> $($_.Exception.Message)" -ForegroundColor Red
    }

} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
