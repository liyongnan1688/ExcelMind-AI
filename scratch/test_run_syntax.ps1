$ErrorActionPreference = "Stop"

Add-Type -Path "bin\ExcelDna.Integration.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.Core.dll"
Add-Type -Path "bin\Microsoft.Web.WebView2.WinForms.dll"
Add-Type -Path "bin\LeeExcel.dll"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$mod = $wb.VBProject.VBComponents.Add(1)
$mod.Name = "MyMod"
$mod.CodeModule.AddFromString(@"
Sub TestProc(targetWb As Workbook)
    targetWb.Sheets(1).Range("A1").Value = "RunSuccess"
End Sub
"@)

Write-Host "Workbook Name: $($wb.Name)"

# Option 1: Module.Procedure
try {
    Write-Host "1. Testing 'MyMod.TestProc'..."
    $excel.Run("MyMod.TestProc", $wb)
    Write-Host "   SUCCESS: MyMod.TestProc worked!"
} catch {
    Write-Host "   FAILED: $($_.Exception.Message)"
}

# Option 2: Procedure
try {
    Write-Host "2. Testing 'TestProc'..."
    $excel.Run("TestProc", $wb)
    Write-Host "   SUCCESS: TestProc worked!"
} catch {
    Write-Host "   FAILED: $($_.Exception.Message)"
}

# Option 3: 'WorkbookName'!Module.Procedure
try {
    Write-Host "3. Testing ''$($wb.Name)'!MyMod.TestProc'..."
    $excel.Run("'" + $wb.Name + "'!MyMod.TestProc", $wb)
    Write-Host "   SUCCESS: ''$($wb.Name)'!MyMod.TestProc' worked!"
} catch {
    Write-Host "   FAILED: $($_.Exception.Message)"
}

# Option 4: 'WorkbookName'!Procedure
try {
    Write-Host "4. Testing ''$($wb.Name)'!TestProc'..."
    $excel.Run("'" + $wb.Name + "'!TestProc", $wb)
    Write-Host "   SUCCESS: ''$($wb.Name)'!TestProc' worked!"
} catch {
    Write-Host "   FAILED: $($_.Exception.Message)"
}

$wb.Close($false)
$excel.Quit()
