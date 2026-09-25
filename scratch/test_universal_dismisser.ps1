Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public class WinDismisser {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
'@

$brokenCode = @"
Sub BrokenMacro()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Columns(1.ColumnWidth) = 10
End Sub
"@

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

$wb = $excel.Workbooks.Add()
$vbProj = $wb.VBProject
$mod = $vbProj.VBComponents.Add(1)
$mod.Name = "TestBrokenMod"
$mod.CodeModule.AddFromString($brokenCode)

$cts = New-Object System.Threading.CancellationTokenSource
$task = [System.Threading.Tasks.Task]::Run([Action]{
    while (!$cts.Token.IsCancellationRequested) {
        [System.Threading.Thread]::Sleep(50)
        [WinDismisser]::EnumWindows({
            param($hwnd, $lparam)
            $cls = New-Object System.Text.StringBuilder 256
            $txt = New-Object System.Text.StringBuilder 256
            [WinDismisser]::GetClassName($hwnd, $cls, 256) | Out-Null
            [WinDismisser]::GetWindowText($hwnd, $txt, 256) | Out-Null
            $c = $cls.ToString()
            $t = $txt.ToString()
            if ($c -eq "#32770" -or $t -like "*Visual Basic*") {
                [Console]::WriteLine("MATCHED DIALOG: Class='$c', Title='$t'")
                [WinDismisser]::PostMessage($hwnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) # WM_CLOSE
            }
            return $true
        }, [IntPtr]::Zero)
    }
})

Write-Host "Calling broken macro with universal dismisser..."
try {
    $excel.Run("TestBrokenMod.BrokenMacro")
    Write-Host "Success"
} catch {
    Write-Host "Caught expected exception: $($_.Exception.Message)"
} finally {
    $cts.Cancel()
    try { $task.Wait(300) } catch {}
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
