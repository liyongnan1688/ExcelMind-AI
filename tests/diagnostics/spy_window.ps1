Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public class WinSpy {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
'@

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excelPid = 0
# 获取 Excel 进程 ID
$wb = $excel.Workbooks.Add()

$code = @"
Sub BrokenMacro()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Columns(1.ColumnWidth) = 10
End Sub
"@

$vbProj = $wb.VBProject
$mod = $vbProj.VBComponents.Add(1)
$mod.Name = "TestSyntaxErr"
$mod.CodeModule.AddFromString($code)

$cts = New-Object System.Threading.CancellationTokenSource
$task = [System.Threading.Tasks.Task]::Run([Action]{
    while (!$cts.Token.IsCancellationRequested) {
        [System.Threading.Thread]::Sleep(100)
        [WinSpy]::EnumWindows({
            param($hwnd, $lparam)
            $cls = New-Object System.Text.StringBuilder 256
            $txt = New-Object System.Text.StringBuilder 256
            [WinSpy]::GetClassName($hwnd, $cls, 256) | Out-Null
            [WinSpy]::GetWindowText($hwnd, $txt, 256) | Out-Null
            if ($cls.ToString() -eq "#32770") {
                $pId = 0
                [WinSpy]::GetWindowThreadProcessId($hwnd, [ref]$pId)
                [Console]::WriteLine("FOUND #32770: Title='$($txt.ToString())', PID=$pId")
                # 尝试关闭
                [WinSpy]::PostMessage($hwnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) # WM_CLOSE
            }
            return $true
        }, [IntPtr]::Zero)
    }
})

try {
    Write-Host "Running broken macro..."
    $excel.Run("TestSyntaxErr.BrokenMacro")
    Write-Host "Excel.Run completed without exception"
} catch {
    Write-Host "Excel.Run threw: $($_.Exception.Message)"
} finally {
    $cts.Cancel()
    try { $task.Wait(500) } catch {}
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
