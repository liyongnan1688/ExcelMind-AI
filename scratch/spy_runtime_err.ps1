Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public class DbgSpy {
    public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

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
$excel.DisplayAlerts = $false

$wb = $excel.Workbooks.Add()
$mod = $wb.VBProject.VBComponents.Add(1)
$mod.Name = "DivZeroMod"
$mod.CodeModule.AddFromString(@"
Sub RunDivZero()
    Dim x As Long, y As Long
    x = 10
    y = 0
    ActiveSheet.Range("A1").Value = x / y
End Sub
"@)

# 启动监视器
$cts = New-Object System.Threading.CancellationTokenSource
$task = [System.Threading.Tasks.Task]::Run([Action]{
    while (!$cts.Token.IsCancellationRequested) {
        [System.Threading.Thread]::Sleep(100)
        [DbgSpy]::EnumWindows({
            param($hwnd, $lparam)
            $cls = New-Object System.Text.StringBuilder 256
            $txt = New-Object System.Text.StringBuilder 256
            [DbgSpy]::GetClassName($hwnd, $cls, 256) | Out-Null
            [DbgSpy]::GetWindowText($hwnd, $txt, 256) | Out-Null
            $pid = 0
            [DbgSpy]::GetWindowThreadProcessId($hwnd, [ref]$pid)
            
            $p = Get-Process -Id $pid -ErrorAction SilentlyContinue
            if ($p -and $p.ProcessName -eq "EXCEL") {
                [Console]::WriteLine("Excel Window found: HWND=$hwnd, Class='$($cls.ToString())', Title='$($txt.ToString())'")
                if ($cls.ToString() -eq "#32770") {
                    [Console]::WriteLine("--> Dismissing #32770 dialog via WM_CLOSE...")
                    [DbgSpy]::PostMessage($hwnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)
                }
            }
            return $true
        }, [IntPtr]::Zero)
    }
})

try {
    Write-Host "Calling RunDivZero..."
    $excel.Run("DivZeroMod.RunDivZero")
    Write-Host "RunDivZero returned successfully!"
} catch {
    Write-Host "Caught expected error: $($_.Exception.Message)"
} finally {
    $cts.Cancel()
    try { $task.Wait(500) } catch {}
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
