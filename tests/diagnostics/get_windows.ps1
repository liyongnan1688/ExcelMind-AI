Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public class WinLister {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}
'@

[WinLister]::EnumWindows({
    param($hwnd, $lparam)
    if ([WinLister]::IsWindowVisible($hwnd)) {
        $cls = New-Object System.Text.StringBuilder 256
        $txt = New-Object System.Text.StringBuilder 256
        [WinLister]::GetClassName($hwnd, $cls, 256) | Out-Null
        [WinLister]::GetWindowText($hwnd, $txt, 256) | Out-Null
        $pid = 0
        [WinLister]::GetWindowThreadProcessId($hwnd, [ref]$pid)
        $proc = Get-Process -Id $pid -ErrorAction SilentlyContinue
        $pName = if ($proc) { $proc.ProcessName } else { "Unknown" }
        if ($pName -eq "EXCEL" -or $cls.ToString() -eq "#32770") {
            Write-Host "HWND: $hwnd | Process: $pName ($pid) | Class: '$($cls.ToString())' | Title: '$($txt.ToString())'"
        }
    }
    return $true
}, [IntPtr]::Zero) | Out-Null
