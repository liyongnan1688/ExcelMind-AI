Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public class WinEnum {
    public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}
'@

$proc = Get-Process excel -ErrorAction SilentlyContinue | Select-Object -First 1
if (!$proc) {
    Write-Host "No Excel process found."
    exit
}

Write-Host "Excel PID: $($proc.Id)"
foreach ($t in $proc.Threads) {
    [WinEnum]::EnumThreadWindows($t.Id, {
        param($hwnd, $lparam)
        $cls = New-Object System.Text.StringBuilder 256
        $txt = New-Object System.Text.StringBuilder 256
        [WinEnum]::GetClassName($hwnd, $cls, 256) | Out-Null
        [WinEnum]::GetWindowText($hwnd, $txt, 256) | Out-Null
        $vis = [WinEnum]::IsWindowVisible($hwnd)
        Write-Host "HWND: $hwnd, Visible: $vis, Class: '$($cls.ToString())', Title: '$($txt.ToString())'"
        return $true
    }, [IntPtr]::Zero) | Out-Null
}
