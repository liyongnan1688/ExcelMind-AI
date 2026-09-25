# scratch/test_dialog_dismiss.ps1
$ErrorActionPreference = "Stop"

Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Win32Helper {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    public const uint WM_COMMAND = 0x0111;
    public const uint WM_CLOSE = 0x0010;
}
"@

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Start a background job/timer to dismiss any VBE dialog if it appears
$timer = New-Object System.Timers.Timer
$timer.Interval = 200 # 200 ms
$timer.AutoReset = $true
Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action {
    $hwnd = [Win32Helper]::FindWindow("#32770", "Microsoft Visual Basic for Applications")
    if ($hwnd -ne [IntPtr]::Zero) {
        Write-Host "`n[Watchdog] Detected VBE Modal Dialog! Dismissing..." -ForegroundColor Yellow
        [Win32Helper]::SendMessage($hwnd, [Win32Helper]::WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
    }
} | Out-Null
$timer.Start()

try {
    $wb = $excel.Workbooks.Add()
    $comp = $wb.VBProject.VBComponents.Add(1)
    
    $code = @"
Sub TestBad()
    Msg = "Hi"
End CleanExit
End Sub
"@
    $comp.CodeModule.AddFromString($code)
    Write-Host "Code with syntax error added."

    Write-Host "Running app.Run (should be unblocked by watchdog if dialog appears)..."
    try {
        $excel.Run("'" + $wb.Name + "'!TestBad")
        Write-Host "Run finished!"
    } catch {
        Write-Host "Caught expected exception: $($_.Exception.Message)" -ForegroundColor Green
    }
} finally {
    $timer.Stop()
    $timer.Dispose()
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
}
