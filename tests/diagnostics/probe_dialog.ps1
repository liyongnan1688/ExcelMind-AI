$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $wb = $excel.Workbooks.Add()
    $c = $wb.VBProject.VBComponents.Add(1)
    $c.CodeModule.AddFromString("Sub TestIf(targetWb As Workbook)`r`nIf 1 = 1 Then`r`nDim a As Long`r`nEnd Sub`r`n")
    
    # 异步触发 app.Run
    $job = Start-Job -ScriptBlock {
        param($modName, $wbPath)
        $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $xl.Run($modName + ".TestIf")
    } -ArgumentList $c.Name, $wb.FullName
    
    Start-Sleep -Seconds 1
    
    # 查找弹出的窗口
    Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class WinProbe {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
"@

    [WinProbe]::EnumWindows({
        param($hWnd, $lParam)
        $sbC = New-Object System.Text.StringBuilder 256
        $sbT = New-Object System.Text.StringBuilder 256
        [WinProbe]::GetClassName($hWnd, $sbC, 256) | Out-Null
        [WinProbe]::GetWindowText($hWnd, $sbT, 256) | Out-Null
        $cls = $sbC.ToString()
        $tit = $sbT.ToString()
        if ($cls -eq "#32770" -or $tit -like "*Visual Basic*") {
            $pidVal = 0
            [WinProbe]::GetWindowThreadProcessId($hWnd, [ref]$pidVal) | Out-Null
            Write-Host "Found Dialog -> hWnd: $hWnd, Class: $cls, Title: $tit, PID: $pidVal"
            # 尝试关闭它
            [WinProbe]::SendMessage($hWnd, 0x0111, [IntPtr]2, [IntPtr]0) | Out-Null
            [WinProbe]::SendMessage($hWnd, 0x0111, [IntPtr]1, [IntPtr]0) | Out-Null
            [WinProbe]::PostMessage($hWnd, 0x0010, [IntPtr]0, [IntPtr]0) | Out-Null
        }
        return $true
    }, [IntPtr]::Zero) | Out-Null
    
    $job | Wait-Job -Timeout 3
    $job | Remove-Job -Force
} finally {
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
