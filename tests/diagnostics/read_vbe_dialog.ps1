<#
.SYNOPSIS
    Diagnose Excel VBE dialog message (Capture #32770 dialog triggered by CommandBar 578).

.DESCRIPTION
    Requires caller to explicitly specify target workbook and VBA code snippet.
    Fails before creating any Excel COM object if inputs are missing.
    Outputs run logs and captured dialog to .artifacts/tests/ or specified OutputDir.
#>
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath = "",

    [Parameter(Mandatory = $false)]
    [string]$VbaSnippetPath = "",

    [Parameter(Mandatory = $false)]
    [string]$OutputDir = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# 1. Resolve project root
$scriptDir = $PSScriptRoot
$projectRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if (-not (Test-Path (Join-Path $projectRoot "tests"))) {
    $projectRoot = (Get-Location).Path
}

# 2. Check input arguments before any Excel/COM interaction
if ([string]::IsNullOrWhiteSpace($WorkbookPath) -or [string]::IsNullOrWhiteSpace($VbaSnippetPath)) {
    Write-Host "[ERROR] Missing required arguments: -WorkbookPath and -VbaSnippetPath must be specified." -ForegroundColor Red
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  powershell -NoProfile -File tests/diagnostics/read_vbe_dialog.ps1 -WorkbookPath <path-to-xlsx> -VbaSnippetPath <path-to-vba> [-OutputDir <path>]"
    Write-Host "Note: Historical scratch sample generated_split_real.txt and desktop wildcard search have been removed."
    exit 1
}

$resolvedWb = if ([System.IO.Path]::IsPathRooted($WorkbookPath)) { $WorkbookPath } else { Join-Path (Get-Location).Path $WorkbookPath }
if (-not (Test-Path -LiteralPath $resolvedWb)) {
    $resolvedWb = Join-Path $projectRoot $WorkbookPath
}
if (-not (Test-Path -LiteralPath $resolvedWb)) {
    Write-Error "[FATAL ERROR] Workbook does not exist: '$WorkbookPath'"
    exit 1
}

$resolvedVba = if ([System.IO.Path]::IsPathRooted($VbaSnippetPath)) { $VbaSnippetPath } else { Join-Path (Get-Location).Path $VbaSnippetPath }
if (-not (Test-Path -LiteralPath $resolvedVba)) {
    $resolvedVba = Join-Path $projectRoot $VbaSnippetPath
}
if (-not (Test-Path -LiteralPath $resolvedVba)) {
    Write-Error "[FATAL ERROR] VBA snippet file does not exist: '$VbaSnippetPath'"
    exit 1
}

# 3. Resolve and initialize output directory
$toolName = "read_vbe_dialog"
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $runId = "diag_vbe_dialog_" + (Get-Date -Format "yyyyMMdd_HHmmss") + "_" + [System.Guid]::NewGuid().ToString("N").Substring(0, 6)
    $OutputDir = Join-Path $projectRoot (Join-Path ".artifacts" (Join-Path "tests" $runId))
} elseif (-not [System.IO.Path]::IsPathRooted($OutputDir)) {
    $OutputDir = Join-Path $projectRoot $OutputDir
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}
$lockFile = Join-Path $OutputDir ".running"
[System.IO.File]::WriteAllText($lockFile, (Get-Date -Format "o"), [System.Text.Encoding]::UTF8)

# 4. Add Win32 PInvoke
$csharpCode = @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class WinApiDiag {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
}
"@
Add-Type -TypeDefinition $csharpCode -ErrorAction SilentlyContinue

$status = "failed"
$details = ""
$excel = $null
$wb = $null

try {
    # 5. Start Excel COM and run diagnosis
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true

    $wb = $excel.Workbooks.Open($resolvedWb, $false, $true)
    $proj = $wb.VBProject

    $code = Get-Content -LiteralPath $resolvedVba -Raw
    if ($code -match '```vba([\s\S]+?)```') {
        $vba = $matches[1].Trim()
    } else {
        $vba = $code
    }

    $mod = $proj.VBComponents.Add(1)
    $mod.Name = "ModDiagnose"
    $mod.CodeModule.AddFromString($vba)

    $wb.Activate()
    $excel.VBE.ActiveVBProject = $proj
    $mod.Activate()

    $btn = $excel.VBE.CommandBars.FindControl([Type]::Missing, 578)
    Write-Host "Btn 578 Enabled before: $($btn.Enabled)"

    $job = Start-Job -ScriptBlock {
        $csharpJobCode = @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class WinApi2Diag {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
"@
        Add-Type -TypeDefinition $csharpJobCode -ErrorAction SilentlyContinue

        for ($i = 0; $i -lt 30; $i++) {
            Start-Sleep -Milliseconds 200
            $hwnd = [WinApi2Diag]::FindWindow("#32770", "Microsoft Visual Basic for Applications")
            if ($hwnd -ne [IntPtr]::Zero) {
                $staticHwnd = [WinApi2Diag]::FindWindowEx($hwnd, [IntPtr]::Zero, "Static", $null)
                $staticHwnd2 = [WinApi2Diag]::FindWindowEx($hwnd, $staticHwnd, "Static", $null)
                $sb = New-Object System.Text.StringBuilder 1024
                [WinApi2Diag]::GetWindowText($staticHwnd2, $sb, 1024)
                $msg = $sb.ToString()
                [WinApi2Diag]::PostMessage($hwnd, 0x0111, [IntPtr]2, [IntPtr]::Zero)
                return $msg
            }
        }
        return "No dialog found"
    }

    $btn.Execute()
    Write-Host "Btn 578 Enabled after: $($btn.Enabled)"

    $dialogResult = Receive-Job $job -Wait
    Write-Host "Dialog Text: $dialogResult"
    Remove-Job $job -Force

    Set-Content -LiteralPath (Join-Path $OutputDir "dialog_result.txt") -Value $dialogResult -Encoding UTF8
    $status = "completed"
    $details = "Dialog text captured: $dialogResult"
} catch {
    $details = "Exception: " + $_.Exception.Message
    Write-Error $_
} finally {
    if ($wb -ne $null) {
        try { $wb.Close($false) } catch { }
    }
    if ($excel -ne $null) {
        try { $excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }

    try {
        $meta = @{
            runId     = Split-Path -Leaf $OutputDir
            toolName  = $toolName
            timestamp = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
            status    = $status
            details   = $details
        } | ConvertTo-Json -Depth 3
        Set-Content -LiteralPath (Join-Path $OutputDir "meta.json") -Value $meta -Encoding UTF8
    } catch { }

    if (Test-Path -LiteralPath $lockFile) {
        try { Remove-Item -LiteralPath $lockFile -Force } catch { }
    }
}
