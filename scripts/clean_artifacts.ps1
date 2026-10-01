# ==============================================================================
# Lee-Excel Artifacts Cleaning Entry
# Path: scripts/clean_artifacts.ps1
#
# Safety Guarantees:
# 1. Preview by default. Requires explicit -Apply switch to delete.
# 2. Hardcoded target scope: ONLY .artifacts/tests/ and .artifacts/tmp/.
# 3. Strictly forbids touching src/, web/, bin/, tests/, scratch/, docs/, .backups, or .backup_*.
# 4. Canonicalizes paths, blocks symlinks, reparse points, and path traversal (..).
# 5. Skips runs marked with .running lockfile.
# 6. Does NOT terminate processes or force-kill open handles.
# 7. Exact RunId Matching: -RunId strictly matches an exact directory name (no wildcards/patterns).
#    Bypasses the 14-day retention requirement when specified, but still requires valid metadata
#    and strictly enforces all safety gates (.running, non-symlink, within scope).
#    Rejects ambiguous RunIds matching across multiple scopes (tests/ and tmp/) without -Scope.
# ==============================================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [int]$RetentionDays = 14,

    [Parameter(Mandatory = $false)]
    [string]$RunId = "",

    [Parameter(Mandatory = $false)]
    [ValidateSet("All", "Tests", "Tmp")]
    [string]$Scope = "All",

    [Parameter(Mandatory = $false)]
    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ProjectRoot = (Resolve-Path (Join-Path $ScriptDir "..")).Path
$ArtifactsRoot = (Join-Path $ProjectRoot ".artifacts")

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Lee-Excel Artifacts Safety Cleaner" -ForegroundColor Cyan
Write-Host " Mode:           $(if ($Apply) { 'APPLY (Physical Deletion)' } else { 'PREVIEW ONLY (Dry-Run)' })" -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'Green' })
Write-Host " Retention Days: $RetentionDays days" -ForegroundColor Gray
Write-Host " Specific RunId: $(if ($RunId) { $RunId } else { '<None - Scan by Age>' })" -ForegroundColor Gray
Write-Host " Scope:          $Scope" -ForegroundColor Gray
Write-Host " Artifacts Root: $ArtifactsRoot" -ForegroundColor Gray
Write-Host "==========================================================" -ForegroundColor Cyan

if (-not (Test-Path -LiteralPath $ArtifactsRoot)) {
    Write-Host "No .artifacts directory found. Nothing to clean." -ForegroundColor Gray
    exit 0
}

# Validate RunId if specified
if ($RunId -ne "") {
    if ($RunId -match '[\*\\/?:<>|]' -or $RunId -match '\.\.') {
        Write-Error "CRITICAL: RunId contains invalid characters or wildcards. Exact alphanumeric run ID required: '$RunId'"
        exit 1
    }
    if (-not ($RunId -match '^[a-zA-Z0-9_\-\.]+$')) {
        Write-Error "CRITICAL: RunId must consist of alphanumeric characters, underscores, hyphens, and dots: '$RunId'"
        exit 1
    }
}

# Determine target directories
$targetRoots = @()
if ($Scope -eq "All" -or $Scope -eq "Tests") {
    $testsDir = Join-Path $ArtifactsRoot "tests"
    if (Test-Path -LiteralPath $testsDir) { $targetRoots += (Resolve-Path -LiteralPath $testsDir).Path }
}
if ($Scope -eq "All" -or $Scope -eq "Tmp") {
    $tmpDir = Join-Path $ArtifactsRoot "tmp"
    if (Test-Path -LiteralPath $tmpDir) { $targetRoots += (Resolve-Path -LiteralPath $tmpDir).Path }
}

$cutoffTime = (Get-Date).AddDays(-$RetentionDays)
$candidates = [System.Collections.Generic.List[PSCustomObject]]::new()
$skippedRunning = [System.Collections.Generic.List[PSCustomObject]]::new()
$skippedNoMeta = [System.Collections.Generic.List[PSCustomObject]]::new()
$skippedReparse = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($rootPath in $targetRoots) {
    # Security check: ensure rootPath is strictly inside .artifacts
    if (-not $rootPath.StartsWith($ArtifactsRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        Write-Error "CRITICAL: Attempted access outside of .artifacts root: $rootPath"
        exit 1
    }

    $subDirs = Get-ChildItem -LiteralPath $rootPath -Directory
    foreach ($dir in $subDirs) {
        $fullPath = $dir.FullName

        # 1. Reject ReparsePoint (Symlink / Junction)
        if ($dir.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            $skippedReparse.Add([PSCustomObject]@{
                Path   = $fullPath
                Reason = "Symlink or ReparsePoint rejected for safety"
            })
            continue
        }

        # 2. Check if marked as running
        $lockFile = Join-Path $fullPath ".running"
        if (Test-Path -LiteralPath $lockFile) {
            $skippedRunning.Add([PSCustomObject]@{
                Path   = $fullPath
                Reason = "Active run lockfile found (.running)"
            })
            continue
        }

        # 3. Check RunId filter: MUST be exact match (no wildcards, no pattern matching)
        if ($RunId -ne "") {
            if ($dir.Name -ne $RunId) {
                continue
            }
            # When explicit RunId is provided, age check is bypassed, but all other safety checks still apply
        } else {
            # Check age
            if ($dir.LastWriteTime -gt $cutoffTime) {
                continue
            }
        }

        # 4. Check metadata or known timestamp pattern (e.g. yyyyMMdd_... or run_... or GUID)
        $hasMeta = (Test-Path (Join-Path $fullPath "meta.json")) -or 
                   (Test-Path (Join-Path $fullPath "run.json")) -or 
                   ($dir.Name -match '^\d{8}_\d{6}') -or 
                   ($dir.Name -match '^[0-9a-fA-F-]{36}$') -or 
                   ($dir.Name.StartsWith("run_") -or $dir.Name.StartsWith("session_") -or $dir.Name.StartsWith("wb_binding_"))

        if (-not $hasMeta) {
            $skippedNoMeta.Add([PSCustomObject]@{
                Path   = $fullPath
                Reason = "Lacks standard run metadata or recognizable timestamp naming convention"
            })
            continue
        }

        # Calculate directory size
        $files = @(Get-ChildItem -LiteralPath $fullPath -Recurse -File -Force -ErrorAction SilentlyContinue)
        $sizeBytes = 0
        foreach ($f in $files) {
            $sizeBytes += $f.Length
        }

        $candidates.Add([PSCustomObject]@{
            DirectoryName = $dir.Name
            FullPath      = $fullPath
            LastWriteTime = $dir.LastWriteTime
            FileCount     = $files.Count
            SizeBytes     = $sizeBytes
        })
    }
}

# Ambiguity check for RunId: Refuse to delete if same ID exists in both tests and tmp
if ($RunId -ne "" -and $candidates.Count -gt 1) {
    $ambiguousPaths = ($candidates | ForEach-Object { $_.FullPath }) -join "`n"
    Write-Error "AMBIGUOUS RUN ID: The exact RunId '$RunId' matches $($candidates.Count) directories across different scopes:`n$ambiguousPaths`nSpecify -Scope Tests or -Scope Tmp to disambiguate. Halting."
    exit 1
}

# Output findings
Write-Host "`n[Scan Results]" -ForegroundColor Cyan
Write-Host "Total Candidate Directories to Clean: $($candidates.Count)" -ForegroundColor $(if ($candidates.Count -gt 0) { 'Yellow' } else { 'Green' })

$totalBytes = 0
foreach ($cand in $candidates) {
    $totalBytes += $cand.SizeBytes
    $kb = [math]::Round($cand.SizeBytes / 1024, 2)
    Write-Host " - [$($cand.DirectoryName)] ($($cand.FileCount) files, $kb KB, Modified: $($cand.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')))"
}

if ($skippedRunning.Count -gt 0) {
    Write-Host "`n[Skipped Running Runs] ($($skippedRunning.Count)):" -ForegroundColor Yellow
    foreach ($s in $skippedRunning) { Write-Host " - $($s.Path): $($s.Reason)" -ForegroundColor Yellow }
}

if ($skippedNoMeta.Count -gt 0) {
    Write-Host "`n[Reported Non-Standard Directories (Not Deleted)] ($($skippedNoMeta.Count)):" -ForegroundColor DarkYellow
    foreach ($s in $skippedNoMeta) { Write-Host " - $($s.Path): $($s.Reason)" -ForegroundColor DarkYellow }
}

if ($skippedReparse.Count -gt 0) {
    Write-Host "`n[Skipped Symlinks / Reparse Points] ($($skippedReparse.Count)):" -ForegroundColor Red
    foreach ($s in $skippedReparse) { Write-Host " - $($s.Path): $($s.Reason)" -ForegroundColor Red }
}

Write-Host "----------------------------------------------------------" -ForegroundColor Gray
Write-Host "Total Reclaimable Size: $([math]::Round($totalBytes / 1024, 2)) KB" -ForegroundColor Cyan

if (-not $Apply) {
    Write-Host "`nPREVIEW COMPLETED. No files were deleted." -ForegroundColor Green
    Write-Host "To execute physical deletion, re-run with: .\scripts\clean_artifacts.ps1 -Apply" -ForegroundColor Yellow
    exit 0
}

# Execution Phase
Write-Host "`n[Executing Deletion]" -ForegroundColor Red
$deletedCount = 0
$failedCount = 0

foreach ($cand in $candidates) {
    try {
        Remove-Item -LiteralPath $cand.FullPath -Recurse -Force -ErrorAction Stop
        Write-Host " [DELETED] $($cand.FullPath)" -ForegroundColor Green
        $deletedCount++
    } catch {
        Write-Warning " [LOCKED / SKIPPED] Failed to remove $($cand.FullPath): $($_.Exception.Message)"
        $failedCount++
    }
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "Execution Summary: Deleted=$deletedCount, Failed/Locked=$failedCount" -ForegroundColor $(if ($failedCount -eq 0) { 'Green' } else { 'Yellow' })
Write-Host "==========================================================" -ForegroundColor Cyan
