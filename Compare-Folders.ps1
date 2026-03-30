# ============================================================
# Compare-Folders.ps1
#
# Compares files between two folders (including subfolders).
# Identifies filenames (excluding extension) present in
# Folder 1 that are NOT present in Folder 2.
# Supports local and network (UNC) paths.
# Runs a second QC scan to verify results.
# ============================================================

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry  = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red    }
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
        default   { Write-Host $logEntry }
    }

    if ($script:LogPath) {
        try {
            Add-Content -Path $script:LogPath -Value $logEntry -ErrorAction Stop
        } catch {
            # Cannot write to log — print to console only, do NOT open anything
            Write-Host "[$timestamp] [WARNING] Unable to write to log file '$($script:LogPath)': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Test-FolderAccessible {
    <#
    .SYNOPSIS
        Attempts a lightweight read of the folder to confirm access before scanning.
    #>
    param([string]$FolderPath)

    try {
        $null = Get-ChildItem -LiteralPath $FolderPath -Force -ErrorAction Stop |
                Select-Object -First 1
        return $true
    } catch [System.UnauthorizedAccessException] {
        Write-Log "ACCESS DENIED to folder: '$FolderPath'. " +
                  "Ensure the account '$($env:USERDOMAIN)\$($env:USERNAME)' has at least Read permission." `
                  -Level ERROR
        return $false
    } catch [System.IO.DirectoryNotFoundException] {
        Write-Log "DIRECTORY NOT FOUND: '$FolderPath'. " +
                  "Verify the path exists and is reachable (check VPN / network share availability)." `
                  -Level ERROR
        return $false
    } catch [System.IO.IOException] {
        Write-Log "I/O ERROR accessing '$FolderPath': $($_.Exception.Message)" -Level ERROR
        return $false
    } catch {
        Write-Log "UNEXPECTED ERROR accessing '$FolderPath': $($_.Exception.GetType().Name) — $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Get-FileBaseNames {
    <#
    .SYNOPSIS
        Recursively collects the base names (no extension) of every file
        under $FolderPath. Returns a case-insensitive HashSet.
    #>
    param(
        [string]$FolderPath,
        [string]$ScanLabel = 'Scan'
    )

    $baseNames  = [System.Collections.Generic.HashSet[string]]::new(
                      [System.StringComparer]::OrdinalIgnoreCase)
    $fileCount  = 0
    $errorCount = 0

    Write-Log "[$ScanLabel] Beginning recursive file enumeration of: '$FolderPath'"

    try {
        # -ErrorVariable captures per-item errors; -ErrorAction SilentlyContinue
        # lets the scan continue past unreadable items.
        $allFiles = Get-ChildItem -LiteralPath $FolderPath `
                                  -Recurse -File -Force `
                                  -ErrorAction SilentlyContinue `
                                  -ErrorVariable itemErrors

        # Report every item-level error without stopping
        foreach ($ie in $itemErrors) {
            $errorCount++
            Write-Log "[$ScanLabel] Could not enumerate item — $($ie.Exception.Message)" -Level WARNING
        }

        foreach ($file in $allFiles) {
            try {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

                if ([string]::IsNullOrWhiteSpace($baseName)) {
                    Write-Log "[$ScanLabel] Skipping file with empty base name: '$($file.FullName)'" -Level WARNING
                    continue
                }

                $null = $baseNames.Add($baseName)
                $fileCount++
            } catch {
                $errorCount++
                Write-Log "[$ScanLabel] Error processing file '$($file.FullName)': " +
                          "$($_.Exception.GetType().Name) — $($_.Exception.Message)" -Level ERROR
            }
        }
    } catch [System.UnauthorizedAccessException] {
        Write-Log "[$ScanLabel] ACCESS DENIED during scan of '$FolderPath': $($_.Exception.Message)" -Level ERROR
    } catch [System.IO.IOException] {
        Write-Log "[$ScanLabel] I/O ERROR during scan of '$FolderPath': $($_.Exception.Message)" -Level ERROR
    } catch {
        Write-Log "[$ScanLabel] CRITICAL ERROR during scan of '$FolderPath': " +
                  "$($_.Exception.GetType().Name) — $($_.Exception.Message)" -Level ERROR
    }

    Write-Log "[$ScanLabel] Enumeration finished. Files processed: $fileCount | Unique base names: $($baseNames.Count) | Errors: $errorCount"
    return $baseNames
}

#endregion

#region Main Script

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "         FOLDER COMPARISON TOOL  —  Missing File Finder     " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Running as : $($env:USERDOMAIN)\$($env:USERNAME)"          -ForegroundColor Yellow
Write-Host "  Date/Time  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"    -ForegroundColor Yellow
Write-Host ""

# ----------------------------------------------------------------
# Step 1 — Collect folder paths from the user
# ----------------------------------------------------------------
Write-Host "--- Step 1: Enter Folder Paths ---" -ForegroundColor Cyan

do {
    $folder1 = (Read-Host "  Folder Path 1 (SOURCE  — files to check)").Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($folder1)) {
        Write-Host "  [ERROR] Folder Path 1 cannot be empty. Please enter a valid path." -ForegroundColor Red
    }
} while ([string]::IsNullOrWhiteSpace($folder1))

do {
    $folder2 = (Read-Host "  Folder Path 2 (COMPARE — files to compare against)").Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($folder2)) {
        Write-Host "  [ERROR] Folder Path 2 cannot be empty. Please enter a valid path." -ForegroundColor Red
    }
} while ([string]::IsNullOrWhiteSpace($folder2))

Write-Host ""

# ----------------------------------------------------------------
# Step 2 — Validate that both paths exist and are directories
# ----------------------------------------------------------------
Write-Host "--- Step 2: Validating Folder Paths ---" -ForegroundColor Cyan

$pathErrors = $false

foreach ($entry in @(
    [pscustomobject]@{ Label = 'Folder 1 (Source) '; Path = $folder1 },
    [pscustomobject]@{ Label = 'Folder 2 (Compare)'; Path = $folder2 }
)) {
    if (-not (Test-Path -LiteralPath $entry.Path -PathType Container)) {
        Write-Host "  [ERROR] $($entry.Label): Path not found or not a directory — '$($entry.Path)'" -ForegroundColor Red
        Write-Host "          Check spelling, network connectivity, and that the share is mounted." -ForegroundColor Red
        $pathErrors = $true
    } else {
        Write-Host "  [OK]    $($entry.Label): '$($entry.Path)'" -ForegroundColor Green
    }
}

if ($pathErrors) {
    Write-Host ""
    Write-Host "  One or more paths could not be validated. Correct the paths and re-run the script." -ForegroundColor Red
    Read-Host "  Press Enter to exit"
    exit 1
}

# ----------------------------------------------------------------
# Step 3 — Set up output directory and log file
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 3: Preparing Output Directory ---" -ForegroundColor Cyan

$outputParentDir = Split-Path -LiteralPath $folder1 -Parent
$outputDir       = Join-Path -Path $outputParentDir -ChildPath "Missing"
$outputFile      = Join-Path -Path $outputDir -ChildPath "Missing.txt"
$script:LogPath  = Join-Path -Path $outputDir -ChildPath "Compare-Folders_Log.txt"

Write-Host "  Output directory : $outputDir"
Write-Host "  Report file      : $outputFile"
Write-Host "  Log file         : $($script:LogPath)"

try {
    if (-not (Test-Path -LiteralPath $outputDir -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDir -Force -ErrorAction Stop | Out-Null
        Write-Host "  [OK] Created output directory." -ForegroundColor Green
    } else {
        Write-Host "  [OK] Output directory already exists." -ForegroundColor Green
    }
} catch [System.UnauthorizedAccessException] {
    Write-Host "  [ERROR] ACCESS DENIED — cannot create output directory '$outputDir'." -ForegroundColor Red
    Write-Host "          Ensure write permission exists one level above Folder 1." -ForegroundColor Red
    Read-Host "  Press Enter to exit"
    exit 1
} catch {
    Write-Host "  [ERROR] Could not create output directory '$outputDir': $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "  Press Enter to exit"
    exit 1
}

# Log file is now available; initialise it
Write-Log "Script started."
Write-Log "Running as      : $($env:USERDOMAIN)\$($env:USERNAME)"
Write-Log "Folder 1 Source : $folder1"
Write-Log "Folder 2 Compare: $folder2"
Write-Log "Output directory: $outputDir"

# ----------------------------------------------------------------
# Step 4 — Permission pre-flight check on both folders
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 4: Checking Read Permissions ---" -ForegroundColor Cyan
Write-Log "Starting permission pre-flight checks..."

$f1Accessible = Test-FolderAccessible -FolderPath $folder1
$f2Accessible = Test-FolderAccessible -FolderPath $folder2

if (-not $f1Accessible) {
    Write-Log "Folder 1 is not readable. Cannot proceed." -Level ERROR
    Read-Host "Press Enter to exit"
    exit 1
}
if (-not $f2Accessible) {
    Write-Log "Folder 2 is not readable. Cannot proceed." -Level ERROR
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "  [OK] Read access confirmed for both folders." -ForegroundColor Green
Write-Log "Permission pre-flight checks passed."

# ----------------------------------------------------------------
# Step 5 — Primary scan
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 5: PRIMARY SCAN ---" -ForegroundColor Cyan
Write-Log "=== PRIMARY SCAN BEGIN ==="

$primaryF1Names = Get-FileBaseNames -FolderPath $folder1 -ScanLabel 'Primary/F1'
$primaryF2Names = Get-FileBaseNames -FolderPath $folder2 -ScanLabel 'Primary/F2'

$primaryMissing = [System.Collections.Generic.List[string]]::new()
foreach ($name in $primaryF1Names) {
    if (-not $primaryF2Names.Contains($name)) {
        $primaryMissing.Add($name)
    }
}

Write-Log "=== PRIMARY SCAN COMPLETE === Missing count: $($primaryMissing.Count)"
Write-Host "  Primary scan — files missing from Folder 2: $($primaryMissing.Count)" -ForegroundColor White

# ----------------------------------------------------------------
# Step 6 — QC (verification) scan
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 6: QC VERIFICATION SCAN ---" -ForegroundColor Cyan
Write-Log "=== QC VERIFICATION SCAN BEGIN ==="

$qcF1Names = Get-FileBaseNames -FolderPath $folder1 -ScanLabel 'QC/F1'
$qcF2Names = Get-FileBaseNames -FolderPath $folder2 -ScanLabel 'QC/F2'

$qcMissing = [System.Collections.Generic.List[string]]::new()
foreach ($name in $qcF1Names) {
    if (-not $qcF2Names.Contains($name)) {
        $qcMissing.Add($name)
    }
}

Write-Log "=== QC VERIFICATION SCAN COMPLETE === Missing count: $($qcMissing.Count)"
Write-Host "  QC scan        — files missing from Folder 2: $($qcMissing.Count)" -ForegroundColor White

# ----------------------------------------------------------------
# Step 7 — Reconcile primary vs QC results
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 7: Reconciling Scan Results ---" -ForegroundColor Cyan
Write-Log "Reconciling primary scan vs QC scan..."

$primarySet = [System.Collections.Generic.HashSet[string]]::new(
                  $primaryMissing, [System.StringComparer]::OrdinalIgnoreCase)
$qcSet      = [System.Collections.Generic.HashSet[string]]::new(
                  $qcMissing, [System.StringComparer]::OrdinalIgnoreCase)

$onlyInPrimary = $primaryMissing | Where-Object { -not $qcSet.Contains($_) }
$onlyInQC      = $qcMissing      | Where-Object { -not $primarySet.Contains($_) }

if ($onlyInPrimary -or $onlyInQC) {
    Write-Log "DISCREPANCY detected between primary and QC scans." -Level WARNING

    if ($onlyInPrimary) {
        Write-Log "  Found in Primary only ($($onlyInPrimary.Count) item(s)): $($onlyInPrimary -join '; ')" -Level WARNING
    }
    if ($onlyInQC) {
        Write-Log "  Found in QC only ($($onlyInQC.Count) item(s)): $($onlyInQC -join '; ')" -Level WARNING
    }

    Write-Log "Resolution: using UNION of both scans to ensure no files are missed." -Level WARNING

    $finalSet = [System.Collections.Generic.HashSet[string]]::new(
                    $primaryMissing, [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($n in $qcMissing) { $null = $finalSet.Add($n) }
    $finalMissing = $finalSet | Sort-Object
} else {
    Write-Log "Scans are consistent — no discrepancies found."
    Write-Host "  [OK] Both scans agree. Results are consistent." -ForegroundColor Green
    $finalMissing = $primaryMissing | Sort-Object
}

# ----------------------------------------------------------------
# Step 8 — Write Missing.txt report
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 8: Writing Report ---" -ForegroundColor Cyan
Write-Log "Writing report to: $outputFile"

try {
    $reportLines = [System.Collections.Generic.List[string]]::new()

    $reportLines.Add("================================================================")
    $reportLines.Add("  MISSING FILES REPORT")
    $reportLines.Add("  Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $reportLines.Add("  Run by    : $($env:USERDOMAIN)\$($env:USERNAME)")
    $reportLines.Add("================================================================")
    $reportLines.Add("")
    $reportLines.Add("  Folder 1 (Source) : $folder1")
    $reportLines.Add("  Folder 2 (Compare): $folder2")
    $reportLines.Add("")
    $reportLines.Add("  Comparison method : Filename only (extension excluded)")
    $reportLines.Add("  Scan method       : Primary scan + QC verification scan")
    $reportLines.Add("")
    $reportLines.Add("  Files in Folder 1 (unique base names) : $($primaryF1Names.Count)")
    $reportLines.Add("  Files in Folder 2 (unique base names) : $($primaryF2Names.Count)")
    $reportLines.Add("  Missing from Folder 2                 : $($finalMissing.Count)")
    $reportLines.Add("")
    $reportLines.Add("================================================================")
    $reportLines.Add("  MISSING FILE NAMES (base name, no extension)")
    $reportLines.Add("================================================================")
    $reportLines.Add("")

    if ($finalMissing.Count -gt 0) {
        foreach ($name in $finalMissing) {
            $reportLines.Add($name)
        }
    } else {
        $reportLines.Add("  *** No missing files detected. ***")
        $reportLines.Add("  All file base names found in Folder 1 are also present in Folder 2.")
    }

    $reportLines.Add("")
    $reportLines.Add("================================================================")
    $reportLines.Add("  END OF REPORT")
    $reportLines.Add("================================================================")

    # Set-Content writes without opening any external viewer
    $reportLines | Set-Content -LiteralPath $outputFile -Encoding UTF8 -ErrorAction Stop

    Write-Log "Report written successfully: $outputFile"
    Write-Host "  [OK] Report saved: $outputFile" -ForegroundColor Green

} catch [System.UnauthorizedAccessException] {
    Write-Log "ACCESS DENIED — cannot write report file '$outputFile': $($_.Exception.Message)" -Level ERROR
} catch [System.IO.IOException] {
    Write-Log "I/O ERROR writing report '$outputFile': $($_.Exception.Message)" -Level ERROR
} catch {
    Write-Log "UNEXPECTED ERROR writing report '$outputFile': $($_.Exception.GetType().Name) — $($_.Exception.Message)" -Level ERROR
}

# ----------------------------------------------------------------
# Summary
# ----------------------------------------------------------------
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ("  Folder 1 unique base names : {0,6}" -f $primaryF1Names.Count)
Write-Host ("  Folder 2 unique base names : {0,6}" -f $primaryF2Names.Count)
Write-Host ("  Missing from Folder 2      : {0,6}" -f $finalMissing.Count)
Write-Host ""
Write-Host "  Report : $outputFile"
Write-Host "  Log    : $($script:LogPath)"
Write-Host "============================================================" -ForegroundColor Cyan

Write-Log "Script completed."
Write-Host ""
Read-Host "Press Enter to exit"

#endregion
