# ============================================================
# Compare-Folders.ps1
#
# Compares files between two folders (including subfolders).
# Identifies filenames (excluding extension) present in
# Folder 1 that are NOT present in Folder 2.
# Copies those files into Folder1\Missing\ as a flat list
# (all files placed directly in Missing\, no subfolders).
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
            Write-Host "[$timestamp] [WARNING] Unable to write to log file '$($script:LogPath)': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Test-FolderAccessible {
    param([string]$FolderPath)

    try {
        $null = Get-ChildItem -LiteralPath $FolderPath -Force -ErrorAction Stop |
                Select-Object -First 1
        return $true
    } catch [System.UnauthorizedAccessException] {
        Write-Log "ACCESS DENIED to folder: '$FolderPath'. Ensure the account '$($env:USERDOMAIN)\$($env:USERNAME)' has at least Read permission." -Level ERROR
        return $false
    } catch [System.IO.DirectoryNotFoundException] {
        Write-Log "DIRECTORY NOT FOUND: '$FolderPath'. Verify the path exists and is reachable (check VPN / network share availability)." -Level ERROR
        return $false
    } catch [System.IO.IOException] {
        Write-Log "I/O ERROR accessing '$FolderPath': $($_.Exception.Message)" -Level ERROR
        return $false
    } catch {
        Write-Log "UNEXPECTED ERROR accessing '$FolderPath': $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

function Get-Folder1Files {
    <#
    .SYNOPSIS
        Recursively collects all file objects from Folder 1, skipping the
        Missing output subdirectory so re-runs do not pick up copied files.
        Returns a hashtable: { BaseNames = HashSet; Files = List[FileInfo] }
    #>
    param(
        [string]$FolderPath,
        [string]$ExcludePath,
        [string]$ScanLabel = 'Scan'
    )

    $baseNames  = [System.Collections.Generic.HashSet[string]]::new(
                      [System.StringComparer]::OrdinalIgnoreCase)
    $fileList   = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    $fileCount  = 0
    $errorCount = 0

    # Normalise the exclude path so string comparison is reliable
    $excludeNorm = $ExcludePath.TrimEnd('\') + '\'

    Write-Log "[$ScanLabel] Scanning Folder 1: '$FolderPath'"
    if ($ExcludePath) {
        Write-Log "[$ScanLabel] Excluding subdirectory from scan: '$ExcludePath'"
    }

    try {
        $allFiles = Get-ChildItem -LiteralPath $FolderPath `
                                  -Recurse -File -Force `
                                  -ErrorAction SilentlyContinue `
                                  -ErrorVariable itemErrors

        foreach ($ie in $itemErrors) {
            $errorCount++
            Write-Log "[$ScanLabel] Could not enumerate item - $($ie.Exception.Message)" -Level WARNING
        }

        foreach ($file in $allFiles) {
            try {
                # Skip anything inside the Missing output directory
                if ($ExcludePath -and $file.FullName.StartsWith($excludeNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

                if ([string]::IsNullOrWhiteSpace($baseName)) {
                    Write-Log "[$ScanLabel] Skipping file with empty base name: '$($file.FullName)'" -Level WARNING
                    continue
                }

                $null = $baseNames.Add($baseName)
                $fileList.Add($file)
                $fileCount++
            } catch {
                $errorCount++
                Write-Log "[$ScanLabel] Error processing file '$($file.FullName)': $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level ERROR
            }
        }
    } catch [System.UnauthorizedAccessException] {
        Write-Log "[$ScanLabel] ACCESS DENIED during scan of '$FolderPath': $($_.Exception.Message)" -Level ERROR
    } catch [System.IO.IOException] {
        Write-Log "[$ScanLabel] I/O ERROR during scan of '$FolderPath': $($_.Exception.Message)" -Level ERROR
    } catch {
        Write-Log "[$ScanLabel] CRITICAL ERROR during scan of '$FolderPath': $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level ERROR
    }

    Write-Log "[$ScanLabel] Done. Files scanned: $fileCount | Unique base names: $($baseNames.Count) | Errors: $errorCount"
    return @{ BaseNames = $baseNames; Files = $fileList }
}

function Get-Folder2BaseNames {
    <#
    .SYNOPSIS
        Recursively collects base names (no extension) from Folder 2.
        Returns a case-insensitive HashSet.
    #>
    param(
        [string]$FolderPath,
        [string]$ScanLabel = 'Scan'
    )

    $baseNames  = [System.Collections.Generic.HashSet[string]]::new(
                      [System.StringComparer]::OrdinalIgnoreCase)
    $fileCount  = 0
    $errorCount = 0

    Write-Log "[$ScanLabel] Scanning Folder 2: '$FolderPath'"

    try {
        $allFiles = Get-ChildItem -LiteralPath $FolderPath `
                                  -Recurse -File -Force `
                                  -ErrorAction SilentlyContinue `
                                  -ErrorVariable itemErrors

        foreach ($ie in $itemErrors) {
            $errorCount++
            Write-Log "[$ScanLabel] Could not enumerate item - $($ie.Exception.Message)" -Level WARNING
        }

        foreach ($file in $allFiles) {
            try {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                if ([string]::IsNullOrWhiteSpace($baseName)) { continue }
                $null = $baseNames.Add($baseName)
                $fileCount++
            } catch {
                $errorCount++
                Write-Log "[$ScanLabel] Error processing file '$($file.FullName)': $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level ERROR
            }
        }
    } catch [System.UnauthorizedAccessException] {
        Write-Log "[$ScanLabel] ACCESS DENIED during scan of '$FolderPath': $($_.Exception.Message)" -Level ERROR
    } catch [System.IO.IOException] {
        Write-Log "[$ScanLabel] I/O ERROR during scan of '$FolderPath': $($_.Exception.Message)" -Level ERROR
    } catch {
        Write-Log "[$ScanLabel] CRITICAL ERROR during scan of '$FolderPath': $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level ERROR
    }

    Write-Log "[$ScanLabel] Done. Files scanned: $fileCount | Unique base names: $($baseNames.Count) | Errors: $errorCount"
    return $baseNames
}

function Copy-MissingFiles {
    <#
    .SYNOPSIS
        Copies files whose base name is not present in $MissingBaseNames
        from Folder 1 into the Missing output directory as a flat list.
        All files are placed directly in Missing\ with no subfolders.
        Returns a hashtable: { CopiedCount; SkippedCount; ErrorCount; CopiedFiles }
    #>
    param(
        [string]$Folder1Path,
        [string]$OutputDir,
        [System.Collections.Generic.List[System.IO.FileInfo]]$Folder1Files,
        [System.Collections.Generic.HashSet[string]]$Folder2BaseNames,
        [string]$ScanLabel = 'Copy'
    )

    $copiedCount  = 0
    $skippedCount = 0
    $errorCount   = 0
    $copiedFiles  = [System.Collections.Generic.List[string]]::new()

    $folder1Norm = $Folder1Path.TrimEnd('\') + '\'

    Write-Log "[$ScanLabel] Beginning file copy to: '$OutputDir'"

    foreach ($file in $Folder1Files) {
        try {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            if ($Folder2BaseNames.Contains($baseName)) {
                $skippedCount++
                continue
            }

            # Flat copy - all files go directly into the Missing folder (no subfolders)
            $destination = Join-Path -Path $OutputDir -ChildPath $file.Name

            # If a file with the same filename already exists in the flat destination,
            # append the base name with a counter to avoid silent overwrites
            if (Test-Path -LiteralPath $destination -PathType Leaf) {
                $basePart = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $ext      = [System.IO.Path]::GetExtension($file.Name)
                $counter  = 1
                do {
                    $destination = Join-Path -Path $OutputDir -ChildPath "$($basePart)_$counter$ext"
                    $counter++
                } while (Test-Path -LiteralPath $destination -PathType Leaf)
                Write-Log "[$ScanLabel] Filename conflict - renamed to: '$destination'" -Level WARNING
            }

            # Copy the file - do not open or alter content
            Copy-Item -LiteralPath $file.FullName -Destination $destination -Force -ErrorAction Stop

            $copiedFiles.Add($file.FullName)
            $copiedCount++
            Write-Log "[$ScanLabel] Copied: '$($file.FullName)' -> '$destination'"

        } catch [System.UnauthorizedAccessException] {
            $errorCount++
            Write-Log "[$ScanLabel] ACCESS DENIED copying '$($file.FullName)': $($_.Exception.Message)" -Level ERROR
        } catch [System.IO.IOException] {
            $errorCount++
            Write-Log "[$ScanLabel] I/O ERROR copying '$($file.FullName)': $($_.Exception.Message)" -Level ERROR
        } catch {
            $errorCount++
            Write-Log "[$ScanLabel] UNEXPECTED ERROR copying '$($file.FullName)': $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-Log "[$ScanLabel] Copy complete. Copied: $copiedCount | Already in F2 (skipped): $skippedCount | Errors: $errorCount"
    return @{ CopiedCount = $copiedCount; SkippedCount = $skippedCount; ErrorCount = $errorCount; CopiedFiles = $copiedFiles }
}

#endregion

#region Main Script

Clear-Host
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "        FOLDER COMPARISON TOOL - Missing File Finder        " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Running as : $($env:USERDOMAIN)\$($env:USERNAME)"          -ForegroundColor Yellow
Write-Host "  Date/Time  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"    -ForegroundColor Yellow
Write-Host ""

# ----------------------------------------------------------------
# Step 1 - Collect folder paths from the user
# ----------------------------------------------------------------
Write-Host "--- Step 1: Enter Folder Paths ---" -ForegroundColor Cyan

do {
    $folder1 = (Read-Host "  Folder Path 1 (SOURCE  - files to check)").Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($folder1)) {
        Write-Host "  [ERROR] Folder Path 1 cannot be empty. Please enter a valid path." -ForegroundColor Red
    }
} while ([string]::IsNullOrWhiteSpace($folder1))

do {
    $folder2 = (Read-Host "  Folder Path 2 (COMPARE - files to compare against)").Trim().TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($folder2)) {
        Write-Host "  [ERROR] Folder Path 2 cannot be empty. Please enter a valid path." -ForegroundColor Red
    }
} while ([string]::IsNullOrWhiteSpace($folder2))

Write-Host ""

# ----------------------------------------------------------------
# Step 2 - Validate that both paths exist and are directories
# ----------------------------------------------------------------
Write-Host "--- Step 2: Validating Folder Paths ---" -ForegroundColor Cyan

$pathErrors = $false

foreach ($entry in @(
    [pscustomobject]@{ Label = 'Folder 1 (Source) '; Path = $folder1 },
    [pscustomobject]@{ Label = 'Folder 2 (Compare)'; Path = $folder2 }
)) {
    if (-not (Test-Path -LiteralPath $entry.Path -PathType Container)) {
        Write-Host "  [ERROR] $($entry.Label): Path not found or not a directory - '$($entry.Path)'" -ForegroundColor Red
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
# Step 3 - Set up output directory and log file (inside Folder 1)
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 3: Preparing Output Directory ---" -ForegroundColor Cyan

$outputDir      = Join-Path -Path $folder1 -ChildPath "Missing"
$outputFile     = Join-Path -Path $outputDir -ChildPath "Missing.txt"
$script:LogPath = Join-Path -Path $outputDir -ChildPath "Compare-Folders_Log.txt"

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
    Write-Host "  [ERROR] ACCESS DENIED - cannot create output directory '$outputDir'." -ForegroundColor Red
    Write-Host "          Ensure write permission exists inside Folder 1." -ForegroundColor Red
    Read-Host "  Press Enter to exit"
    exit 1
} catch {
    Write-Host "  [ERROR] Could not create output directory '$outputDir': $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "  Press Enter to exit"
    exit 1
}

Write-Log "Script started."
Write-Log "Running as      : $($env:USERDOMAIN)\$($env:USERNAME)"
Write-Log "Folder 1 Source : $folder1"
Write-Log "Folder 2 Compare: $folder2"
Write-Log "Output directory: $outputDir  (inside Folder 1 - excluded from scans)"

# ----------------------------------------------------------------
# Step 4 - Permission pre-flight check on both folders
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
# Step 5 - Primary scan
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 5: PRIMARY SCAN ---" -ForegroundColor Cyan
Write-Log "=== PRIMARY SCAN BEGIN ==="

$primaryF1Result = Get-Folder1Files  -FolderPath $folder1 -ExcludePath $outputDir -ScanLabel 'Primary/F1'
$primaryF2Names  = Get-Folder2BaseNames -FolderPath $folder2 -ScanLabel 'Primary/F2'

$primaryF1Names = $primaryF1Result.BaseNames
$primaryF1Files = $primaryF1Result.Files

# Identify missing base names
$primaryMissingNames = [System.Collections.Generic.List[string]]::new()
foreach ($name in $primaryF1Names) {
    if (-not $primaryF2Names.Contains($name)) {
        $primaryMissingNames.Add($name)
    }
}

Write-Log "=== PRIMARY SCAN COMPLETE === Missing base names: $($primaryMissingNames.Count)"
Write-Host "  Primary scan - files missing from Folder 2: $($primaryMissingNames.Count)" -ForegroundColor White

# ----------------------------------------------------------------
# Step 6 - QC (verification) scan
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 6: QC VERIFICATION SCAN ---" -ForegroundColor Cyan
Write-Log "=== QC VERIFICATION SCAN BEGIN ==="

$qcF1Result = Get-Folder1Files  -FolderPath $folder1 -ExcludePath $outputDir -ScanLabel 'QC/F1'
$qcF2Names  = Get-Folder2BaseNames -FolderPath $folder2 -ScanLabel 'QC/F2'

$qcF1Names = $qcF1Result.BaseNames

$qcMissingNames = [System.Collections.Generic.List[string]]::new()
foreach ($name in $qcF1Names) {
    if (-not $qcF2Names.Contains($name)) {
        $qcMissingNames.Add($name)
    }
}

Write-Log "=== QC VERIFICATION SCAN COMPLETE === Missing base names: $($qcMissingNames.Count)"
Write-Host "  QC scan        - files missing from Folder 2: $($qcMissingNames.Count)" -ForegroundColor White

# ----------------------------------------------------------------
# Step 7 - Reconcile primary vs QC results
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 7: Reconciling Scan Results ---" -ForegroundColor Cyan
Write-Log "Reconciling primary scan vs QC scan..."

$primarySet = [System.Collections.Generic.HashSet[string]]::new(
                  $primaryMissingNames, [System.StringComparer]::OrdinalIgnoreCase)
$qcSet      = [System.Collections.Generic.HashSet[string]]::new(
                  $qcMissingNames, [System.StringComparer]::OrdinalIgnoreCase)

$onlyInPrimary = $primaryMissingNames | Where-Object { -not $qcSet.Contains($_) }
$onlyInQC      = $qcMissingNames      | Where-Object { -not $primarySet.Contains($_) }

# Build final set of missing base names
if ($onlyInPrimary -or $onlyInQC) {
    Write-Log "DISCREPANCY detected between primary and QC scans." -Level WARNING

    if ($onlyInPrimary) {
        Write-Log "  Found in Primary only ($($onlyInPrimary.Count) item(s)): $($onlyInPrimary -join '; ')" -Level WARNING
    }
    if ($onlyInQC) {
        Write-Log "  Found in QC only ($($onlyInQC.Count) item(s)): $($onlyInQC -join '; ')" -Level WARNING
    }

    Write-Log "Resolution: using UNION of both scans to ensure no files are missed." -Level WARNING

    $finalMissingNamesSet = [System.Collections.Generic.HashSet[string]]::new(
                                $primaryMissingNames, [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($n in $qcMissingNames) { $null = $finalMissingNamesSet.Add($n) }
} else {
    Write-Log "Scans are consistent - no discrepancies found."
    Write-Host "  [OK] Both scans agree. Results are consistent." -ForegroundColor Green
    $finalMissingNamesSet = $primarySet
}

$finalMissingNamesSorted = $finalMissingNamesSet | Sort-Object

# ----------------------------------------------------------------
# Step 8 - Copy missing files into Folder1\Missing\
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 8: Copying Missing Files ---" -ForegroundColor Cyan
Write-Log "=== FILE COPY BEGIN ==="

# Use the primary scan's file list (QC confirmed the names match)
$copyResult = Copy-MissingFiles -Folder1Path $folder1 `
                                -OutputDir $outputDir `
                                -Folder1Files $primaryF1Files `
                                -Folder2BaseNames $primaryF2Names `
                                -ScanLabel 'Copy'

Write-Log "=== FILE COPY COMPLETE ==="

if ($copyResult.ErrorCount -gt 0) {
    Write-Host "  [WARNING] $($copyResult.ErrorCount) file(s) could not be copied. See log for details." -ForegroundColor Yellow
} else {
    Write-Host "  [OK] All missing files copied successfully." -ForegroundColor Green
}

# ----------------------------------------------------------------
# Step 9 - QC verify copies actually landed on disk
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 9: QC Verifying Copied Files ---" -ForegroundColor Cyan
Write-Log "=== QC COPY VERIFICATION BEGIN ==="

$qcVerifyPassed  = 0
$qcVerifyFailed  = 0

foreach ($srcPath in $copyResult.CopiedFiles) {
    try {
        $srcFileName = [System.IO.Path]::GetFileName($srcPath)
        # Check for the flat-copied file in the Missing folder
        $destMatches = Get-ChildItem -LiteralPath $outputDir -File -Filter $srcFileName -ErrorAction SilentlyContinue

        if ($destMatches) {
            $qcVerifyPassed++
        } else {
            $qcVerifyFailed++
            Write-Log "QC VERIFY FAILED - '$srcFileName' not found in Missing folder (source: '$srcPath')" -Level WARNING
        }
    } catch {
        $qcVerifyFailed++
        Write-Log "QC VERIFY ERROR for '$srcPath': $($_.Exception.Message)" -Level ERROR
    }
}

Write-Log "=== QC COPY VERIFICATION COMPLETE === Passed: $qcVerifyPassed | Failed: $qcVerifyFailed"

if ($qcVerifyFailed -gt 0) {
    Write-Host "  [WARNING] QC verification found $qcVerifyFailed file(s) not confirmed on disk. Check the log." -ForegroundColor Yellow
} else {
    Write-Host "  [OK] QC verification passed. All $qcVerifyPassed copied file(s) confirmed on disk." -ForegroundColor Green
}

# ----------------------------------------------------------------
# Step 10 - Write Missing.txt report
# ----------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 10: Writing Report ---" -ForegroundColor Cyan
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
    $reportLines.Add("  Copied files to   : $outputDir")
    $reportLines.Add("")
    $reportLines.Add("  Comparison method : Filename only (extension excluded)")
    $reportLines.Add("  Scan method       : Primary scan + QC verification scan")
    $reportLines.Add("")
    $reportLines.Add("  Files in Folder 1 (unique base names) : $($primaryF1Names.Count)")
    $reportLines.Add("  Files in Folder 2 (unique base names) : $($primaryF2Names.Count)")
    $reportLines.Add("  Missing from Folder 2                 : $($finalMissingNamesSet.Count)")
    $reportLines.Add("  Files copied to Missing folder        : $($copyResult.CopiedCount)")
    $reportLines.Add("  Copy errors                           : $($copyResult.ErrorCount)")
    $reportLines.Add("  QC verify passed                      : $qcVerifyPassed")
    $reportLines.Add("  QC verify failed                      : $qcVerifyFailed")
    $reportLines.Add("")
    $reportLines.Add("================================================================")
    $reportLines.Add("  MISSING FILE NAMES (base name, no extension)")
    $reportLines.Add("================================================================")
    $reportLines.Add("")

    if ($finalMissingNamesSorted.Count -gt 0) {
        foreach ($name in $finalMissingNamesSorted) {
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

    $reportLines | Set-Content -LiteralPath $outputFile -Encoding UTF8 -ErrorAction Stop

    Write-Log "Report written successfully: $outputFile"
    Write-Host "  [OK] Report saved: $outputFile" -ForegroundColor Green

} catch [System.UnauthorizedAccessException] {
    Write-Log "ACCESS DENIED - cannot write report file '$outputFile': $($_.Exception.Message)" -Level ERROR
} catch [System.IO.IOException] {
    Write-Log "I/O ERROR writing report '$outputFile': $($_.Exception.Message)" -Level ERROR
} catch {
    Write-Log "UNEXPECTED ERROR writing report '$outputFile': $($_.Exception.GetType().Name) - $($_.Exception.Message)" -Level ERROR
}

# ----------------------------------------------------------------
# Summary
# ----------------------------------------------------------------
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ("  Folder 1 unique base names    : {0,6}" -f $primaryF1Names.Count)
Write-Host ("  Folder 2 unique base names    : {0,6}" -f $primaryF2Names.Count)
Write-Host ("  Missing from Folder 2         : {0,6}" -f $finalMissingNamesSet.Count)
Write-Host ("  Files copied to Missing\      : {0,6}" -f $copyResult.CopiedCount)
Write-Host ("  Copy errors                   : {0,6}" -f $copyResult.ErrorCount)
Write-Host ("  QC copy verification passed   : {0,6}" -f $qcVerifyPassed)
Write-Host ("  QC copy verification failed   : {0,6}" -f $qcVerifyFailed)
Write-Host ""
Write-Host "  Missing folder : $outputDir"
Write-Host "  Report         : $outputFile"
Write-Host "  Log            : $($script:LogPath)"
Write-Host "============================================================" -ForegroundColor Cyan

Write-Log "Script completed."
Write-Host ""
Read-Host "Press Enter to exit"

#endregion
