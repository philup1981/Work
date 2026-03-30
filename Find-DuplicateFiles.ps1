# =============================================================================
# Find-DuplicateFiles.ps1
# Compares files in two folders (including subfolders) and identifies
# duplicates where the file name (excluding extension) exists in both.
# Supports local and network (UNC) paths.
# Runs under the current user account - no elevation required.
# =============================================================================

#region ---- Helper Functions -------------------------------------------------

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped message to the console and appends it to the
        shared error/event log list.  Never opens external windows.
    #>
    param (
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','QC')]
        [string]$Level = 'INFO',
        [System.Collections.Generic.List[string]]$Log
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry     = "[$timestamp][$Level] $Message"

    switch ($Level) {
        'INFO'    { Write-Host   $entry -ForegroundColor Cyan    }
        'WARNING' { Write-Host   $entry -ForegroundColor Yellow  }
        'ERROR'   { Write-Host   $entry -ForegroundColor Red     }
        'QC'      { Write-Host   $entry -ForegroundColor Magenta }
    }

    if ($null -ne $Log) {
        $Log.Add($entry)
    }
}

# --------------------------------------------------------------------------- #

function Test-FolderReadAccess {
    <#
    .SYNOPSIS
        Attempts a lightweight read of the folder to confirm the current user
        has at least list-directory access.  Returns $true / $false.
    #>
    param ([string]$FolderPath)

    try {
        $null = [System.IO.Directory]::GetFileSystemEntries($FolderPath)
        return $true
    } catch [UnauthorizedAccessException] {
        return $false
    } catch {
        return $false
    }
}

# --------------------------------------------------------------------------- #

function Get-FilesRecursive {
    <#
    .SYNOPSIS
        Recursively enumerates all files under $FolderPath.
        Access-denied and unreadable paths are logged and skipped;
        no external windows are opened.
    .OUTPUTS
        Array of [PSCustomObject] with FullPath, FileName, BaseName, Extension.
    #>
    param (
        [string]$FolderPath,
        [string]$ScanLabel,
        [System.Collections.Generic.List[string]]$ErrorLog
    )

    $results = [System.Collections.Generic.List[object]]::new()

    # Use a manual stack-based DFS so we can catch per-directory errors
    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($FolderPath)

    while ($queue.Count -gt 0) {
        $currentDir = $queue.Dequeue()

        # --- enumerate sub-directories ---
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($currentDir)
            foreach ($sub in $subDirs) { $queue.Enqueue($sub) }
        } catch [UnauthorizedAccessException] {
            $msg = "[$ScanLabel] ACCESS DENIED reading sub-directories of '$currentDir'. Skipping this path. Error: $($_.Exception.Message)"
            Write-Log -Message $msg -Level 'ERROR' -Log $ErrorLog
            continue
        } catch {
            $msg = "[$ScanLabel] Unexpected error enumerating sub-directories of '$currentDir'. Skipping. Error: $($_.Exception.Message)"
            Write-Log -Message $msg -Level 'ERROR' -Log $ErrorLog
            continue
        }

        # --- enumerate files in this directory ---
        try {
            $filesInDir = [System.IO.Directory]::GetFiles($currentDir)
        } catch [UnauthorizedAccessException] {
            $msg = "[$ScanLabel] ACCESS DENIED reading files in '$currentDir'. Skipping this directory. Error: $($_.Exception.Message)"
            Write-Log -Message $msg -Level 'ERROR' -Log $ErrorLog
            continue
        } catch {
            $msg = "[$ScanLabel] Unexpected error reading files in '$currentDir'. Skipping. Error: $($_.Exception.Message)"
            Write-Log -Message $msg -Level 'ERROR' -Log $ErrorLog
            continue
        }

        foreach ($filePath in $filesInDir) {
            try {
                $fi = [System.IO.FileInfo]::new($filePath)
                $results.Add([PSCustomObject]@{
                    FullPath  = $fi.FullName
                    FileName  = $fi.Name
                    BaseName  = $fi.BaseName          # name without extension
                    Extension = $fi.Extension
                })
            } catch {
                $msg = "[$ScanLabel] Unable to read file info for '$filePath'. Skipping. Error: $($_.Exception.Message)"
                Write-Log -Message $msg -Level 'ERROR' -Log $ErrorLog
            }
        }
    }

    return ,$results   # comma forces array return even when empty
}

# --------------------------------------------------------------------------- #

function Find-Duplicates {
    <#
    .SYNOPSIS
        Compares two lists of file objects and returns the set of base names
        (lower-cased, sorted) that appear in both lists.
    #>
    param (
        [System.Collections.Generic.List[object]]$FileList1,
        [System.Collections.Generic.List[object]]$FileList2
    )

    $set2 = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($f in $FileList2) { $null = $set2.Add($f.BaseName) }

    $duplicates = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($f in $FileList1) {
        if ($set2.Contains($f.BaseName)) {
            $null = $duplicates.Add($f.BaseName.ToLower())
        }
    }

    # Always return a string array (never $null) so callers can safely
    # pass the result to HashSet constructors and Count checks.
    [string[]]$sorted = @($duplicates | Sort-Object)
    return ,$sorted
}

#endregion ---- Helper Functions ----------------------------------------------


#region ---- Main Script ------------------------------------------------------

Clear-Host
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host '   DUPLICATE FILE FINDER' -ForegroundColor Cyan
Write-Host "   Running as: $env:USERDOMAIN\$env:USERNAME" -ForegroundColor Cyan
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host ''

# ---------- Shared error / event log --------------------------------------- #
$sharedLog = [System.Collections.Generic.List[string]]::new()

# ---------- Collect and validate Folder Path 1 ----------------------------- #
Write-Host 'Enter the two folder paths to compare.' -ForegroundColor White
Write-Host 'UNC paths (e.g. \\server\share\folder) are supported.' -ForegroundColor Gray
Write-Host ''

do {
    $Folder1 = (Read-Host 'Folder Path 1').Trim().TrimEnd('\')

    if ([string]::IsNullOrWhiteSpace($Folder1)) {
        Write-Host 'ERROR: Folder Path 1 cannot be empty. Please try again.' -ForegroundColor Red
        continue
    }

    if (-not [System.IO.Directory]::Exists($Folder1)) {
        Write-Host "ERROR: Folder Path 1 does not exist or is unreachable: '$Folder1'" -ForegroundColor Red
        Write-Host '       Check the path spelling and network connectivity, then try again.' -ForegroundColor Yellow
        $Folder1 = ''
        continue
    }

    if (-not (Test-FolderReadAccess -FolderPath $Folder1)) {
        Write-Host "ERROR: The current user ($env:USERDOMAIN\$env:USERNAME) does not have read access to Folder Path 1: '$Folder1'" -ForegroundColor Red
        Write-Host '       Resolve the permission issue and re-run the script.' -ForegroundColor Yellow
        $Folder1 = ''
        continue
    }

    Write-Host "OK  - Folder Path 1 verified: $Folder1" -ForegroundColor Green

} while ([string]::IsNullOrWhiteSpace($Folder1))

Write-Host ''

# ---------- Collect and validate Folder Path 2 ----------------------------- #
do {
    $Folder2 = (Read-Host 'Folder Path 2').Trim().TrimEnd('\')

    if ([string]::IsNullOrWhiteSpace($Folder2)) {
        Write-Host 'ERROR: Folder Path 2 cannot be empty. Please try again.' -ForegroundColor Red
        continue
    }

    if (-not [System.IO.Directory]::Exists($Folder2)) {
        Write-Host "ERROR: Folder Path 2 does not exist or is unreachable: '$Folder2'" -ForegroundColor Red
        Write-Host '       Check the path spelling and network connectivity, then try again.' -ForegroundColor Yellow
        $Folder2 = ''
        continue
    }

    if (-not (Test-FolderReadAccess -FolderPath $Folder2)) {
        Write-Host "ERROR: The current user ($env:USERDOMAIN\$env:USERNAME) does not have read access to Folder Path 2: '$Folder2'" -ForegroundColor Red
        Write-Host '       Resolve the permission issue and re-run the script.' -ForegroundColor Yellow
        $Folder2 = ''
        continue
    }

    Write-Host "OK  - Folder Path 2 verified: $Folder2" -ForegroundColor Green

} while ([string]::IsNullOrWhiteSpace($Folder2))

Write-Host ''

# ---------- Prepare output directory --------------------------------------- #
# Output goes to a "Duplicates" folder one level above Folder Path 1
$outputParent   = Split-Path -Path $Folder1 -Parent
$duplicatesDir  = Join-Path -Path $outputParent -ChildPath 'Duplicates'
$duplicatesTxt  = Join-Path -Path $duplicatesDir -ChildPath 'Duplicates.txt'
$errorLogTxt    = Join-Path -Path $duplicatesDir -ChildPath 'ErrorLog.txt'

Write-Host "Output directory : $duplicatesDir" -ForegroundColor White

try {
    if (-not (Test-Path -Path $duplicatesDir -PathType Container)) {
        $null = New-Item -Path $duplicatesDir -ItemType Directory -Force -ErrorAction Stop
        Write-Host "Created output directory." -ForegroundColor Green
    } else {
        Write-Host "Output directory already exists - files will be overwritten." -ForegroundColor Yellow
    }
} catch {
    Write-Host "CRITICAL ERROR: Cannot create output directory '$duplicatesDir'." -ForegroundColor Red
    Write-Host "  Detail: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Ensure you have write access to: $outputParent" -ForegroundColor Yellow
    exit 1
}

Write-Host ''

# =========================================================================== #
#  PRIMARY SCAN                                                                #
# =========================================================================== #
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host '   PRIMARY SCAN' -ForegroundColor Cyan
Write-Host ('=' * 60) -ForegroundColor Cyan

Write-Host "Scanning Folder 1 (primary): $Folder1" -ForegroundColor White
$primaryFiles1 = Get-FilesRecursive -FolderPath $Folder1 -ScanLabel 'PRIMARY-F1' -ErrorLog $sharedLog
Write-Host "  -> $($primaryFiles1.Count) file(s) found." -ForegroundColor Green

Write-Host "Scanning Folder 2 (primary): $Folder2" -ForegroundColor White
$primaryFiles2 = Get-FilesRecursive -FolderPath $Folder2 -ScanLabel 'PRIMARY-F2' -ErrorLog $sharedLog
Write-Host "  -> $($primaryFiles2.Count) file(s) found." -ForegroundColor Green

Write-Host ''
Write-Host 'Comparing file names (primary scan)...' -ForegroundColor White
$primaryDuplicates = Find-Duplicates -FileList1 $primaryFiles1 -FileList2 $primaryFiles2
Write-Host "  -> Primary scan identified $($primaryDuplicates.Count) duplicate base name(s)." -ForegroundColor $(if ($primaryDuplicates.Count -gt 0) {'Yellow'} else {'Green'})

Write-Host ''

# =========================================================================== #
#  QC VERIFICATION SCAN                                                        #
# =========================================================================== #
Write-Host ('=' * 60) -ForegroundColor Magenta
Write-Host '   QC VERIFICATION SCAN' -ForegroundColor Magenta
Write-Host ('=' * 60) -ForegroundColor Magenta

Write-Host "Re-scanning Folder 1 (QC): $Folder1" -ForegroundColor White
$qcFiles1 = Get-FilesRecursive -FolderPath $Folder1 -ScanLabel 'QC-F1' -ErrorLog $sharedLog
Write-Host "  -> $($qcFiles1.Count) file(s) found." -ForegroundColor Green

Write-Host "Re-scanning Folder 2 (QC): $Folder2" -ForegroundColor White
$qcFiles2 = Get-FilesRecursive -FolderPath $Folder2 -ScanLabel 'QC-F2' -ErrorLog $sharedLog
Write-Host "  -> $($qcFiles2.Count) file(s) found." -ForegroundColor Green

Write-Host ''
Write-Host 'Comparing file names (QC scan)...' -ForegroundColor White
$qcDuplicates = Find-Duplicates -FileList1 $qcFiles1 -FileList2 $qcFiles2
Write-Host "  -> QC scan identified $($qcDuplicates.Count) duplicate base name(s)." -ForegroundColor $(if ($qcDuplicates.Count -gt 0) {'Yellow'} else {'Green'})

Write-Host ''

# =========================================================================== #
#  CROSS-CHECK PRIMARY vs QC                                                   #
# =========================================================================== #
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host '   CROSS-CHECK: PRIMARY vs QC' -ForegroundColor Cyan
Write-Host ('=' * 60) -ForegroundColor Cyan

$primarySet = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@($primaryDuplicates), [System.StringComparer]::OrdinalIgnoreCase
)
$qcSet = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@($qcDuplicates), [System.StringComparer]::OrdinalIgnoreCase
)

# Items in primary but not QC
$onlyInPrimary = $primaryDuplicates | Where-Object { -not $qcSet.Contains($_) }
# Items in QC but not primary
$onlyInQC      = $qcDuplicates      | Where-Object { -not $primarySet.Contains($_) }

$discrepancyCount = ($onlyInPrimary.Count + $onlyInQC.Count)

if ($discrepancyCount -eq 0) {
    Write-Host 'QC PASS: Both scans returned identical results.' -ForegroundColor Green
    $finalDuplicates = $primaryDuplicates   # either list is fine; they match
} else {
    Write-Host "QC WARNING: $discrepancyCount discrepanc(ies) detected between scans." -ForegroundColor Red

    foreach ($item in $onlyInPrimary) {
        $msg = "Discrepancy - found in PRIMARY scan only: '$item'"
        Write-Log -Message $msg -Level 'QC' -Log $sharedLog
    }
    foreach ($item in $onlyInQC) {
        $msg = "Discrepancy - found in QC scan only: '$item'"
        Write-Log -Message $msg -Level 'QC' -Log $sharedLog
    }

    # Conservative approach: take the UNION so nothing is missed
    $unionSet = [System.Collections.Generic.HashSet[string]]::new(
        $primarySet, [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($item in $qcDuplicates) { $null = $unionSet.Add($item) }
    $finalDuplicates = ($unionSet | Sort-Object)

    Write-Host "Using union of both scans ($($finalDuplicates.Count) total) as the final result." -ForegroundColor Yellow
}

# File-count consistency check
if ($primaryFiles1.Count -ne $qcFiles1.Count) {
    $msg = "File count mismatch for Folder 1 between scans: Primary=$($primaryFiles1.Count), QC=$($qcFiles1.Count). The folder contents may have changed during the scan."
    Write-Log -Message $msg -Level 'QC' -Log $sharedLog
}
if ($primaryFiles2.Count -ne $qcFiles2.Count) {
    $msg = "File count mismatch for Folder 2 between scans: Primary=$($primaryFiles2.Count), QC=$($qcFiles2.Count). The folder contents may have changed during the scan."
    Write-Log -Message $msg -Level 'QC' -Log $sharedLog
}

Write-Host ''

# =========================================================================== #
#  BUILD DUPLICATES.TXT REPORT                                                 #
# =========================================================================== #
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host '   WRITING REPORT' -ForegroundColor Cyan
Write-Host ('=' * 60) -ForegroundColor Cyan

$report = [System.Text.StringBuilder]::new()

$null = $report.AppendLine('=' * 60)
$null = $report.AppendLine('DUPLICATE FILE FINDER - RESULTS REPORT')
$null = $report.AppendLine('=' * 60)
$null = $report.AppendLine("Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$null = $report.AppendLine("Run by     : $env:USERDOMAIN\$env:USERNAME")
$null = $report.AppendLine("Computer   : $env:COMPUTERNAME")
$null = $report.AppendLine('')
$null = $report.AppendLine("Folder 1   : $Folder1")
$null = $report.AppendLine("Folder 2   : $Folder2")
$null = $report.AppendLine('')
$null = $report.AppendLine('--- Scan Statistics ----------------------------------------')
$null = $report.AppendLine("Primary scan  - Folder 1 files : $($primaryFiles1.Count)")
$null = $report.AppendLine("Primary scan  - Folder 2 files : $($primaryFiles2.Count)")
$null = $report.AppendLine("QC scan       - Folder 1 files : $($qcFiles1.Count)")
$null = $report.AppendLine("QC scan       - Folder 2 files : $($qcFiles2.Count)")
$null = $report.AppendLine("QC Result                      : $(if ($discrepancyCount -eq 0) {'PASS - Both scans consistent'} else {"WARNING - $discrepancyCount discrepanc(ies) found (see ErrorLog.txt)"})")
$null = $report.AppendLine('')
$null = $report.AppendLine('--- Duplicate Summary --------------------------------------')
$null = $report.AppendLine("Total duplicate base names found : $($finalDuplicates.Count)")
$null = $report.AppendLine('')

if ($finalDuplicates.Count -eq 0) {
    $null = $report.AppendLine('No duplicate file names were found between the two folders.')
} else {
    $null = $report.AppendLine('The following file base names (name without extension) exist in')
    $null = $report.AppendLine('BOTH Folder 1 and Folder 2 (including their sub-folders):')
    $null = $report.AppendLine('')
    $null = $report.AppendLine('=' * 60)

    foreach ($dupName in $finalDuplicates) {
        $null = $report.AppendLine('')
        $null = $report.AppendLine("DUPLICATE BASE NAME: $dupName")
        $null = $report.AppendLine('-' * 40)

        # Gather matching files from both primary scan lists
        $matches1 = $primaryFiles1 | Where-Object { $_.BaseName -ieq $dupName }
        $matches2 = $primaryFiles2 | Where-Object { $_.BaseName -ieq $dupName }

        $null = $report.AppendLine("  Folder 1 match(es): $($matches1.Count)")
        foreach ($m in $matches1) {
            $null = $report.AppendLine("    File : $($m.FileName)")
            $null = $report.AppendLine("    Path : $($m.FullPath)")
        }

        $null = $report.AppendLine('')
        $null = $report.AppendLine("  Folder 2 match(es): $($matches2.Count)")
        foreach ($m in $matches2) {
            $null = $report.AppendLine("    File : $($m.FileName)")
            $null = $report.AppendLine("    Path : $($m.FullPath)")
        }

        $null = $report.AppendLine('')
        $null = $report.AppendLine('=' * 60)
    }
}

$null = $report.AppendLine('')
$null = $report.AppendLine("--- End of Report $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ---")

# Write Duplicates.txt (read-only check: we never open the user's source files)
try {
    [System.IO.File]::WriteAllText($duplicatesTxt, $report.ToString(),
        [System.Text.Encoding]::UTF8)
    Write-Host "Duplicates report saved : $duplicatesTxt" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Failed to write Duplicates.txt - $($_.Exception.Message)" -ForegroundColor Red
    $sharedLog.Add("[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')][ERROR] Failed to write '$duplicatesTxt': $($_.Exception.Message)")
}

# Write ErrorLog.txt (only if there were any logged entries)
if ($sharedLog.Count -gt 0) {
    try {
        [System.IO.File]::WriteAllLines($errorLogTxt, $sharedLog,
            [System.Text.Encoding]::UTF8)
        Write-Host "Error/event log saved   : $errorLogTxt" -ForegroundColor Yellow
    } catch {
        Write-Host "ERROR: Failed to write ErrorLog.txt - $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host 'No errors or events to log - ErrorLog.txt not created.' -ForegroundColor Green
}

Write-Host ''

# =========================================================================== #
#  FINAL SUMMARY                                                               #
# =========================================================================== #
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host '   FINAL SUMMARY' -ForegroundColor Cyan
Write-Host ('=' * 60) -ForegroundColor Cyan
Write-Host "Folder 1 files scanned   : $($primaryFiles1.Count)" -ForegroundColor White
Write-Host "Folder 2 files scanned   : $($primaryFiles2.Count)" -ForegroundColor White
Write-Host "QC verification          : $(if ($discrepancyCount -eq 0) {'PASSED'} else {'WARNING - see ErrorLog.txt'})" -ForegroundColor $(if ($discrepancyCount -eq 0) {'Green'} else {'Yellow'})
Write-Host "Duplicates found         : $($finalDuplicates.Count)" -ForegroundColor $(if ($finalDuplicates.Count -gt 0) {'Yellow'} else {'Green'})
Write-Host "Errors / events logged   : $($sharedLog.Count)" -ForegroundColor $(if ($sharedLog.Count -gt 0) {'Yellow'} else {'Green'})
Write-Host ''
Write-Host "Output folder            : $duplicatesDir" -ForegroundColor Cyan
Write-Host "  Duplicates.txt         : $duplicatesTxt" -ForegroundColor Cyan
if ($sharedLog.Count -gt 0) {
    Write-Host "  ErrorLog.txt           : $errorLogTxt" -ForegroundColor Cyan
}
Write-Host ''

if ($finalDuplicates.Count -gt 0) {
    Write-Host 'Duplicate base names:' -ForegroundColor Yellow
    foreach ($d in $finalDuplicates) {
        Write-Host "  - $d" -ForegroundColor Yellow
    }
    Write-Host ''
}

Write-Host 'Script completed.' -ForegroundColor Green

#endregion ---- Main Script ---------------------------------------------------
