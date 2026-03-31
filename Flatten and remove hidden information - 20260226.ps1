#Requires -Version 5.1
<#
.SYNOPSIS
    Flattens Excel workbook content and removes hidden items across an entire folder of spreadsheets.

.DESCRIPTION
    Prompts the user for a folder path, then for every Excel file (.xlsx / .xlsm / .xls) found
    directly in that folder:
      - Flattens all pivot tables to static values (cell highlights and formatting preserved)
      - Converts chart objects to static images (position and size preserved)
      - Removes external data connections
      - Removes named ranges referencing external workbooks
      - Strips VBA macros from macro-enabled workbooks (.xlsm); output saved as .xlsx
      - Deletes hidden and very-hidden worksheets
      - Removes hidden rows and columns (within used range)
      - Removes hidden ListObjects (tables)
      - Runs a QC pass to verify no residual hidden content
      - Saves cleaned files to <source folder>\Output\Passed\ (QC passed) or
        <source folder>\Output\Failed_QC\ (QC issues), using the original filename
        (macro-enabled files are saved with .xlsx extension)

    Results.txt (per-file detail + grand totals) and Error.txt are written to
    <source folder>\Output\.
    Excel COM is restarted every 25 files to prevent memory pressure on large batches (100+ files).

.NOTES
    - No external modules required.
    - Runs under the account that launched the PowerShell terminal.
    - Supports local and UNC/network paths.
    - "Trust access to the VBA project object model" (Excel Trust Center) enables in-place VBA
      stripping; if unavailable, saving as .xlsx eliminates all macros regardless.
#>

[CmdletBinding()]
param()   # No parameters - folder path is prompted interactively.

Set-StrictMode -Off   # Allow unset variables without terminating.

# Restart Excel COM every this many files to prevent memory buildup on large batches.
$ExcelRestartInterval = 25

# ===========================================================================
# HELPER FUNCTIONS
# ===========================================================================

function Write-Log {
    param([string]$FilePath, [string]$Message)
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Add-Content -LiteralPath $FilePath -Value $line -Encoding UTF8
    Write-Host $line
}

function Write-ErrorLog {
    param([string]$FilePath, [string]$Message)
    $line = "[ERROR][$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Add-Content -LiteralPath $FilePath -Value $line -Encoding UTF8
    Write-Warning $line
}

function Release-Com {
    param($obj)
    if ($null -ne $obj) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
    }
}

function New-ExcelInstance {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible                = $false
    $xl.DisplayAlerts          = $false
    $xl.AskToUpdateLinks       = $false
    $xl.AlertBeforeOverwriting = $false
    $xl.AutomationSecurity     = 3   # msoAutomationSecurityForceDisable — prevents macros running on open
    return $xl
}

function New-FileCounts {
    # Returns a fresh ordered hashtable of counters for one file.
    return [ordered]@{
        "Pivot Tables Flattened"                   = 0
        "Formulas Flattened"                       = 0
        "Charts Converted to Images"               = 0
        "Macro Modules Removed"                    = 0
        "External Connections Removed"             = 0
        "Named Ranges Pointing Externally Removed" = 0
        "Hidden Sheets Removed"                    = 0
        "Very Hidden Sheets Removed"               = 0
        "Hidden Rows Removed"                      = 0
        "Hidden Columns Removed"                   = 0
        "Hidden Tables (ListObjects) Removed"      = 0
        "QC Issues Found After Processing"         = 0
    }
}

# ---------------------------------------------------------------------------
# Processes a single workbook. Returns a hashtable with:
#   Counts   - ordered hashtable of action counts
#   Errors   - list of error strings
#   Skipped  - $true if the file was skipped entirely
# ---------------------------------------------------------------------------
function Invoke-ProcessWorkbook {
    param(
        [object]$Excel,          # Live Excel COM application object
        [string]$FilePath,       # Full path to source workbook
        [string]$PassedDir,      # Output\Passed\ — files that pass QC
        [string]$FailedQcDir,    # Output\Failed_QC\ — files with QC issues
        [string]$ResultsFile,    # Path to Results.txt (already open for append)
        [string]$ErrorFile       # Path to Error.txt  (already open for append)
    )

    $counts    = New-FileCounts
    $errorList = [System.Collections.Generic.List[string]]::new()
    $fileItem  = Get-Item -LiteralPath $FilePath
    $isMacroFile = $fileItem.Extension.ToLower() -eq ".xlsm"

    # Excel visibility constants
    $xlSheetVisible    = -1
    $xlSheetHidden     =  0
    $xlSheetVeryHidden =  2

    Write-Log $ResultsFile ""
    Write-Log $ResultsFile ("=" * 60)
    Write-Log $ResultsFile "  FILE: $($fileItem.Name)$(if ($isMacroFile) { '  [macro-enabled .xlsm]' })"
    Write-Log $ResultsFile ("=" * 60)

    # --- Verify read access before handing to Excel ---
    try {
        $fs = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'ReadWrite')
        $fs.Close(); $fs.Dispose()
    } catch {
        $msg = "Cannot open '$FilePath' for reading. File may be locked or permissions denied. Details: $($_.Exception.Message)"
        Write-ErrorLog $ErrorFile $msg
        $errorList.Add($msg)
        return @{ Counts = $counts; Errors = $errorList; Skipped = $true; QcFailed = $false }
    }

    # --- Open workbook ---
    $workbook = $null
    try {
        $workbook = $Excel.Workbooks.Open(
            $FilePath,
            0,        # UpdateLinks - don't update
            $false,   # ReadOnly
            5,        # Format
            "",       # Password
            "",       # WriteResPassword
            $true,    # IgnoreReadOnlyRecommended
            [System.Reflection.Missing]::Value,
            [System.Reflection.Missing]::Value,
            $false,
            $false,
            [System.Reflection.Missing]::Value,
            $false
        )
    } catch {
        $msg = "Failed to open workbook '$FilePath'. Details: $($_.Exception.Message)"
        Write-ErrorLog $ErrorFile $msg
        $errorList.Add($msg)
        return @{ Counts = $counts; Errors = $errorList; Skipped = $true; QcFailed = $false }
    }

    Write-Log $ResultsFile "  Workbook opened successfully."

    # -----------------------------------------------------------------------
    # STEP A - Remove external data connections
    # -----------------------------------------------------------------------
    Write-Log $ResultsFile "  --- External connections ---"
    try {
        $connCount = $workbook.Connections.Count
        Write-Log $ResultsFile "    Found: $connCount"
        for ($c = $connCount; $c -ge 1; $c--) {
            try {
                $conn = $workbook.Connections.Item($c)
                $connName = $conn.Name
                $conn.Delete()
                $counts["External Connections Removed"]++
                Write-Log $ResultsFile "    Removed connection: '$connName'"
                Release-Com $conn
            } catch {
                $msg = "    Could not remove connection [$c] [$($fileItem.Name)]: $($_.Exception.Message)"
                Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
            }
        }
    } catch {
        $msg = "    Error enumerating connections [$($fileItem.Name)]: $($_.Exception.Message)"
        Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
    }

    # -----------------------------------------------------------------------
    # STEP B - Remove externally-referencing named ranges
    # -----------------------------------------------------------------------
    Write-Log $ResultsFile "  --- External named ranges ---"
    try {
        $nameCount = $workbook.Names.Count
        Write-Log $ResultsFile "    Named ranges found: $nameCount"
        $externalNames = @()
        for ($n = 1; $n -le $nameCount; $n++) {
            try {
                $nm = $workbook.Names.Item($n)
                if ($nm.RefersTo -match '\[') { $externalNames += $nm.Name }
                Release-Com $nm
            } catch {}
        }
        Write-Log $ResultsFile "    External named ranges found: $($externalNames.Count)"
        foreach ($eName in $externalNames) {
            try {
                $nm = $workbook.Names.Item($eName)
                $nm.Delete()
                $counts["Named Ranges Pointing Externally Removed"]++
                Write-Log $ResultsFile "    Removed: '$eName'"
                Release-Com $nm
            } catch {
                $msg = "    Could not remove named range '$eName' [$($fileItem.Name)]: $($_.Exception.Message)"
                Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
            }
        }
    } catch {
        $msg = "    Error enumerating named ranges [$($fileItem.Name)]: $($_.Exception.Message)"
        Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
    }

    # -----------------------------------------------------------------------
    # STEP C - Catalogue sheet visibility
    # -----------------------------------------------------------------------
    $hiddenSheetNames     = @()
    $veryHiddenSheetNames = @()
    $visibleSheetNames    = @()
    $totalSheets          = $workbook.Sheets.Count

    Write-Log $ResultsFile "  --- Sheet inventory (total: $totalSheets) ---"

    for ($s = 1; $s -le $totalSheets; $s++) {
        try {
            $sh = $workbook.Sheets.Item($s)
            switch ($sh.Visible) {
                $xlSheetVisible    { $visibleSheetNames    += $sh.Name }
                $xlSheetHidden     { $hiddenSheetNames     += $sh.Name }
                $xlSheetVeryHidden { $veryHiddenSheetNames += $sh.Name }
            }
            Release-Com $sh
        } catch {
            $msg = "    Could not read sheet index $s visibility [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }
    }

    Write-Log $ResultsFile "    Visible sheets      : $($visibleSheetNames.Count)  -> $($visibleSheetNames -join ', ')"
    Write-Log $ResultsFile "    Hidden sheets       : $($hiddenSheetNames.Count)  -> $($hiddenSheetNames -join ', ')"
    Write-Log $ResultsFile "    Very-hidden sheets  : $($veryHiddenSheetNames.Count)  -> $($veryHiddenSheetNames -join ', ')"

    # -----------------------------------------------------------------------
    # STEP D - Process each visible sheet
    # -----------------------------------------------------------------------
    Write-Log $ResultsFile "  --- Processing visible sheets ---"

    foreach ($shName in $visibleSheetNames) {
        Write-Log $ResultsFile "    >> Sheet: '$shName'"

        $ws = $null
        try {
            $ws = $workbook.Sheets.Item($shName)
        } catch {
            $msg = "      Could not access sheet '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
            continue
        }

        # Unprotect the sheet so that all operations (delete rows/cols, pivot flatten, etc.) succeed.
        # The output is a sanitised copy so sheet protection is intentionally not restored.
        try {
            if ($ws.ProtectContents -or $ws.ProtectDrawingObjects -or $ws.ProtectScenarios) {
                $ws.Unprotect()
                Write-Log $ResultsFile "      Sheet unprotected (no password)"
            }
        } catch {
            Write-Log $ResultsFile "      WARNING: Sheet '$shName' is password-protected and could not be unprotected; some operations may fail"
        }

        # D1: Flatten pivot tables
        try {
            $ptCount = $ws.PivotTables().Count
            Write-Log $ResultsFile "      Pivot tables found: $ptCount"
            for ($p = $ptCount; $p -ge 1; $p--) {
                try {
                    $pt      = $ws.PivotTables($p)
                    $ptName  = $pt.Name
                    $ptRange = $pt.TableRange2

                    if ($null -eq $ptRange) {
                        # Pivot table has no resolvable range (broken/disconnected source).
                        # Clear just the cache and skip copy-paste.
                        try { $pt.PivotCache().MissingItemsLimit = 0 } catch {}
                        Write-Log $ResultsFile "      Skipped pivot table '$ptName' (null range - broken data source)"
                        Release-Com $pt
                        continue
                    }

                    # Stage a copy on a temp sheet so formats bake as real Interior.Color
                    # (pasting formats onto a live pivot table is overridden by the pivot engine;
                    #  formats only persist once the pivot is gone from the target cells)
                    $ptRange.Copy()
                    $tmpSheet = $null
                    try {
                        $tmpSheet = $ws.Parent.Worksheets.Add()
                        $tmpSheet.Cells(1, 1).PasteSpecial(-4163)   # xlPasteValues  — stage values
                        $tmpSheet.Cells(1, 1).PasteSpecial(-4122)   # xlPasteFormats — stage formats (no pivot engine here, they stick)

                        # Flatten the original range — this destroys the pivot data connection
                        $ptRange.PasteSpecial(-4163)   # xlPasteValues

                        # Discard the now-dead pivot object
                        try {
                            $ptObj = $ws.PivotTables($p)
                            $ptObj.TableRange2.ClearOutline()
                            Release-Com $ptObj
                        } catch {}

                        # Wipe residual pivot-table cell formatting so it cannot interfere with the restore
                        $ptRange.ClearFormats()

                        # Restore colours from temp sheet — pivot engine is gone and cells are clean, so formats persist
                        $tmpRange = $tmpSheet.Range(
                            $tmpSheet.Cells(1, 1),
                            $tmpSheet.Cells($ptRange.Rows.Count, $ptRange.Columns.Count)
                        )
                        $tmpRange.Copy()
                        $ptRange.PasteSpecial(-4122)   # xlPasteFormats — colours now stick on static cells
                        Release-Com $tmpRange
                    } finally {
                        if ($null -ne $tmpSheet) {
                            $Excel.DisplayAlerts = $false
                            $tmpSheet.Delete()
                            $Excel.DisplayAlerts = $true
                            Release-Com $tmpSheet
                        }
                    }

                    $counts["Pivot Tables Flattened"]++
                    Write-Log $ResultsFile "      Flattened pivot table: '$ptName'"
                    Release-Com $ptRange; Release-Com $pt
                } catch {
                    $msg = "      Error flattening pivot table $p in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
            }
        } catch {
            $msg = "      Error accessing pivot tables in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        # D2: Flatten remaining formulas to static values
        # SpecialCells(-4123) = xlCellTypeFormulas selects only formula cells so we
        # never allocate a single array covering the whole used range (which OOMs on
        # large sheets).  Each Area is a small contiguous block; Value2=Value2 on each
        # Area replaces formulas in-place without the clipboard, so merged cells are
        # handled correctly.  SpecialCells throws when no formula cells exist — caught
        # by the inner try/catch and treated as "nothing to do".
        try {
            $usedRng = $ws.UsedRange
            try {
                $fCells = $usedRng.SpecialCells(-4123)   # xlCellTypeFormulas
                foreach ($area in $fCells.Areas) {
                    $area.Value2 = $area.Value2
                    Release-Com $area
                }
                Release-Com $fCells
            } catch {
                # No formula cells found on this sheet — nothing to flatten.
            }
            $counts["Formulas Flattened"]++
            Write-Log $ResultsFile "      Formulas flattened (used range)"
            Release-Com $usedRng
        } catch {
            $msg = "      Error flattening formulas in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        # D3: Convert charts to static images
        try {
            $coCount = $ws.ChartObjects().Count
            Write-Log $ResultsFile "      Chart objects found: $coCount"
            for ($ch = $coCount; $ch -ge 1; $ch--) {
                try {
                    $co     = $ws.ChartObjects($ch)
                    $coName = $co.Name
                    $coLeft = $co.Left; $coTop = $co.Top
                    $coW    = $co.Width; $coH  = $co.Height

                    $co.CopyPicture(1, -4147)   # xlScreen, xlPicture
                    $ws.Paste()
                    # Note: do NOT set $Excel.CutCopyMode here — the property only accepts
                    # xlCopy/xlCut enum values; clearing it is unnecessary in automation
                    # (DisplayAlerts=false suppresses any clipboard prompt on exit).

                    $pic = $ws.Shapes.Item($ws.Shapes.Count)
                    $pic.Left = $coLeft; $pic.Top  = $coTop
                    $pic.Width= $coW;    $pic.Height= $coH
                    Release-Com $pic

                    $co.Delete()
                    $counts["Charts Converted to Images"]++
                    Write-Log $ResultsFile "      Converted chart to image: '$coName'"
                    Release-Com $co
                } catch {
                    $msg = "      Error converting chart $ch in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
            }
        } catch {
            $msg = "      Error accessing charts in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        # D4: Remove hidden rows
        # Scan to build a Union of all hidden rows, then delete in one operation.
        # This is O(n) instead of the O(n^2) caused by shifting rows after each
        # individual Delete() call.
        try {
            $usedRange      = $ws.UsedRange
            $firstRow       = $usedRange.Row
            $lastRow        = $firstRow + $usedRange.Rows.Count - 1
            $hiddenRowUnion = $null
            $hiddenRowCount = 0

            for ($r = $firstRow; $r -le $lastRow; $r++) {
                try {
                    $rObj = $ws.Rows.Item($r)
                    if ($rObj.Hidden) {
                        $hiddenRowCount++
                        $hiddenRowUnion = if ($null -eq $hiddenRowUnion) { $rObj } else { $Excel.Union($hiddenRowUnion, $rObj) }
                    } else {
                        Release-Com $rObj
                    }
                } catch {}
            }
            Write-Log $ResultsFile "      Hidden rows found: $hiddenRowCount"
            if ($null -ne $hiddenRowUnion) {
                try {
                    $hiddenRowUnion.Delete()
                    $counts["Hidden Rows Removed"] += $hiddenRowCount
                    Write-Log $ResultsFile "      Removed $hiddenRowCount hidden row(s)."
                } catch {
                    $msg = "      Error deleting hidden rows in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
                Release-Com $hiddenRowUnion
            }
            Release-Com $usedRange
        } catch {
            $msg = "      Error processing hidden rows in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        # D5: Remove hidden columns
        # Same union-then-delete approach as D4 — one Delete() call regardless of
        # how many hidden columns exist.
        try {
            $usedRange      = $ws.UsedRange
            $firstCol       = $usedRange.Column
            $lastCol        = $firstCol + $usedRange.Columns.Count - 1
            $hiddenColUnion = $null
            $hiddenColCount = 0

            for ($c = $firstCol; $c -le $lastCol; $c++) {
                try {
                    $cObj = $ws.Columns.Item($c)
                    if ($cObj.Hidden) {
                        $hiddenColCount++
                        $hiddenColUnion = if ($null -eq $hiddenColUnion) { $cObj } else { $Excel.Union($hiddenColUnion, $cObj) }
                    } else {
                        Release-Com $cObj
                    }
                } catch {}
            }
            Write-Log $ResultsFile "      Hidden columns found: $hiddenColCount"
            if ($null -ne $hiddenColUnion) {
                try {
                    $hiddenColUnion.Delete()
                    $counts["Hidden Columns Removed"] += $hiddenColCount
                    Write-Log $ResultsFile "      Removed $hiddenColCount hidden column(s)."
                } catch {
                    $msg = "      Error deleting hidden columns in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
                Release-Com $hiddenColUnion
            }
            Release-Com $usedRange
        } catch {
            $msg = "      Error processing hidden columns in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        # D6: Remove hidden ListObjects (tables)
        try {
            $loCount     = $ws.ListObjects.Count
            $hiddenTbls  = @()
            Write-Log $ResultsFile "      ListObjects found: $loCount"
            for ($lo = 1; $lo -le $loCount; $lo++) {
                try {
                    $loObj   = $ws.ListObjects.Item($lo)
                    $loRange = $loObj.Range
                    if ($loRange.EntireRow.Hidden -and $loRange.EntireColumn.Hidden) {
                        $hiddenTbls += $loObj.Name
                    }
                    Release-Com $loRange; Release-Com $loObj
                } catch {}
            }
            Write-Log $ResultsFile "      Hidden tables found: $($hiddenTbls.Count)"
            foreach ($tblName in $hiddenTbls) {
                try {
                    $loObj = $ws.ListObjects.Item($tblName)
                    $loObj.Delete()
                    $counts["Hidden Tables (ListObjects) Removed"]++
                    Write-Log $ResultsFile "      Removed hidden table: '$tblName'"
                    Release-Com $loObj
                } catch {
                    $msg = "      Error removing table '$tblName' in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
            }
        } catch {
            $msg = "      Error processing tables in '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        Release-Com $ws
    }

    # -----------------------------------------------------------------------
    # STEP E - Delete hidden and very-hidden sheets
    # -----------------------------------------------------------------------
    Write-Log $ResultsFile "  --- Deleting hidden sheets ---"
    foreach ($shName in $hiddenSheetNames) {
        try {
            $sh = $workbook.Sheets.Item($shName)
            $sh.Visible = $xlSheetVisible
            $sh.Delete()
            $counts["Hidden Sheets Removed"]++
            Write-Log $ResultsFile "    Deleted hidden sheet: '$shName'"
            Release-Com $sh
        } catch {
            $msg = "    Error deleting hidden sheet '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }
    }
    foreach ($shName in $veryHiddenSheetNames) {
        try {
            $sh = $workbook.Sheets.Item($shName)
            $sh.Visible = $xlSheetVisible
            $sh.Delete()
            $counts["Very Hidden Sheets Removed"]++
            Write-Log $ResultsFile "    Deleted very-hidden sheet: '$shName'"
            Release-Com $sh
        } catch {
            $msg = "    Error deleting very-hidden sheet '$shName' [$($fileItem.Name)]: $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }
    }

    # -----------------------------------------------------------------------
    # STEP F - QC pass
    # -----------------------------------------------------------------------
    Write-Log $ResultsFile "  --- QC pass ---"
    $qcIssues = [System.Collections.Generic.List[string]]::new()

    # F1: Sheet visibility
    for ($s = 1; $s -le $workbook.Sheets.Count; $s++) {
        try {
            $sh = $workbook.Sheets.Item($s)
            if ($sh.Visible -ne $xlSheetVisible) {
                $issue = "QC ISSUE: Sheet '$($sh.Name)' is still hidden (Visible=$($sh.Visible))."
                $qcIssues.Add($issue); Write-Log $ResultsFile "    $issue"
            }
            Release-Com $sh
        } catch {}
    }

    # F2: Pivot tables, charts, hidden rows/cols on visible sheets
    for ($s = 1; $s -le $workbook.Sheets.Count; $s++) {
        try {
            $sh = $workbook.Sheets.Item($s)
            if ($sh.Visible -ne $xlSheetVisible) { Release-Com $sh; continue }
            $shQC = $sh.Name

            try {
                $ptQC = $sh.PivotTables().Count
                if ($ptQC -gt 0) {
                    $issue = "QC ISSUE: Sheet '$shQC' still has $ptQC pivot table(s)."
                    $qcIssues.Add($issue); Write-Log $ResultsFile "    $issue"
                }
            } catch {}

            try {
                $coQC = $sh.ChartObjects().Count
                if ($coQC -gt 0) {
                    $issue = "QC ISSUE: Sheet '$shQC' still has $coQC chart object(s)."
                    $qcIssues.Add($issue); Write-Log $ResultsFile "    $issue"
                }
            } catch {}

            try {
                $ur     = $sh.UsedRange
                $rStart = $ur.Row; $rEnd = $rStart + $ur.Rows.Count - 1
                $hidR   = 0
                for ($r = $rStart; $r -le $rEnd; $r++) {
                    $rObj = $sh.Rows.Item($r)
                    if ($rObj.Hidden) { $hidR++ }
                    Release-Com $rObj
                }
                if ($hidR -gt 0) {
                    $issue = "QC ISSUE: Sheet '$shQC' still has $hidR hidden row(s)."
                    $qcIssues.Add($issue); Write-Log $ResultsFile "    $issue"
                }
                Release-Com $ur
            } catch {}

            try {
                $ur     = $sh.UsedRange
                $cStart = $ur.Column; $cEnd = $cStart + $ur.Columns.Count - 1
                $hidC   = 0
                for ($c = $cStart; $c -le $cEnd; $c++) {
                    $cObj = $sh.Columns.Item($c)
                    if ($cObj.Hidden) { $hidC++ }
                    Release-Com $cObj
                }
                if ($hidC -gt 0) {
                    $issue = "QC ISSUE: Sheet '$shQC' still has $hidC hidden column(s)."
                    $qcIssues.Add($issue); Write-Log $ResultsFile "    $issue"
                }
                Release-Com $ur
            } catch {}

            Release-Com $sh
        } catch {}
    }

    # F3: Connections
    try {
        $connQC = $workbook.Connections.Count
        if ($connQC -gt 0) {
            $issue = "QC ISSUE: Workbook still has $connQC external connection(s)."
            $qcIssues.Add($issue); Write-Log $ResultsFile "    $issue"
        }
    } catch {}

    $counts["QC Issues Found After Processing"] = $qcIssues.Count

    if ($qcIssues.Count -eq 0) {
        Write-Log $ResultsFile "    QC PASSED - No residual hidden content found."
    } else {
        Write-Log $ResultsFile "    QC COMPLETED WITH $($qcIssues.Count) ISSUE(S). See details above."
        foreach ($qi in $qcIssues) { Write-ErrorLog $ErrorFile "    [$($fileItem.Name)] $qi" }
    }

    # -----------------------------------------------------------------------
    # STEP G0 - Strip VBA macros from macro-enabled workbooks (.xlsm)
    # -----------------------------------------------------------------------
    if ($isMacroFile) {
        Write-Log $ResultsFile "  --- Stripping VBA macros (.xlsm) ---"
        try {
            $vbp      = $workbook.VBProject
            # vbext_pp_locked = 1 means the VBProject is password-protected
            if ($vbp.Protection -eq 1) {
                Write-Log $ResultsFile "    VBProject is password-protected; macros will be eliminated by saving as .xlsx"
            } else {
                $vbComps      = $vbp.VBComponents
                $removedCount = 0
                for ($vi = $vbComps.Count; $vi -ge 1; $vi--) {
                    try {
                        $vbc = $vbComps.Item($vi)
                        # vbext_ct_Document (type 100) = sheet/workbook module — cannot remove, clear code only
                        if ($vbc.Type -eq 100) {
                            $cm = $vbc.CodeModule
                            if ($cm.CountOfLines -gt 0) {
                                $cm.DeleteLines(1, $cm.CountOfLines)
                                $removedCount++
                            }
                        } else {
                            $vbComps.Remove($vbc)
                            $removedCount++
                        }
                    } catch {}
                }
                $counts["Macro Modules Removed"] += $removedCount
                Write-Log $ResultsFile "    Stripped $removedCount VBA component(s) via VBProject"
            }
        } catch {
            # VBProject access requires 'Trust access to the VBA project object model' in Trust Center.
            # If unavailable, saving as .xlsx below is the fallback guarantee.
            Write-Log $ResultsFile "    VBProject access unavailable (Trust Center setting required); macros eliminated by saving as .xlsx"
        }
    }

    # -----------------------------------------------------------------------
    # STEP G - Save to Passed or Failed_QC folder
    # Macro-enabled files are always saved as .xlsx to guarantee macro elimination.
    # -----------------------------------------------------------------------
    $hasQcIssues = $qcIssues.Count -gt 0
    $saveExt     = if ($isMacroFile) { ".xlsx" } else { $fileItem.Extension.ToLower() }
    $saveName    = [System.IO.Path]::GetFileNameWithoutExtension($fileItem.Name) + $saveExt
    $destDir     = if ($hasQcIssues) { $FailedQcDir } else { $PassedDir }
    $savePath    = Join-Path $destDir $saveName
    $destLabel   = if ($hasQcIssues) { "Failed_QC" } else { "Passed" }

    Write-Log $ResultsFile "  --- Saving [$destLabel]: $savePath ---"

    $xlFileFormat = switch ($saveExt) {
        ".xlsx" { 51 }   # xlOpenXMLWorkbook       — also strips any residual macros from xlsm
        ".xls"  { 56 }   # xlExcel8
        default { 51 }
    }

    try {
        $workbook.SaveAs(
            $savePath,
            $xlFileFormat,
            [System.Reflection.Missing]::Value,
            [System.Reflection.Missing]::Value,
            $false, $false, 1,
            [System.Reflection.Missing]::Value,
            $false,
            [System.Reflection.Missing]::Value,
            [System.Reflection.Missing]::Value,
            $false
        )
        $saveNote = if ($isMacroFile) { " (saved as .xlsx - macros eliminated)" } else { "" }
        Write-Log $ResultsFile "  Saved successfully$saveNote : $savePath"
    } catch {
        $msg = "  Could not save '$savePath' [$($fileItem.Name)]. Details: $($_.Exception.Message)"
        Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
    }

    # -----------------------------------------------------------------------
    # STEP H - Close workbook (don't save again)
    # -----------------------------------------------------------------------
    try { $workbook.Close($false) } catch {}
    Release-Com $workbook

    # Per-file summary in results
    Write-Log $ResultsFile "  --- File summary: $($fileItem.Name) ---"
    $maxLen = ($counts.Keys | Measure-Object -Property Length -Maximum).Maximum
    foreach ($key in $counts.Keys) {
        $pad  = " " * ($maxLen - $key.Length)
        $line = "    $key$pad : $($counts[$key])"
        Add-Content -LiteralPath $ResultsFile -Value $line -Encoding UTF8
        Write-Host $line
    }

    return @{ Counts = $counts; Errors = $errorList; Skipped = $false; QcFailed = $hasQcIssues }
}

# ===========================================================================
# MAIN SCRIPT BODY
# ===========================================================================

Write-Host ""
Write-Host "========================================================"
Write-Host "  Flatten & Remove Hidden Information - Batch Mode"
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "========================================================"
Write-Host ""

# ---------------------------------------------------------------------------
# Prompt for folder path
# ---------------------------------------------------------------------------
do {
    $FolderPath = (Read-Host "Enter the full path to the folder containing the spreadsheets").Trim()

    if ([string]::IsNullOrWhiteSpace($FolderPath)) {
        Write-Warning "No path entered. Please try again."
        continue
    }

    # Resolve to absolute path (handles relative paths and env vars)
    try {
        $FolderPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($FolderPath)
    } catch {
        Write-Warning "Could not resolve path '$FolderPath'. Please enter a valid path."
        $FolderPath = ""
        continue
    }

    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        Write-Warning "The path '$FolderPath' does not exist or is not a folder. Please try again."
        $FolderPath = ""
    }
} while ([string]::IsNullOrWhiteSpace($FolderPath))

Write-Host ""
Write-Host "Source folder : $FolderPath"

# ---------------------------------------------------------------------------
# Discover Excel files
# ---------------------------------------------------------------------------
$excelFiles = @(Get-ChildItem -LiteralPath $FolderPath -File |
    Where-Object { $_.Extension -match '^\.(xlsx|xlsm|xls)$' } |
    Sort-Object Name)

if ($excelFiles.Count -eq 0) {
    Write-Error "No Excel files (.xlsx, .xlsm, .xls) found in '$FolderPath'. Nothing to process."
    exit 1
}

Write-Host "Excel files found: $($excelFiles.Count)"
$excelFiles | ForEach-Object { Write-Host "  - $($_.Name)" }
Write-Host ""

# ---------------------------------------------------------------------------
# Build Output folders - inside the source folder, at the same level as the files
# ---------------------------------------------------------------------------
$outputDir   = Join-Path $FolderPath "Output"
$passedDir   = Join-Path $outputDir "Passed"
$failedQcDir = Join-Path $outputDir "Failed_QC"

foreach ($dir in @($outputDir, $passedDir, $failedQcDir)) {
    if (-not (Test-Path -LiteralPath $dir)) {
        try {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Write-Host "Created: $dir"
        } catch {
            Write-Error "FATAL: Could not create directory '$dir'. Details: $($_.Exception.Message)"
            exit 1
        }
    }
}

Write-Host "Output (passed)  : $passedDir"
Write-Host "Output (failed)  : $failedQcDir"

$resultsFile = Join-Path $outputDir "Results.txt"
$errorFile   = Join-Path $outputDir "Error.txt"

# Initialise log files
$runHeader = "=" * 60
$runDate   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

Set-Content -LiteralPath $resultsFile -Value $runHeader -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Results - Flatten & Remove Hidden Information (Batch)" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Run date      : $runDate" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Source folder : $FolderPath" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Files found   : $($excelFiles.Count)" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value $runHeader -Encoding UTF8

Set-Content -LiteralPath $errorFile -Value $runHeader -Encoding UTF8
Add-Content -LiteralPath $errorFile -Value "  Error Log - Flatten & Remove Hidden Information (Batch)" -Encoding UTF8
Add-Content -LiteralPath $errorFile -Value "  Run date      : $runDate" -Encoding UTF8
Add-Content -LiteralPath $errorFile -Value "  Source folder : $FolderPath" -Encoding UTF8
Add-Content -LiteralPath $errorFile -Value $runHeader -Encoding UTF8

Write-Log $resultsFile "Batch processing started. Files to process: $($excelFiles.Count)"

# ---------------------------------------------------------------------------
# Launch Excel once for the entire batch
# ---------------------------------------------------------------------------
Write-Log $resultsFile "Launching Microsoft Excel COM (current user context)."

$excel = $null
try {
    $excel = New-ExcelInstance
} catch {
    $msg = "FATAL: Could not create Excel COM object. Is Microsoft Excel installed? Details: $($_.Exception.Message)"
    Write-ErrorLog $errorFile $msg
    Write-Error $msg
    exit 1
}

Write-Log $resultsFile "Excel COM object ready. AutomationSecurity=ForceDisable (macros suppressed on open)."

# ---------------------------------------------------------------------------
# Initialise grand-total counters
# ---------------------------------------------------------------------------
$grandTotals      = New-FileCounts
$skippedFiles     = [System.Collections.Generic.List[string]]::new()
$allErrors        = [System.Collections.Generic.List[string]]::new()
$filesProcessed   = 0
$qcFailedFiles    = [System.Collections.Generic.List[string]]::new()
$batchStart       = [datetime]::UtcNow

# ---------------------------------------------------------------------------
# Process each file
# ---------------------------------------------------------------------------
$fileIndex = 0
foreach ($fileItem in $excelFiles) {
    $fileIndex++
    $pct     = [math]::Round($fileIndex / $excelFiles.Count * 100, 1)
    $elapsed = ([datetime]::UtcNow - $batchStart).TotalSeconds
    $eta     = if ($fileIndex -gt 1) {
                   $secsPerFile = $elapsed / ($fileIndex - 1)
                   $remaining   = [int]($secsPerFile * ($excelFiles.Count - $fileIndex + 1))
                   "ETA ~${remaining}s"
               } else { "ETA calculating..." }

    Write-Host ""
    Write-Host "[$fileIndex / $($excelFiles.Count)] ($pct%)  $($fileItem.Name)  - $eta"

    # Restart Excel every $ExcelRestartInterval files to release COM memory pressure
    if ($fileIndex -gt 1 -and (($fileIndex - 1) % $ExcelRestartInterval) -eq 0) {
        Write-Log $resultsFile "Restarting Excel COM after $($fileIndex - 1) files (memory management)."
        try { $excel.Quit() } catch {}
        Release-Com $excel
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        try {
            $excel = New-ExcelInstance
            Write-Log $resultsFile "Excel COM restarted successfully."
        } catch {
            $msg = "FATAL: Could not restart Excel COM at file $fileIndex. Aborting."
            Write-ErrorLog $errorFile $msg; Write-Error $msg; break
        }
    }

    $result = Invoke-ProcessWorkbook `
        -Excel       $excel `
        -FilePath    $fileItem.FullName `
        -PassedDir   $passedDir `
        -FailedQcDir $failedQcDir `
        -ResultsFile $resultsFile `
        -ErrorFile   $errorFile

    if ($result.Skipped) {
        $skippedFiles.Add($fileItem.Name)
        Write-Log $resultsFile "  !! SKIPPED: $($fileItem.Name)"
    } else {
        $filesProcessed++
        foreach ($key in @($grandTotals.Keys)) {
            $grandTotals[$key] += $result.Counts[$key]
        }
        if ($result.QcFailed) { $qcFailedFiles.Add($fileItem.Name) }
    }

    foreach ($e in $result.Errors) { $allErrors.Add($e) }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# ---------------------------------------------------------------------------
# Quit Excel
# ---------------------------------------------------------------------------
try { $excel.Quit() } catch {}
Release-Com $excel
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()

Write-Log $resultsFile "Excel closed."

# ---------------------------------------------------------------------------
# Grand-total summary in Results.txt
# ---------------------------------------------------------------------------
$totalElapsed = [math]::Round(([datetime]::UtcNow - $batchStart).TotalSeconds)

Add-Content -LiteralPath $resultsFile -Value "" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value ("=" * 60) -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  GRAND TOTAL SUMMARY" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Files in folder  : $($excelFiles.Count)" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Files processed  : $filesProcessed" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Files skipped    : $($skippedFiles.Count)  -> $($skippedFiles -join ', ')" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  QC passed        : $($filesProcessed - $qcFailedFiles.Count)" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  QC failed        : $($qcFailedFiles.Count)  -> $($qcFailedFiles -join ', ')" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Total elapsed    : ${totalElapsed}s" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value ("=" * 60) -Encoding UTF8

$maxLen = ($grandTotals.Keys | Measure-Object -Property Length -Maximum).Maximum
foreach ($key in $grandTotals.Keys) {
    $pad  = " " * ($maxLen - $key.Length)
    $line = "  $key$pad : $($grandTotals[$key])"
    Add-Content -LiteralPath $resultsFile -Value $line -Encoding UTF8
    Write-Host $line
}

Add-Content -LiteralPath $resultsFile -Value "" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Output (passed) : $passedDir" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Output (failed) : $failedQcDir" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Results log     : $resultsFile" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Error log       : $errorFile" -Encoding UTF8

# Error.txt footer
if ($allErrors.Count -gt 0) {
    Add-Content -LiteralPath $errorFile -Value "" -Encoding UTF8
    Add-Content -LiteralPath $errorFile -Value ("=" * 60) -Encoding UTF8
    Add-Content -LiteralPath $errorFile -Value "  ERRORS ENCOUNTERED ACROSS ALL FILES ($($allErrors.Count) total)" -Encoding UTF8
    Add-Content -LiteralPath $errorFile -Value ("=" * 60) -Encoding UTF8
    foreach ($e in $allErrors) {
        Add-Content -LiteralPath $errorFile -Value "  $e" -Encoding UTF8
    }
} else {
    Add-Content -LiteralPath $errorFile -Value "" -Encoding UTF8
    Add-Content -LiteralPath $errorFile -Value "  No errors encountered during this run." -Encoding UTF8
}

Write-Host ""
Write-Host "========================================================"
Write-Host "  Batch complete."
Write-Host "  Total elapsed    : ${totalElapsed}s"
Write-Host "  Files processed  : $filesProcessed / $($excelFiles.Count)"
Write-Host "  Files skipped    : $($skippedFiles.Count)"
Write-Host "  QC passed        : $($filesProcessed - $qcFailedFiles.Count)  -> $passedDir"
Write-Host "  QC failed        : $($qcFailedFiles.Count)  -> $failedQcDir"
Write-Host "  Results log      : $resultsFile"
Write-Host "  Error log        : $errorFile"
Write-Host "========================================================"
Write-Host ""
