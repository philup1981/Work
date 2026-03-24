#Requires -Version 5.1
<#
.SYNOPSIS
    Flattens Excel workbook content and removes hidden items across an entire folder of spreadsheets.

.DESCRIPTION
    Prompts the user for a folder path, then for every Excel file (.xlsx / .xlsm / .xls) found
    directly in that folder:
      - Flattens all pivot tables to static values
      - Converts chart objects to static images
      - Removes external data connections
      - Removes named ranges referencing external workbooks
      - Deletes hidden and very-hidden worksheets
      - Removes hidden rows and columns (within used range)
      - Removes hidden ListObjects (tables)
      - Runs a QC pass to verify no residual hidden content
      - Saves an Updated_<filename> copy to an Output folder one level above the source folder

    A single Results.txt (per-file detail + grand total) and Error.txt are written to the Output folder.

.NOTES
    - No external modules required.
    - Runs under the account that launched the PowerShell terminal.
    - Supports local and UNC/network paths.
#>

[CmdletBinding()]
param()   # No parameters - folder path is prompted interactively.

Set-StrictMode -Off   # Allow unset variables without terminating.

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

function New-FileCounts {
    # Returns a fresh ordered hashtable of counters for one file.
    return [ordered]@{
        "Pivot Tables Flattened"                   = 0
        "Formulas Flattened"                       = 0
        "Charts Converted to Images"               = 0
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
        [string]$OutputDir,      # Destination folder for Updated_ file
        [string]$ResultsFile,    # Path to Results.txt (already open for append)
        [string]$ErrorFile       # Path to Error.txt  (already open for append)
    )

    $counts    = New-FileCounts
    $errorList = [System.Collections.Generic.List[string]]::new()
    $fileItem  = Get-Item -LiteralPath $FilePath

    # Excel visibility constants
    $xlSheetVisible    = -1
    $xlSheetHidden     =  0
    $xlSheetVeryHidden =  2

    Write-Log $ResultsFile ""
    Write-Log $ResultsFile ("=" * 60)
    Write-Log $ResultsFile "  FILE: $($fileItem.Name)"
    Write-Log $ResultsFile ("=" * 60)

    # --- Verify read access before handing to Excel ---
    try {
        $fs = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'ReadWrite')
        $fs.Close(); $fs.Dispose()
    } catch {
        $msg = "Cannot open '$FilePath' for reading. File may be locked or permissions denied. Details: $($_.Exception.Message)"
        Write-ErrorLog $ErrorFile $msg
        $errorList.Add($msg)
        return @{ Counts = $counts; Errors = $errorList; Skipped = $true }
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
        return @{ Counts = $counts; Errors = $errorList; Skipped = $true }
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
                $msg = "    Could not remove connection [$c]: $($_.Exception.Message)"
                Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
            }
        }
    } catch {
        $msg = "    Error enumerating connections: $($_.Exception.Message)"
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
                $msg = "    Could not remove named range '$eName': $($_.Exception.Message)"
                Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
            }
        }
    } catch {
        $msg = "    Error enumerating named ranges: $($_.Exception.Message)"
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
            $msg = "    Could not read sheet index $s visibility: $($_.Exception.Message)"
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
            $msg = "      Could not access sheet '$shName': $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
            continue
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

                    $ptRange.Copy()
                    $ptRange.PasteSpecial(-4163)   # xlPasteValues
                    $pt.TableRange2.ClearContents()
                    $ptRange.PasteSpecial(-4163)

                    try {
                        $ptObj = $ws.PivotTables($p)
                        $ptObj.TableRange2.ClearOutline()
                        $ptObj.TableRange1.Clear()
                        $ptRange.PasteSpecial(-4122)   # xlPasteFormats  — restore cell highlights/colours
                        $ptRange.PasteSpecial(-4163)   # xlPasteValues
                        Release-Com $ptObj
                    } catch {}

                    $counts["Pivot Tables Flattened"]++
                    Write-Log $ResultsFile "      Flattened pivot table: '$ptName'"
                    Release-Com $ptRange; Release-Com $pt
                } catch {
                    $msg = "      Error flattening pivot table $p in '$shName': $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
            }
        } catch {
            $msg = "      Error accessing pivot tables in '$shName': $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        # D2: Flatten remaining formulas to static values
        try {
            $usedRng = $ws.UsedRange
            # Count formula cells for logging (SpecialCells throws when none exist)
            $formulaCount = 0
            try {
                $formulaCells = $usedRng.SpecialCells(-4123)   # xlCellTypeFormulas
                $formulaCount = $formulaCells.Count
                Release-Com $formulaCells
            } catch {}
            Write-Log $ResultsFile "      Formulas found: $formulaCount"
            if ($formulaCount -gt 0) {
                # Paste values over the entire UsedRange (always contiguous — avoids
                # the non-contiguous-range limitation of PasteSpecial on SpecialCells)
                $usedRng.Copy()
                $usedRng.PasteSpecial(-4163)                   # xlPasteValues
                $Excel.CutCopyMode = $false
                $counts["Formulas Flattened"] += $formulaCount
                Write-Log $ResultsFile "      Formulas flattened: $formulaCount"
            }
            Release-Com $usedRng
        } catch {
            $msg = "      Error flattening formulas in '$shName': $($_.Exception.Message)"
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
                    $Excel.CutCopyMode = $false

                    $pic = $ws.Shapes.Item($ws.Shapes.Count)
                    $pic.Left = $coLeft; $pic.Top  = $coTop
                    $pic.Width= $coW;    $pic.Height= $coH
                    Release-Com $pic

                    $co.Delete()
                    $counts["Charts Converted to Images"]++
                    Write-Log $ResultsFile "      Converted chart to image: '$coName'"
                    Release-Com $co
                } catch {
                    $msg = "      Error converting chart $ch in '$shName': $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
            }
        } catch {
            $msg = "      Error accessing charts in '$shName': $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        # D4: Remove hidden rows
        try {
            $usedRange        = $ws.UsedRange
            $firstRow         = $usedRange.Row
            $lastRow          = $firstRow + $usedRange.Rows.Count - 1
            $hiddenRowIndices = [System.Collections.Generic.List[int]]::new()

            for ($r = $firstRow; $r -le $lastRow; $r++) {
                try {
                    $rObj = $ws.Rows.Item($r)
                    if ($rObj.Hidden) { $hiddenRowIndices.Add($r) }
                    Release-Com $rObj
                } catch {}
            }
            Write-Log $ResultsFile "      Hidden rows found: $($hiddenRowIndices.Count)"
            $hiddenRowIndices.Reverse()
            foreach ($ri in $hiddenRowIndices) {
                try {
                    $rObj = $ws.Rows.Item($ri)
                    $rObj.Delete()
                    $counts["Hidden Rows Removed"]++
                    Release-Com $rObj
                } catch {
                    $msg = "      Error deleting hidden row $ri in '$shName': $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
            }
            if ($hiddenRowIndices.Count -gt 0) {
                Write-Log $ResultsFile "      Removed $($hiddenRowIndices.Count) hidden row(s)."
            }
            Release-Com $usedRange
        } catch {
            $msg = "      Error processing hidden rows in '$shName': $($_.Exception.Message)"
            Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
        }

        # D5: Remove hidden columns
        try {
            $usedRange        = $ws.UsedRange
            $firstCol         = $usedRange.Column
            $lastCol          = $firstCol + $usedRange.Columns.Count - 1
            $hiddenColIndices = [System.Collections.Generic.List[int]]::new()

            for ($c = $firstCol; $c -le $lastCol; $c++) {
                try {
                    $cObj = $ws.Columns.Item($c)
                    if ($cObj.Hidden) { $hiddenColIndices.Add($c) }
                    Release-Com $cObj
                } catch {}
            }
            Write-Log $ResultsFile "      Hidden columns found: $($hiddenColIndices.Count)"
            $hiddenColIndices.Reverse()
            foreach ($ci in $hiddenColIndices) {
                try {
                    $cObj = $ws.Columns.Item($ci)
                    $cObj.Delete()
                    $counts["Hidden Columns Removed"]++
                    Release-Com $cObj
                } catch {
                    $msg = "      Error deleting hidden column $ci in '$shName': $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
            }
            if ($hiddenColIndices.Count -gt 0) {
                Write-Log $ResultsFile "      Removed $($hiddenColIndices.Count) hidden column(s)."
            }
            Release-Com $usedRange
        } catch {
            $msg = "      Error processing hidden columns in '$shName': $($_.Exception.Message)"
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
                    $msg = "      Error removing table '$tblName' in '$shName': $($_.Exception.Message)"
                    Write-ErrorLog $ErrorFile $msg; $errorList.Add($msg)
                }
            }
        } catch {
            $msg = "      Error processing tables in '$shName': $($_.Exception.Message)"
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
            $msg = "    Error deleting hidden sheet '$shName': $($_.Exception.Message)"
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
            $msg = "    Error deleting very-hidden sheet '$shName': $($_.Exception.Message)"
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
        foreach ($qi in $qcIssues) { Write-ErrorLog $ErrorFile "    $qi" }
    }

    # -----------------------------------------------------------------------
    # STEP G - Save Updated_ copy
    # -----------------------------------------------------------------------
    $updatedName = "Updated_" + $fileItem.Name
    $updatedPath = Join-Path $OutputDir $updatedName

    Write-Log $ResultsFile "  --- Saving: $updatedPath ---"

    try {
        $ext = $fileItem.Extension.ToLower()
        $xlFileFormat = switch ($ext) {
            ".xlsx" { 51 }   # xlOpenXMLWorkbook
            ".xlsm" { 52 }   # xlOpenXMLWorkbookMacroEnabled
            ".xls"  { 56 }   # xlExcel8
            default { 51 }
        }

        $workbook.SaveAs(
            $updatedPath,
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
        Write-Log $ResultsFile "  Saved successfully: $updatedPath"
    } catch {
        $msg = "  Could not save '$updatedPath'. Details: $($_.Exception.Message)"
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

    return @{ Counts = $counts; Errors = $errorList; Skipped = $false }
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
# Build Output folder - one level above the source folder
# ---------------------------------------------------------------------------
$parentDir = Split-Path $FolderPath -Parent
$outputDir = Join-Path $parentDir "Output"

Write-Host "Output directory: $outputDir"

if (-not (Test-Path -LiteralPath $outputDir)) {
    try {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Host "Created output directory: $outputDir"
    } catch {
        Write-Error "FATAL: Could not create output directory '$outputDir'. Details: $($_.Exception.Message)"
        exit 1
    }
}

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
    $excel = New-Object -ComObject Excel.Application
} catch {
    $msg = "FATAL: Could not create Excel COM object. Is Microsoft Excel installed? Details: $($_.Exception.Message)"
    Write-ErrorLog $errorFile $msg
    Write-Error $msg
    exit 1
}

$excel.Visible               = $false
$excel.DisplayAlerts         = $false
$excel.AskToUpdateLinks      = $false
$excel.AlertBeforeOverwriting = $false

Write-Log $resultsFile "Excel COM object ready."

# ---------------------------------------------------------------------------
# Initialise grand-total counters
# ---------------------------------------------------------------------------
$grandTotals      = New-FileCounts
$skippedFiles     = [System.Collections.Generic.List[string]]::new()
$allErrors        = [System.Collections.Generic.List[string]]::new()
$filesProcessed   = 0

# ---------------------------------------------------------------------------
# Process each file
# ---------------------------------------------------------------------------
$fileIndex = 0
foreach ($fileItem in $excelFiles) {
    $fileIndex++
    Write-Host ""
    Write-Host "[$fileIndex / $($excelFiles.Count)] Processing: $($fileItem.Name)"

    $result = Invoke-ProcessWorkbook `
        -Excel       $excel `
        -FilePath    $fileItem.FullName `
        -OutputDir   $outputDir `
        -ResultsFile $resultsFile `
        -ErrorFile   $errorFile

    if ($result.Skipped) {
        $skippedFiles.Add($fileItem.Name)
        Write-Log $resultsFile "  !! SKIPPED: $($fileItem.Name)"
    } else {
        $filesProcessed++
        # Accumulate grand totals
        foreach ($key in @($grandTotals.Keys)) {
            $grandTotals[$key] += $result.Counts[$key]
        }
    }

    foreach ($e in $result.Errors) { $allErrors.Add($e) }

    # Collect garbage between files to free COM memory
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
Add-Content -LiteralPath $resultsFile -Value "" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value ("=" * 60) -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  GRAND TOTAL SUMMARY" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Files in folder  : $($excelFiles.Count)" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Files processed  : $filesProcessed" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Files skipped    : $($skippedFiles.Count)  -> $($skippedFiles -join ', ')" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value ("=" * 60) -Encoding UTF8

$maxLen = ($grandTotals.Keys | Measure-Object -Property Length -Maximum).Maximum
foreach ($key in $grandTotals.Keys) {
    $pad  = " " * ($maxLen - $key.Length)
    $line = "  $key$pad : $($grandTotals[$key])"
    Add-Content -LiteralPath $resultsFile -Value $line -Encoding UTF8
    Write-Host $line
}

Add-Content -LiteralPath $resultsFile -Value "" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Output location : $outputDir" -Encoding UTF8
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
Write-Host "  Files processed : $filesProcessed / $($excelFiles.Count)"
Write-Host "  Files skipped   : $($skippedFiles.Count)"
Write-Host "  Output folder   : $outputDir"
Write-Host "  Results log     : $resultsFile"
Write-Host "  Error log       : $errorFile"
Write-Host "========================================================"
Write-Host ""
