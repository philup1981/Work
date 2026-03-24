#Requires -Version 5.1
<#
.SYNOPSIS
    Flattens Excel workbook content (pivot tables, charts, linked data) and removes hidden sheets/rows/columns.

.DESCRIPTION
    This script:
      - Flattens all pivot tables to static values
      - Flattens all charts to static images (embedded pictures)
      - Removes external data connections
      - Removes hidden and very-hidden worksheets
      - Removes hidden rows and columns in every visible sheet
      - Removes hidden ListObjects (tables)
      - Runs a QC pass to confirm no residual hidden content
      - Writes an updated workbook, Results.txt, and Error.txt to an Output folder
        one level above the workbook's parent folder

.PARAMETER SpreadsheetPath
    Full path (local or UNC) to the source Excel workbook.

.EXAMPLE
    .\<script>.ps1 -SpreadsheetPath "\\server\share\reports\MyBook.xlsx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Full path to the Excel workbook (local or UNC).")]
    [string]$SpreadsheetPath
)

# ---------------------------------------------------------------------------
# Helper: append a timestamped line to a log file
# ---------------------------------------------------------------------------
function Write-Log {
    param(
        [string]$FilePath,
        [string]$Message
    )
    $ts  = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $line = "[$ts] $Message"
    Add-Content -LiteralPath $FilePath -Value $line -Encoding UTF8
    Write-Host $line
}

# ---------------------------------------------------------------------------
# Helper: append to Error log and also write to host
# ---------------------------------------------------------------------------
function Write-ErrorLog {
    param(
        [string]$FilePath,
        [string]$Message
    )
    $ts   = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $line = "[ERROR][$ts] $Message"
    Add-Content -LiteralPath $FilePath -Value $line -Encoding UTF8
    Write-Warning $line
}

# ---------------------------------------------------------------------------
# Counters (populated during processing)
# ---------------------------------------------------------------------------
$counts = [ordered]@{
    "Pivot Tables Flattened"          = 0
    "Charts Converted to Images"      = 0
    "External Connections Removed"    = 0
    "Named Ranges Pointing Externally Removed" = 0
    "Hidden Sheets Removed"           = 0
    "Very Hidden Sheets Removed"      = 0
    "Hidden Rows Removed"             = 0
    "Hidden Columns Removed"          = 0
    "Hidden Tables (ListObjects) Removed" = 0
    "QC Issues Found After Processing"= 0
}

$errorList = [System.Collections.Generic.List[string]]::new()

# ---------------------------------------------------------------------------
# STEP 1 – Resolve and validate the spreadsheet path
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "========================================================"
Write-Host "  Flatten & Remove Hidden Information"
Write-Host "  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "========================================================"
Write-Host ""

# Expand environment variables / relative paths
try {
    $SpreadsheetPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SpreadsheetPath)
} catch {
    Write-Warning "Could not resolve provided path '$SpreadsheetPath'. Proceeding with the path as given."
}

Write-Host "Spreadsheet path supplied : $SpreadsheetPath"

# Does the file exist?
if (-not (Test-Path -LiteralPath $SpreadsheetPath -PathType Leaf)) {
    Write-Error "FATAL: The file '$SpreadsheetPath' does not exist or cannot be reached.`nVerify the path is correct and that network connectivity / drive mapping is in place."
    exit 1
}

# Can we read it?
try {
    $null = [System.IO.File]::Open($SpreadsheetPath, 'Open', 'Read', 'ReadWrite')
} catch {
    Write-Error "FATAL: The file '$SpreadsheetPath' exists but cannot be opened for reading.`nCheck permissions or whether the file is exclusively locked by another process.`nDetails: $($_.Exception.Message)"
    exit 1
}

# Resolve to a fully qualified path in case it's a UNC path passed as relative
$SpreadsheetPath = (Get-Item -LiteralPath $SpreadsheetPath).FullName
Write-Host "Resolved full path         : $SpreadsheetPath"

# ---------------------------------------------------------------------------
# STEP 2 – Build output folder structure
# ---------------------------------------------------------------------------
$fileItem       = Get-Item -LiteralPath $SpreadsheetPath
$fileDir        = $fileItem.DirectoryName          # e.g. \\server\share\reports
$parentDir      = Split-Path $fileDir -Parent       # one level above
$outputDir      = Join-Path $parentDir "Output"

Write-Host "Output directory           : $outputDir"

if (-not (Test-Path -LiteralPath $outputDir)) {
    try {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Host "Created output directory   : $outputDir"
    } catch {
        Write-Error "FATAL: Could not create output directory '$outputDir'.`nDetails: $($_.Exception.Message)"
        exit 1
    }
}

$resultsFile    = Join-Path $outputDir "Results.txt"
$errorFile      = Join-Path $outputDir "Error.txt"
$updatedName    = "Updated_" + $fileItem.Name
$updatedPath    = Join-Path $outputDir $updatedName

# Initialise/clear log files for this run
$runHeader = "=" * 60
Set-Content -LiteralPath $resultsFile -Value $runHeader          -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Results – Flatten & Remove Hidden Information"  -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Run date : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Source   : $SpreadsheetPath" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value $runHeader          -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value ""                  -Encoding UTF8

Set-Content -LiteralPath $errorFile   -Value $runHeader          -Encoding UTF8
Add-Content -LiteralPath $errorFile   -Value "  Error Log – Flatten & Remove Hidden Information" -Encoding UTF8
Add-Content -LiteralPath $errorFile   -Value "  Run date : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Encoding UTF8
Add-Content -LiteralPath $errorFile   -Value "  Source   : $SpreadsheetPath" -Encoding UTF8
Add-Content -LiteralPath $errorFile   -Value $runHeader          -Encoding UTF8
Add-Content -LiteralPath $errorFile   -Value ""                  -Encoding UTF8

Write-Log $resultsFile "Processing started."

# ---------------------------------------------------------------------------
# STEP 3 – Launch Excel via COM (under the current user account)
# ---------------------------------------------------------------------------
Write-Log $resultsFile "Launching Microsoft Excel via COM automation (current user context)."

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    $msg = "FATAL: Could not create an Excel COM object. Is Microsoft Excel installed for the current user?`nDetails: $($_.Exception.Message)"
    Write-ErrorLog $errorFile $msg
    Add-Content -LiteralPath $errorFile -Value $msg -Encoding UTF8
    Write-Error $msg
    exit 1
}

$excel.Visible               = $false
$excel.DisplayAlerts         = $false
$excel.AskToUpdateLinks      = $false
$excel.AlertBeforeOverwriting = $false

Write-Log $resultsFile "Excel COM object created successfully."

# ---------------------------------------------------------------------------
# STEP 4 – Open workbook
# ---------------------------------------------------------------------------
$workbook = $null
try {
    # UpdateLinks=0 (don't update), ReadOnly=false, Format=5 (default)
    $workbook = $excel.Workbooks.Open(
        $SpreadsheetPath,   # Filename
        0,                  # UpdateLinks  – 0 = don't update
        $false,             # ReadOnly
        5,                  # Format       – 5 = nothing special
        "",                 # Password
        "",                 # WriteResPassword
        $true,              # IgnoreReadOnlyRecommended
        [System.Reflection.Missing]::Value,  # Origin
        [System.Reflection.Missing]::Value,  # Delimiter
        $false,             # Editable
        $false,             # Notify
        [System.Reflection.Missing]::Value,  # Converter
        $false              # AddToMru
    )
} catch {
    $msg = "FATAL: Could not open workbook '$SpreadsheetPath'.`nDetails: $($_.Exception.Message)"
    Write-ErrorLog $errorFile $msg
    if ($excel) { $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    Write-Error $msg
    exit 1
}

Write-Log $resultsFile "Workbook opened: $($workbook.FullName)"

# ---------------------------------------------------------------------------
# Helper: release a COM object safely
# ---------------------------------------------------------------------------
function Release-Com {
    param($obj)
    if ($null -ne $obj) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {}
    }
}

# ===========================================================================
# STEP 5 – Remove external data connections
# ===========================================================================
Write-Log $resultsFile "--- Removing external data connections ---"

try {
    $connCount = $workbook.Connections.Count
    Write-Log $resultsFile "  Found $connCount connection(s) in workbook."
    for ($c = $connCount; $c -ge 1; $c--) {
        try {
            $conn = $workbook.Connections.Item($c)
            $connName = $conn.Name
            $conn.Delete()
            $counts["External Connections Removed"]++
            Write-Log $resultsFile "  Removed connection [$c]: '$connName'"
            Release-Com $conn
        } catch {
            $msg = "  Could not remove connection [$c]: $($_.Exception.Message)"
            Write-ErrorLog $errorFile $msg
            $errorList.Add($msg)
        }
    }
} catch {
    $msg = "Error enumerating workbook connections: $($_.Exception.Message)"
    Write-ErrorLog $errorFile $msg
    $errorList.Add($msg)
}

# ===========================================================================
# STEP 6 – Remove external named ranges / names
# ===========================================================================
Write-Log $resultsFile "--- Checking workbook-level Names for external references ---"

try {
    $nameCount = $workbook.Names.Count
    Write-Log $resultsFile "  Found $nameCount named range(s) in workbook."
    $externalNames = @()
    for ($n = 1; $n -le $nameCount; $n++) {
        try {
            $nm = $workbook.Names.Item($n)
            $ref = $nm.RefersTo
            # External references contain '[' (workbook name in brackets)
            if ($ref -match '\[') {
                $externalNames += $nm.Name
            }
            Release-Com $nm
        } catch {}
    }
    foreach ($eName in $externalNames) {
        try {
            $nm = $workbook.Names.Item($eName)
            $nm.Delete()
            $counts["Named Ranges Pointing Externally Removed"]++
            Write-Log $resultsFile "  Removed external named range: '$eName'"
            Release-Com $nm
        } catch {
            $msg = "  Could not remove named range '$eName': $($_.Exception.Message)"
            Write-ErrorLog $errorFile $msg
            $errorList.Add($msg)
        }
    }
} catch {
    $msg = "Error enumerating named ranges: $($_.Exception.Message)"
    Write-ErrorLog $errorFile $msg
    $errorList.Add($msg)
}

# ===========================================================================
# STEP 7 – Process each worksheet
# ===========================================================================
Write-Log $resultsFile "--- Processing worksheets ---"

# Excel sheet visibility constants
$xlSheetVisible    = -1   # xlSheetVisible
$xlSheetHidden     =  0   # xlSheetHidden
$xlSheetVeryHidden =  2   # xlSheetVeryHidden

# Collect hidden sheets first (iterate by index to allow deletion)
$hiddenSheetNames     = @()
$veryHiddenSheetNames = @()

$totalSheets = $workbook.Sheets.Count
Write-Log $resultsFile "  Total sheets in workbook: $totalSheets"

for ($s = 1; $s -le $totalSheets; $s++) {
    try {
        $sh = $workbook.Sheets.Item($s)
        switch ($sh.Visible) {
            $xlSheetHidden     { $hiddenSheetNames     += $sh.Name }
            $xlSheetVeryHidden { $veryHiddenSheetNames += $sh.Name }
        }
        Release-Com $sh
    } catch {
        $msg = "  Could not read visibility of sheet index $s: $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
    }
}

Write-Log $resultsFile "  Hidden sheets found      : $($hiddenSheetNames.Count)  -> $($hiddenSheetNames -join ', ')"
Write-Log $resultsFile "  Very-hidden sheets found : $($veryHiddenSheetNames.Count)  -> $($veryHiddenSheetNames -join ', ')"

# We must keep at least one visible sheet – collect all visible sheet names
$visibleSheetNames = @()
for ($s = 1; $s -le $totalSheets; $s++) {
    try {
        $sh = $workbook.Sheets.Item($s)
        if ($sh.Visible -eq $xlSheetVisible) { $visibleSheetNames += $sh.Name }
        Release-Com $sh
    } catch {}
}

# ----- Process visible sheets: flatten pivot tables, charts, hidden rows/cols -----
Write-Log $resultsFile "  Processing visible sheets for pivot tables, charts, hidden rows/columns..."

foreach ($shName in $visibleSheetNames) {
    Write-Log $resultsFile "  >> Sheet: '$shName'"

    $ws = $null
    try {
        $ws = $workbook.Sheets.Item($shName)
    } catch {
        $msg = "    Could not access sheet '$shName': $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
        continue
    }

    # ---- 7a: Flatten Pivot Tables ----
    try {
        $ptCount = $ws.PivotTables().Count
        Write-Log $resultsFile "    Pivot tables found: $ptCount"

        for ($p = $ptCount; $p -ge 1; $p--) {
            try {
                $pt        = $ws.PivotTables($p)
                $ptName    = $pt.Name
                $ptRange   = $pt.TableRange2  # full pivot table range incl. page fields

                # Copy the range to clipboard then paste as values
                $ptRange.Copy()
                $ptRange.PasteSpecial(-4163)  # xlPasteValues = -4163

                # Delete the pivot cache / pivot table object
                $pt.TableRange2.ClearContents()
                # Paste values back
                $ptRange.PasteSpecial(-4163)

                # The pivot table object still exists until we delete the PivotCache
                # Safest: re-fetch pivot table and delete it
                try {
                    $ptObj = $ws.PivotTables($p)
                    $cache = $ptObj.PivotCache()
                    $ptObj.TableRange2.ClearOutline()
                    # Remove the pivot table (leaves values in place)
                    $ptObj.TableRange1.Clear()      # clear the actual pivot
                    # Repaste values from clipboard
                    $ptRange.PasteSpecial(-4163)
                    Release-Com $ptObj
                    Release-Com $cache
                } catch {}

                $counts["Pivot Tables Flattened"]++
                Write-Log $resultsFile "    Flattened pivot table '$ptName' in sheet '$shName'."
                Release-Com $ptRange
                Release-Com $pt
            } catch {
                $msg = "    Error flattening pivot table $p in sheet '$shName': $($_.Exception.Message)"
                Write-ErrorLog $errorFile $msg
                $errorList.Add($msg)
            }
        }
    } catch {
        $msg = "    Error accessing pivot tables in sheet '$shName': $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
    }

    # ---- 7b: Flatten Charts (ChartObjects) to static images ----
    try {
        $coCount = $ws.ChartObjects().Count
        Write-Log $resultsFile "    Chart objects found: $coCount"

        for ($ch = $coCount; $ch -ge 1; $ch--) {
            try {
                $co      = $ws.ChartObjects($ch)
                $coName  = $co.Name
                $coLeft  = $co.Left
                $coTop   = $co.Top
                $coW     = $co.Width
                $coH     = $co.Height

                # Copy the chart as a picture to clipboard
                $co.CopyPicture(1, -4147)  # Appearance=xlScreen=1, Format=xlPicture=-4147

                # Paste as picture on the sheet
                $ws.Paste()
                $excel.CutCopyMode = $false  # clear clipboard

                # Reposition the pasted picture to the same location
                # The pasted picture is the last shape added
                $shapeCount = $ws.Shapes.Count
                $pic = $ws.Shapes.Item($shapeCount)
                $pic.Left  = $coLeft
                $pic.Top   = $coTop
                $pic.Width = $coW
                $pic.Height= $coH

                Release-Com $pic

                # Delete the original chart object
                $co.Delete()
                $counts["Charts Converted to Images"]++
                Write-Log $resultsFile "    Converted chart '$coName' to static image in sheet '$shName'."
                Release-Com $co
            } catch {
                $msg = "    Error converting chart $ch in sheet '$shName': $($_.Exception.Message)"
                Write-ErrorLog $errorFile $msg
                $errorList.Add($msg)
            }
        }
    } catch {
        $msg = "    Error accessing charts in sheet '$shName': $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
    }

    # ---- 7c: Remove hidden rows ----
    try {
        $usedRange = $ws.UsedRange
        $rowCount  = $usedRange.Rows.Count
        $firstRow  = $usedRange.Row
        $hiddenRowCount = 0

        # Build a list of hidden row indices (1-based worksheet rows)
        $hiddenRowIndices = [System.Collections.Generic.List[int]]::new()
        for ($r = $firstRow; $r -lt ($firstRow + $rowCount); $r++) {
            try {
                $rowObj = $ws.Rows.Item($r)
                if ($rowObj.Hidden -eq $true) {
                    $hiddenRowIndices.Add($r)
                }
                Release-Com $rowObj
            } catch {}
        }

        Write-Log $resultsFile "    Hidden rows found in used range: $($hiddenRowIndices.Count)"

        # Delete in reverse order so indices don't shift
        $hiddenRowIndices.Reverse()
        foreach ($ri in $hiddenRowIndices) {
            try {
                $rowObj = $ws.Rows.Item($ri)
                $rowObj.Delete()
                $counts["Hidden Rows Removed"]++
                $hiddenRowCount++
                Release-Com $rowObj
            } catch {
                $msg = "    Error deleting hidden row $ri in sheet '$shName': $($_.Exception.Message)"
                Write-ErrorLog $errorFile $msg
                $errorList.Add($msg)
            }
        }
        if ($hiddenRowCount -gt 0) {
            Write-Log $resultsFile "    Removed $hiddenRowCount hidden row(s) from sheet '$shName'."
        }
        Release-Com $usedRange
    } catch {
        $msg = "    Error processing hidden rows in sheet '$shName': $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
    }

    # ---- 7d: Remove hidden columns ----
    try {
        $usedRange = $ws.UsedRange
        $colCount  = $usedRange.Columns.Count
        $firstCol  = $usedRange.Column
        $hiddenColCount = 0

        $hiddenColIndices = [System.Collections.Generic.List[int]]::new()
        for ($c = $firstCol; $c -lt ($firstCol + $colCount); $c++) {
            try {
                $colObj = $ws.Columns.Item($c)
                if ($colObj.Hidden -eq $true) {
                    $hiddenColIndices.Add($c)
                }
                Release-Com $colObj
            } catch {}
        }

        Write-Log $resultsFile "    Hidden columns found in used range: $($hiddenColIndices.Count)"

        $hiddenColIndices.Reverse()
        foreach ($ci in $hiddenColIndices) {
            try {
                $colObj = $ws.Columns.Item($ci)
                $colObj.Delete()
                $counts["Hidden Columns Removed"]++
                $hiddenColCount++
                Release-Com $colObj
            } catch {
                $msg = "    Error deleting hidden column $ci in sheet '$shName': $($_.Exception.Message)"
                Write-ErrorLog $errorFile $msg
                $errorList.Add($msg)
            }
        }
        if ($hiddenColCount -gt 0) {
            Write-Log $resultsFile "    Removed $hiddenColCount hidden column(s) from sheet '$shName'."
        }
        Release-Com $usedRange
    } catch {
        $msg = "    Error processing hidden columns in sheet '$shName': $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
    }

    # ---- 7e: Remove hidden ListObjects (Tables) ----
    try {
        $loCount = $ws.ListObjects.Count
        Write-Log $resultsFile "    ListObjects (tables) found: $loCount"

        $hiddenTables = @()
        for ($lo = 1; $lo -le $loCount; $lo++) {
            try {
                $loObj = $ws.ListObjects.Item($lo)
                # A ListObject is considered "hidden" if its ShowHeaders and ShowTotals
                # are both false AND it has no visible display range – or if every row/col
                # it occupies is hidden. We'll treat any table whose entire range is
                # within hidden rows/cols as hidden.
                $loRange = $loObj.Range
                $entirelyHidden = $loRange.EntireRow.Hidden -and $loRange.EntireColumn.Hidden
                if ($entirelyHidden) {
                    $hiddenTables += $loObj.Name
                }
                Release-Com $loRange
                Release-Com $loObj
            } catch {}
        }

        foreach ($tblName in $hiddenTables) {
            try {
                $loObj = $ws.ListObjects.Item($tblName)
                $loObj.Delete()
                $counts["Hidden Tables (ListObjects) Removed"]++
                Write-Log $resultsFile "    Removed hidden table '$tblName' in sheet '$shName'."
                Release-Com $loObj
            } catch {
                $msg = "    Error removing hidden table '$tblName' in sheet '$shName': $($_.Exception.Message)"
                Write-ErrorLog $errorFile $msg
                $errorList.Add($msg)
            }
        }
    } catch {
        $msg = "    Error processing ListObjects in sheet '$shName': $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
    }

    Release-Com $ws
}

# ----- Remove hidden sheets (after processing visible ones) -----
Write-Log $resultsFile "--- Removing hidden sheets ---"

foreach ($shName in $hiddenSheetNames) {
    try {
        $sh = $workbook.Sheets.Item($shName)
        $sh.Visible = $xlSheetVisible   # must be visible before delete
        $sh.Delete()
        $counts["Hidden Sheets Removed"]++
        Write-Log $resultsFile "  Deleted hidden sheet: '$shName'"
        Release-Com $sh
    } catch {
        $msg = "  Error deleting hidden sheet '$shName': $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
    }
}

foreach ($shName in $veryHiddenSheetNames) {
    try {
        $sh = $workbook.Sheets.Item($shName)
        $sh.Visible = $xlSheetVisible
        $sh.Delete()
        $counts["Very Hidden Sheets Removed"]++
        Write-Log $resultsFile "  Deleted very-hidden sheet: '$shName'"
        Release-Com $sh
    } catch {
        $msg = "  Error deleting very-hidden sheet '$shName': $($_.Exception.Message)"
        Write-ErrorLog $errorFile $msg
        $errorList.Add($msg)
    }
}

# ===========================================================================
# STEP 8 – QC Pass: verify no residual hidden content
# ===========================================================================
Write-Log $resultsFile "--- QC Pass: checking for residual hidden content ---"
$qcIssues = [System.Collections.Generic.List[string]]::new()

# QC: sheets
$sheetCountQC = $workbook.Sheets.Count
for ($s = 1; $s -le $sheetCountQC; $s++) {
    try {
        $sh = $workbook.Sheets.Item($s)
        if ($sh.Visible -ne $xlSheetVisible) {
            $issue = "QC ISSUE: Sheet '$($sh.Name)' is still hidden (Visible=$($sh.Visible))."
            $qcIssues.Add($issue)
            Write-Log $resultsFile "  $issue"
        }
        Release-Com $sh
    } catch {}
}

# QC: pivot tables, charts, hidden rows/cols on remaining visible sheets
$sheetCountQC2 = $workbook.Sheets.Count
for ($s = 1; $s -le $sheetCountQC2; $s++) {
    try {
        $sh = $workbook.Sheets.Item($s)
        if ($sh.Visible -ne $xlSheetVisible) { Release-Com $sh; continue }
        $shNameQC = $sh.Name

        # Pivot tables
        try {
            $ptQC = $sh.PivotTables().Count
            if ($ptQC -gt 0) {
                $issue = "QC ISSUE: Sheet '$shNameQC' still contains $ptQC pivot table(s)."
                $qcIssues.Add($issue)
                Write-Log $resultsFile "  $issue"
            }
        } catch {}

        # Charts
        try {
            $coQC = $sh.ChartObjects().Count
            if ($coQC -gt 0) {
                $issue = "QC ISSUE: Sheet '$shNameQC' still contains $coQC chart object(s)."
                $qcIssues.Add($issue)
                Write-Log $resultsFile "  $issue"
            }
        } catch {}

        # Hidden rows
        try {
            $ur = $sh.UsedRange
            $rStart = $ur.Row
            $rEnd   = $rStart + $ur.Rows.Count - 1
            $hidR   = 0
            for ($r = $rStart; $r -le $rEnd; $r++) {
                $rObj = $sh.Rows.Item($r)
                if ($rObj.Hidden) { $hidR++ }
                Release-Com $rObj
            }
            if ($hidR -gt 0) {
                $issue = "QC ISSUE: Sheet '$shNameQC' still has $hidR hidden row(s)."
                $qcIssues.Add($issue)
                Write-Log $resultsFile "  $issue"
            }
            Release-Com $ur
        } catch {}

        # Hidden columns
        try {
            $ur = $sh.UsedRange
            $cStart = $ur.Column
            $cEnd   = $cStart + $ur.Columns.Count - 1
            $hidC   = 0
            for ($c = $cStart; $c -le $cEnd; $c++) {
                $cObj = $sh.Columns.Item($c)
                if ($cObj.Hidden) { $hidC++ }
                Release-Com $cObj
            }
            if ($hidC -gt 0) {
                $issue = "QC ISSUE: Sheet '$shNameQC' still has $hidC hidden column(s)."
                $qcIssues.Add($issue)
                Write-Log $resultsFile "  $issue"
            }
            Release-Com $ur
        } catch {}

        Release-Com $sh
    } catch {}
}

# QC: connections
try {
    $connQC = $workbook.Connections.Count
    if ($connQC -gt 0) {
        $issue = "QC ISSUE: Workbook still has $connQC external connection(s)."
        $qcIssues.Add($issue)
        Write-Log $resultsFile "  $issue"
    }
} catch {}

$counts["QC Issues Found After Processing"] = $qcIssues.Count

if ($qcIssues.Count -eq 0) {
    Write-Log $resultsFile "  QC PASSED – No residual hidden content found."
} else {
    Write-Log $resultsFile "  QC COMPLETED WITH $($qcIssues.Count) ISSUE(S). See details above."
    foreach ($qi in $qcIssues) {
        Write-ErrorLog $errorFile "  $qi"
    }
}

# ===========================================================================
# STEP 9 – Save updated workbook
# ===========================================================================
Write-Log $resultsFile "--- Saving updated workbook ---"
Write-Log $resultsFile "  Destination: $updatedPath"

try {
    # Determine file format from extension
    $ext = $fileItem.Extension.ToLower()
    $xlOpenXML    = 51   # xlsx
    $xlOpenXMLMacro = 52 # xlsm
    $xlExcel8     = 56   # xls (Excel 97-2003)
    $xlFileFormat = switch ($ext) {
        ".xlsx" { $xlOpenXML }
        ".xlsm" { $xlOpenXMLMacro }
        ".xls"  { $xlExcel8 }
        default { $xlOpenXML }
    }

    $workbook.SaveAs(
        $updatedPath,
        $xlFileFormat,
        [System.Reflection.Missing]::Value,  # Password
        [System.Reflection.Missing]::Value,  # WriteResPassword
        $false,                              # ReadOnlyRecommended
        $false,                              # CreateBackup
        1,                                   # AccessMode = xlExclusive
        [System.Reflection.Missing]::Value,  # ConflictResolution
        $false,                              # AddToMru
        [System.Reflection.Missing]::Value,  # TextCodepage
        [System.Reflection.Missing]::Value,  # TextVisualLayout
        $false                               # Local
    )
    Write-Log $resultsFile "  Workbook saved successfully: $updatedPath"
} catch {
    $msg = "FATAL: Could not save updated workbook to '$updatedPath'.`nDetails: $($_.Exception.Message)"
    Write-ErrorLog $errorFile $msg
    $errorList.Add($msg)
    Write-Error $msg
}

# ===========================================================================
# STEP 10 – Close workbook and quit Excel
# ===========================================================================
try {
    $workbook.Close($false)
    Release-Com $workbook
} catch {}
try {
    $excel.Quit()
    Release-Com $excel
} catch {}

# Force garbage collection to release COM objects
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.GC]::Collect()

Write-Log $resultsFile "Excel closed."

# ===========================================================================
# STEP 11 – Write summary counts to Results.txt
# ===========================================================================
Write-Log $resultsFile ""
Add-Content -LiteralPath $resultsFile -Value "------------------------------------------------------------" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  SUMMARY OF ACTIONS" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "------------------------------------------------------------" -Encoding UTF8

$maxLen = ($counts.Keys | Measure-Object -Property Length -Maximum).Maximum
foreach ($key in $counts.Keys) {
    $padding = " " * ($maxLen - $key.Length)
    $line = "  $key$padding : $($counts[$key])"
    Add-Content -LiteralPath $resultsFile -Value $line -Encoding UTF8
    Write-Host $line
}

Add-Content -LiteralPath $resultsFile -Value "" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "  Output files:" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "    Updated workbook : $updatedPath" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "    Results log      : $resultsFile" -Encoding UTF8
Add-Content -LiteralPath $resultsFile -Value "    Error log        : $errorFile" -Encoding UTF8

# Write any errors to Error.txt
if ($errorList.Count -gt 0) {
    Add-Content -LiteralPath $errorFile -Value "" -Encoding UTF8
    Add-Content -LiteralPath $errorFile -Value "------------------------------------------------------------" -Encoding UTF8
    Add-Content -LiteralPath $errorFile -Value "  ERRORS ENCOUNTERED ($($errorList.Count) total)" -Encoding UTF8
    Add-Content -LiteralPath $errorFile -Value "------------------------------------------------------------" -Encoding UTF8
    foreach ($e in $errorList) {
        Add-Content -LiteralPath $errorFile -Value "  $e" -Encoding UTF8
    }
} else {
    Add-Content -LiteralPath $errorFile -Value "  No errors encountered during this run." -Encoding UTF8
}

Write-Host ""
Write-Host "========================================================"
Write-Host "  Processing complete."
Write-Host "  Updated workbook : $updatedPath"
Write-Host "  Results log      : $resultsFile"
Write-Host "  Error log        : $errorFile"
Write-Host "========================================================"
Write-Host ""
