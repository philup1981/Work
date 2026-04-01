<#
.SYNOPSIS
    Identifies filtered rows and columns in an Excel spreadsheet, reveals them,
    and copies the result to an Output directory. Reports counts per search term.

.DESCRIPTION
    This script:
      - Searches every worksheet in the specified Excel workbook.
      - Distinguishes rows hidden by AutoFilter ("filtered") from manually hidden rows.
      - Distinguishes columns hidden by outline/group collapse ("filtered") from manually hidden columns.
      - Only processes ENTIRE rows or columns that are filtered (AutoFilter / outline group).
      - Counts how many times each search term appears in those filtered rows/columns.
      - Reveals (unhides) the filtered rows/columns in a saved copy of the workbook.
      - Removes AutoFilter arrows so the content cannot be re-filtered accidentally.
      - Runs a QC pass to confirm no filtered content remains.
      - Writes Results.txt and Error.txt to the Output directory.

    Output directory is created one level above the spreadsheet's parent folder, named "Output".

.PARAMETER SpreadsheetPath
    Full or UNC path to the Excel file  (e.g. C:\Data\Report.xlsx  or  \\srv\share\Report.xlsx).

.PARAMETER SearchTermsFile
    Full or UNC path to a plain-text file containing search terms, one per line.

.EXAMPLE
    .\Search for filtered rows and columns.ps1 `
        -SpreadsheetPath "C:\Reports\Q1.xlsx" `
        -SearchTermsFile "C:\Reports\terms.txt"

.EXAMPLE
    .\Search for filtered rows and columns.ps1 `
        -SpreadsheetPath "\\fileserver\share\Q1.xlsx" `
        -SearchTermsFile "\\fileserver\share\terms.txt"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true,
               HelpMessage = "Full or UNC path to the Excel spreadsheet")]
    [string]$SpreadsheetPath,

    [Parameter(Mandatory = $true,
               HelpMessage = "Full or UNC path to the text file containing search terms (one per line)")]
    [string]$SearchTermsFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
# GLOBALS
# ─────────────────────────────────────────────────────────────────────────────
$script:ErrorLog   = [System.Collections.Generic.List[string]]::new()
$script:ExcelApp   = $null
$script:Workbook   = $null

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR','QC')]
        [string]$Level = 'INFO'
    )
    $ts     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $tag    = switch ($Level) {
        'INFO'    { '[INFO]   ' }
        'SUCCESS' { '[SUCCESS]' }
        'WARNING' { '[WARNING]' }
        'ERROR'   { '[ERROR]  ' }
        'QC'      { '[QC]     ' }
    }
    $line = "$ts $tag $Message"
    Write-Host $line

    if ($Level -in 'ERROR','WARNING') {
        $script:ErrorLog.Add($line)
    }
}

function Close-Excel {
    param([switch]$Save)
    try {
        if ($null -ne $script:Workbook) {
            if ($Save) { $script:Workbook.Save() }
            $script:Workbook.Close($false)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:Workbook)
            $script:Workbook = $null
        }
    } catch {
        Write-Log "Warning while closing workbook: $_" 'WARNING'
    }
    try {
        if ($null -ne $script:ExcelApp) {
            $script:ExcelApp.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:ExcelApp)
            $script:ExcelApp = $null
        }
    } catch {
        Write-Log "Warning while quitting Excel: $_" 'WARNING'
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Test-ReadAccess {
    param([string]$FilePath)
    try {
        $fs = [System.IO.File]::OpenRead($FilePath)
        $fs.Close()
        return $true
    } catch {
        return $false
    }
}

function Test-WriteAccess {
    param([string]$DirectoryPath)
    $probe = Join-Path $DirectoryPath "writeprobe_$(Get-Random).tmp"
    try {
        [System.IO.File]::WriteAllText($probe, 'test')
        Remove-Item $probe -Force
        return $true
    } catch {
        return $false
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BANNER
# ─────────────────────────────────────────────────────────────────────────────
Write-Log ('=' * 72)
Write-Log 'Search for Filtered Rows and Columns'
Write-Log ('=' * 72)
Write-Log "Run time         : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Log "Running as       : $env:USERDOMAIN\$env:USERNAME"
Write-Log "Spreadsheet Path : $SpreadsheetPath"
Write-Log "Search Terms File: $SearchTermsFile"
Write-Log ('=' * 72)

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — VALIDATE INPUTS
# ─────────────────────────────────────────────────────────────────────────────
Write-Log 'STEP 1: Validating input files...'

$abortEarly = $false

# --- Spreadsheet ---
if ([string]::IsNullOrWhiteSpace($SpreadsheetPath)) {
    Write-Log 'SpreadsheetPath parameter is empty.' 'ERROR'
    $abortEarly = $true
} elseif (-not (Test-Path -LiteralPath $SpreadsheetPath -PathType Leaf)) {
    Write-Log "Spreadsheet file not found or inaccessible: '$SpreadsheetPath'" 'ERROR'
    Write-Log "  Check that the file exists and the path is correct." 'ERROR'
    Write-Log "  For network (UNC) paths, verify the share is reachable and you have access." 'ERROR'
    $abortEarly = $true
} else {
    $xlItem = Get-Item -LiteralPath $SpreadsheetPath
    Write-Log "Spreadsheet found: '$SpreadsheetPath'  [$('{0:N2}' -f ($xlItem.Length / 1KB)) KB]"

    if (-not (Test-ReadAccess $SpreadsheetPath)) {
        Write-Log "Cannot open '$SpreadsheetPath' for reading.  The file may be locked or you lack permission." 'ERROR'
        $abortEarly = $true
    } else {
        Write-Log "Read access confirmed for spreadsheet."
    }

    $validExts = @('.xlsx','.xls','.xlsm','.xlsb','.xltx','.xltm')
    if ($xlItem.Extension.ToLower() -notin $validExts) {
        Write-Log "File extension '$($xlItem.Extension)' is not a recognised Excel format." 'WARNING'
        Write-Log "  Supported extensions: $($validExts -join ', ')" 'WARNING'
    }
}

# --- Search Terms File ---
if ([string]::IsNullOrWhiteSpace($SearchTermsFile)) {
    Write-Log 'SearchTermsFile parameter is empty.' 'ERROR'
    $abortEarly = $true
} elseif (-not (Test-Path -LiteralPath $SearchTermsFile -PathType Leaf)) {
    Write-Log "Search terms file not found or inaccessible: '$SearchTermsFile'" 'ERROR'
    Write-Log "  Check that the file exists and the path is correct." 'ERROR'
    Write-Log "  For network (UNC) paths, verify the share is reachable and you have access." 'ERROR'
    $abortEarly = $true
} else {
    Write-Log "Search terms file found: '$SearchTermsFile'"
    if (-not (Test-ReadAccess $SearchTermsFile)) {
        Write-Log "Cannot open '$SearchTermsFile' for reading.  The file may be locked or you lack permission." 'ERROR'
        $abortEarly = $true
    } else {
        Write-Log "Read access confirmed for search terms file."
    }
}

if ($abortEarly) {
    Write-Log 'Validation failed — aborting.  Resolve the errors above and re-run.' 'ERROR'
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — LOAD SEARCH TERMS
# ─────────────────────────────────────────────────────────────────────────────
Write-Log 'STEP 2: Loading search terms...'

try {
    $searchTerms = Get-Content -LiteralPath $SearchTermsFile -Encoding UTF8 |
                   Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                   ForEach-Object { $_.Trim() } |
                   Select-Object -Unique
} catch {
    Write-Log "Failed to read '$SearchTermsFile': $_" 'ERROR'
    exit 1
}

if ($searchTerms.Count -eq 0) {
    Write-Log "Search terms file '$SearchTermsFile' is empty or contains only blank lines." 'ERROR'
    exit 1
}

Write-Log "Loaded $($searchTerms.Count) unique search term(s):"
foreach ($t in $searchTerms) { Write-Log "  '$t'" }

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — SET UP OUTPUT DIRECTORY
# ─────────────────────────────────────────────────────────────────────────────
Write-Log 'STEP 3: Setting up output directory...'

try {
    $spreadsheetAbsPath = (Resolve-Path -LiteralPath $SpreadsheetPath).ProviderPath
} catch {
    Write-Log "Cannot resolve absolute path for '$SpreadsheetPath': $_" 'ERROR'
    exit 1
}

$spreadsheetDir = Split-Path -Path $spreadsheetAbsPath -Parent
$parentDir      = Split-Path -Path $spreadsheetDir     -Parent

# Guard: if the spreadsheet is already at a root (e.g. C:\file.xlsx), parentDir would be empty
if ([string]::IsNullOrWhiteSpace($parentDir)) {
    Write-Log "Cannot go one level above '$spreadsheetDir' — the spreadsheet appears to be at a root path." 'ERROR'
    exit 1
}

$outputDir = Join-Path $parentDir 'Output'

if (-not (Test-Path -LiteralPath $outputDir)) {
    try {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        Write-Log "Created output directory: $outputDir"
    } catch {
        Write-Log "Could not create output directory '$outputDir': $_" 'ERROR'
        Write-Log "  Ensure you have write permission to: $parentDir" 'ERROR'
        exit 1
    }
} else {
    Write-Log "Output directory already exists: $outputDir"
}

if (-not (Test-WriteAccess $outputDir)) {
    Write-Log "No write access to output directory '$outputDir'." 'ERROR'
    exit 1
}
Write-Log "Write access confirmed for output directory."

$xlFileName      = Split-Path $spreadsheetAbsPath -Leaf
$updatedFilePath = Join-Path $outputDir "Updated_$xlFileName"
$resultsFilePath = Join-Path $outputDir 'Results.txt'
$errorFilePath   = Join-Path $outputDir 'Error.txt'

Write-Log "Output files:"
Write-Log "  Updated spreadsheet : $updatedFilePath"
Write-Log "  Results             : $resultsFilePath"
Write-Log "  Error log           : $errorFilePath"

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — COPY SPREADSHEET
# ─────────────────────────────────────────────────────────────────────────────
Write-Log 'STEP 4: Copying spreadsheet to output directory...'
try {
    Copy-Item -LiteralPath $spreadsheetAbsPath -Destination $updatedFilePath -Force
    Write-Log "Spreadsheet copied to: $updatedFilePath" 'SUCCESS'
} catch {
    Write-Log "Failed to copy spreadsheet to '$updatedFilePath': $_" 'ERROR'
    Write-Log "  Source     : $spreadsheetAbsPath" 'ERROR'
    Write-Log "  Destination: $updatedFilePath" 'ERROR'
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 5 — OPEN EXCEL VIA COM AUTOMATION
# ─────────────────────────────────────────────────────────────────────────────
Write-Log 'STEP 5: Launching Excel via COM automation...'
try {
    $script:ExcelApp                = New-Object -ComObject Excel.Application
    $script:ExcelApp.Visible        = $false
    $script:ExcelApp.DisplayAlerts  = $false
    $script:ExcelApp.ScreenUpdating = $false
    $script:ExcelApp.EnableEvents   = $false
    Write-Log "Excel application launched."
} catch {
    Write-Log "Failed to launch Excel.  Ensure Microsoft Excel is installed on this machine." 'ERROR'
    Write-Log "  Detail: $_" 'ERROR'
    exit 1
}

Write-Log "Opening workbook: $updatedFilePath"
try {
    # Open(Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword,
    #      IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify,
    #      Converter, AddToMru)
    $script:Workbook = $script:ExcelApp.Workbooks.Open(
        $updatedFilePath,
        0,       # UpdateLinks  — do not update
        $false,  # ReadOnly     — we need to edit
        5,       # Format
        '',      # Password
        '',      # WriteResPassword
        $true,   # IgnoreReadOnlyRecommended
        2,       # Origin (Windows)
        ',',     # Delimiter
        $false,  # Editable
        $false,  # Notify
        0,       # Converter
        $false   # AddToMru
    )
    Write-Log "Workbook opened successfully."
} catch {
    Write-Log "Failed to open workbook '$updatedFilePath': $_" 'ERROR'
    Close-Excel
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 6 — PROCESS EVERY WORKSHEET
# ─────────────────────────────────────────────────────────────────────────────
Write-Log ('=' * 72)
Write-Log 'STEP 6: Processing worksheets...'
Write-Log ('=' * 72)

# Initialise per-term counters
$termCounts = [ordered]@{}
foreach ($t in $searchTerms) { $termCounts[$t] = 0 }

$totalFilteredRows = 0
$totalFilteredCols = 0
$sheetsProcessed   = 0
$sheetCount        = $script:Workbook.Worksheets.Count

for ($si = 1; $si -le $sheetCount; $si++) {

    $ws     = $script:Workbook.Worksheets.Item($si)
    $wsName = $ws.Name

    Write-Log "--- Sheet $si / $sheetCount : '$wsName' ---"

    try {
        # ── Determine used range ──────────────────────────────────────────────
        $used = $ws.UsedRange
        if ($null -eq $used -or $used.Count -eq 0) {
            Write-Log "  '$wsName': No used range — skipping." 'WARNING'
            continue
        }

        $rowFirst = $used.Row
        $rowLast  = $used.Row + $used.Rows.Count    - 1
        $colFirst = $used.Column
        $colLast  = $used.Column + $used.Columns.Count - 1

        Write-Log "  '$wsName': Used range — rows $rowFirst–$rowLast, cols $colFirst–$colLast"

        # ── ROWS — identify filtered vs manually hidden ───────────────────────
        #
        #   AutoFilter rows (FilterMode = True):
        #     Step A: record every hidden row in the used range.
        #     Step B: call ShowAllData() — this reveals only AutoFilter-hidden rows.
        #     Step C: rows that are STILL hidden = manually hidden.
        #             rows that are NOW visible  = were AutoFilter-filtered.
        #
        #   No AutoFilter (FilterMode = False):
        #     Any hidden rows are manually hidden — ignored.
        #
        $filteredRows = [System.Collections.Generic.HashSet[int]]::new()
        $manualRows   = [System.Collections.Generic.HashSet[int]]::new()

        $hasAutoFilter = $ws.AutoFilterMode   # arrows present (may or may not be filtering)
        $isFiltered    = $ws.FilterMode       # criteria are currently applied → rows are hidden

        Write-Log "  '$wsName': AutoFilterMode=$hasAutoFilter  FilterMode=$isFiltered"

        if ($isFiltered) {
            Write-Log "  '$wsName': Detecting AutoFilter-hidden rows..."

            # A — snapshot hidden rows
            $hiddenBefore = [System.Collections.Generic.HashSet[int]]::new()
            for ($r = $rowFirst; $r -le $rowLast; $r++) {
                if ($ws.Rows.Item($r).Hidden) { [void]$hiddenBefore.Add($r) }
            }
            Write-Log "  '$wsName': Hidden rows before ShowAllData: $($hiddenBefore.Count)"

            # B — remove filter criteria (reveals AutoFilter-hidden rows)
            try {
                $ws.ShowAllData()
                Write-Log "  '$wsName': ShowAllData() executed — AutoFilter criteria cleared."
            } catch {
                Write-Log "  '$wsName': ShowAllData() failed: $_" 'WARNING'
            }

            # C — snapshot still-hidden rows
            $hiddenAfter = [System.Collections.Generic.HashSet[int]]::new()
            for ($r = $rowFirst; $r -le $rowLast; $r++) {
                if ($ws.Rows.Item($r).Hidden) { [void]$hiddenAfter.Add($r) }
            }
            Write-Log "  '$wsName': Hidden rows after ShowAllData (manually hidden): $($hiddenAfter.Count)"

            foreach ($r in $hiddenBefore) {
                if ($hiddenAfter.Contains($r)) { [void]$manualRows.Add($r)   }
                else                           { [void]$filteredRows.Add($r) }
            }

        } else {
            # No active filter — any hidden rows are manual
            for ($r = $rowFirst; $r -le $rowLast; $r++) {
                if ($ws.Rows.Item($r).Hidden) { [void]$manualRows.Add($r) }
            }
        }

        Write-Log "  '$wsName': AutoFilter-filtered rows : $($filteredRows.Count)"
        Write-Log "  '$wsName': Manually hidden rows     : $($manualRows.Count)"

        if ($manualRows.Count -gt 0) {
            Write-Log "  '$wsName': Manually hidden rows are EXCLUDED from processing (as requested)." 'INFO'
        }

        # ── COLUMNS — filtered (outline-grouped) vs manually hidden ──────────
        #
        #   Excel has no column AutoFilter.  The closest "filtered" equivalent is
        #   columns collapsed by an outline group (OutlineLevel > 1 when grouped).
        #   Columns with OutlineLevel = 1 (ungrouped) that are hidden = manually hidden.
        #
        $filteredCols = [System.Collections.Generic.HashSet[int]]::new()
        $manualCols   = [System.Collections.Generic.HashSet[int]]::new()

        for ($c = $colFirst; $c -le $colLast; $c++) {
            $col = $ws.Columns.Item($c)
            if ($col.Hidden) {
                if ($col.OutlineLevel -gt 1) {
                    # Hidden as part of a collapsed outline group — treated as "filtered"
                    [void]$filteredCols.Add($c)
                } else {
                    # Plain manually hidden column
                    [void]$manualCols.Add($c)
                }
            }
        }

        Write-Log "  '$wsName': Outline-grouped filtered columns : $($filteredCols.Count)"
        Write-Log "  '$wsName': Manually hidden columns           : $($manualCols.Count)"

        if ($manualCols.Count -gt 0) {
            Write-Log "  '$wsName': Manually hidden columns are EXCLUDED from processing (as requested)." 'INFO'
        }

        # ── Nothing to do? ────────────────────────────────────────────────────
        if ($filteredRows.Count -eq 0 -and $filteredCols.Count -eq 0) {
            Write-Log "  '$wsName': No AutoFilter-filtered rows or outline-filtered columns found — skipping." 'INFO'
            # Still remove AutoFilter arrows if present to avoid confusion
            if ($hasAutoFilter) {
                try { $ws.AutoFilterMode = $false } catch { }
            }
            continue
        }

        # ── Search filtered cells for each term ───────────────────────────────
        Write-Log "  '$wsName': Scanning filtered rows/columns for search terms..."

        $sheetCounts = [ordered]@{}
        foreach ($t in $searchTerms) { $sheetCounts[$t] = 0 }

        $hitCells = [System.Collections.Generic.List[string]]::new()

        for ($r = $rowFirst; $r -le $rowLast; $r++) {
            $inFilteredRow = $filteredRows.Contains($r)

            for ($c = $colFirst; $c -le $colLast; $c++) {
                $inFilteredCol = $filteredCols.Contains($c)

                # Only examine cells in a filtered row OR a filtered column
                # (i.e. the ENTIRE row/column is filtered — not just the cell)
                if (-not $inFilteredRow -and -not $inFilteredCol) { continue }

                # Skip cells whose column is merely manually hidden
                if ($manualCols.Contains($c)) { continue }

                $cell      = $ws.Cells.Item($r, $c)
                $rawVal    = $cell.Value2
                $cellText  = if ($null -ne $rawVal) { [string]$rawVal } else { '' }

                if ([string]::IsNullOrEmpty($cellText)) { continue }

                foreach ($t in $searchTerms) {
                    # Case-insensitive literal match
                    if ($cellText -match [regex]::Escape($t)) {
                        $sheetCounts[$t]++
                        $termCounts[$t]++
                        $addr = $cell.Address($false, $false)
                        $msg  = "    HIT  Term='$t'  Sheet='$wsName'  Cell=$addr  Value='$cellText'"
                        Write-Log $msg
                        $hitCells.Add($msg)
                    }
                }
            }
        }

        Write-Log "  '$wsName': Per-term counts this sheet:"
        foreach ($t in $searchTerms) {
            Write-Log ("    {0,-45} : {1}" -f "'$t'", $sheetCounts[$t])
        }

        # ── Ensure filtered rows are fully visible in the output ──────────────
        if ($filteredRows.Count -gt 0) {
            Write-Log "  '$wsName': Confirming filtered rows are unhidden in output file..."
            $stillHidden = 0
            foreach ($r in $filteredRows) {
                if ($ws.Rows.Item($r).Hidden) {
                    $ws.Rows.Item($r).Hidden = $false
                    $stillHidden++
                }
            }
            if ($stillHidden -gt 0) {
                Write-Log "  '$wsName': Force-unhid $stillHidden row(s) that were still hidden after ShowAllData." 'WARNING'
            }
            $totalFilteredRows += $filteredRows.Count
            Write-Log "  '$wsName': $($filteredRows.Count) filtered row(s) revealed." 'SUCCESS'
        }

        # ── Unhide filtered (outline-grouped) columns ─────────────────────────
        if ($filteredCols.Count -gt 0) {
            Write-Log "  '$wsName': Unhiding $($filteredCols.Count) outline-filtered column(s)..."
            foreach ($c in ($filteredCols | Sort-Object)) {
                try {
                    $ws.Columns.Item($c).Hidden = $false
                    $colLetter = $script:ExcelApp.ActiveWorkbook.Sheets.Item($si).Columns.Item($c).Address($false,$false).Split(':')[0] -replace '\d',''
                    Write-Log "    Unhid column $c ($colLetter)"
                } catch {
                    Write-Log "    Could not unhide column $c : $_" 'WARNING'
                }
            }
            $totalFilteredCols += $filteredCols.Count
            Write-Log "  '$wsName': $($filteredCols.Count) filtered column(s) revealed." 'SUCCESS'
        }

        # ── Remove AutoFilter arrows from this sheet ──────────────────────────
        if ($hasAutoFilter) {
            try {
                $ws.AutoFilterMode = $false
                Write-Log "  '$wsName': AutoFilter arrows removed to prevent re-filtering."
            } catch {
                Write-Log "  '$wsName': Could not remove AutoFilter arrows: $_" 'WARNING'
            }
        }

        $sheetsProcessed++
        Write-Log "  '$wsName': Done." 'SUCCESS'

    } catch {
        Write-Log "Unexpected error on sheet '$wsName': $_" 'ERROR'
        Write-Log "  Script stack: $($_.ScriptStackTrace)" 'ERROR'
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 7 — SAVE WORKBOOK
# ─────────────────────────────────────────────────────────────────────────────
Write-Log ('=' * 72)
Write-Log 'STEP 7: Saving updated workbook...'
try {
    $script:Workbook.Save()
    Write-Log "Workbook saved: $updatedFilePath" 'SUCCESS'
} catch {
    Write-Log "Failed to save workbook: $_" 'ERROR'
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 8 — QC CHECK
# ─────────────────────────────────────────────────────────────────────────────
Write-Log ('=' * 72)
Write-Log 'STEP 8: Running QC — checking for any remaining filtered content...'
Write-Log ('=' * 72)

$qcNotes = [System.Collections.Generic.List[string]]::new()

for ($si = 1; $si -le $script:Workbook.Worksheets.Count; $si++) {
    $ws     = $script:Workbook.Worksheets.Item($si)
    $wsName = $ws.Name

    try {
        # Remaining AutoFilter with criteria
        if ($ws.FilterMode) {
            $msg = "QC ISSUE — Sheet '$wsName': AutoFilter criteria still active (rows may still be hidden by filter)."
            Write-Log $msg 'QC'
            $qcNotes.Add($msg)
        }

        # AutoFilter arrows still present (non-blocking but noted)
        if ($ws.AutoFilterMode) {
            $msg = "QC NOTE  — Sheet '$wsName': AutoFilter arrows still present (no criteria active)."
            Write-Log $msg 'QC'
            $qcNotes.Add($msg)
        }

        $used = $ws.UsedRange
        if ($null -ne $used -and $used.Count -gt 0) {
            $rFirst = $used.Row
            $rLast  = $used.Row + $used.Rows.Count    - 1
            $cFirst = $used.Column
            $cLast  = $used.Column + $used.Columns.Count - 1

            $hiddenRowsLeft = 0
            for ($r = $rFirst; $r -le $rLast; $r++) {
                if ($ws.Rows.Item($r).Hidden) { $hiddenRowsLeft++ }
            }
            $hiddenColsLeft = 0
            for ($c = $cFirst; $c -le $cLast; $c++) {
                if ($ws.Columns.Item($c).Hidden) { $hiddenColsLeft++ }
            }

            if ($hiddenRowsLeft -gt 0) {
                $msg = "QC NOTE  — Sheet '$wsName': $hiddenRowsLeft hidden row(s) remain (expected: manually hidden — not filtered)."
                Write-Log $msg 'QC'
                $qcNotes.Add($msg)
            } else {
                Write-Log "QC OK    — Sheet '$wsName': No hidden rows remain." 'QC'
            }

            if ($hiddenColsLeft -gt 0) {
                $msg = "QC NOTE  — Sheet '$wsName': $hiddenColsLeft hidden column(s) remain (expected: manually hidden — not filtered)."
                Write-Log $msg 'QC'
                $qcNotes.Add($msg)
            } else {
                Write-Log "QC OK    — Sheet '$wsName': No hidden columns remain." 'QC'
            }
        }

    } catch {
        Write-Log "QC error on sheet '$wsName': $_" 'WARNING'
    }
}

if ($qcNotes.Count -eq 0) {
    Write-Log 'QC PASSED — No remaining filtered content detected.' 'SUCCESS'
} else {
    Write-Log "QC COMPLETED WITH $($qcNotes.Count) NOTE(S) — See QC lines above." 'WARNING'
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 9 — CLOSE EXCEL
# ─────────────────────────────────────────────────────────────────────────────
Write-Log 'STEP 9: Closing Excel...'
Close-Excel
Write-Log 'Excel closed.'

# ─────────────────────────────────────────────────────────────────────────────
# STEP 10 — WRITE Results.txt
# ─────────────────────────────────────────────────────────────────────────────
Write-Log "STEP 10: Writing $resultsFilePath..."

$sb = [System.Text.StringBuilder]::new()
$div60 = '=' * 72
$sep40 = '-' * 40

[void]$sb.AppendLine($div60)
[void]$sb.AppendLine('Search for Filtered Rows and Columns — Results')
[void]$sb.AppendLine($div60)
[void]$sb.AppendLine("Generated        : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$sb.AppendLine("Run as           : $env:USERDOMAIN\$env:USERNAME")
[void]$sb.AppendLine("Source workbook  : $spreadsheetAbsPath")
[void]$sb.AppendLine("Output workbook  : $updatedFilePath")
[void]$sb.AppendLine("Search terms file: $SearchTermsFile")
[void]$sb.AppendLine($div60)
[void]$sb.AppendLine('')
[void]$sb.AppendLine('PROCESSING SUMMARY')
[void]$sb.AppendLine($sep40)
[void]$sb.AppendLine("Sheets processed                   : $sheetCount total, $sheetsProcessed contained filtered content")
[void]$sb.AppendLine("Filtered rows revealed             : $totalFilteredRows")
[void]$sb.AppendLine("Filtered (outline-grouped) columns : $totalFilteredCols")
[void]$sb.AppendLine('')
[void]$sb.AppendLine('SEARCH TERM COUNTS')
[void]$sb.AppendLine('(occurrences found in AutoFilter-filtered rows or outline-grouped columns)')
[void]$sb.AppendLine($sep40)

foreach ($t in $searchTerms) {
    $line = "{0,-55} : {1}" -f "'$t'", $termCounts[$t]
    [void]$sb.AppendLine($line)
}

[void]$sb.AppendLine('')
[void]$sb.AppendLine('QC CHECK')
[void]$sb.AppendLine($sep40)
if ($qcNotes.Count -eq 0) {
    [void]$sb.AppendLine('PASSED — No remaining filtered content detected in the output workbook.')
} else {
    [void]$sb.AppendLine("COMPLETED WITH $($qcNotes.Count) NOTE(S):")
    foreach ($n in $qcNotes) { [void]$sb.AppendLine("  $n") }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Note: "Manually hidden" rows/columns are NOT filtered — they were hidden')
    [void]$sb.AppendLine('manually before or independently of any AutoFilter/outline operation.')
    [void]$sb.AppendLine('These are left unchanged in the output file, as per the script requirements.')
}

try {
    [System.IO.File]::WriteAllText($resultsFilePath, $sb.ToString(), [System.Text.Encoding]::UTF8)
    Write-Log "Results written to: $resultsFilePath" 'SUCCESS'
} catch {
    Write-Log "Failed to write results file '$resultsFilePath': $_" 'ERROR'
}

# ─────────────────────────────────────────────────────────────────────────────
# STEP 11 — WRITE Error.txt
# ─────────────────────────────────────────────────────────────────────────────
Write-Log "STEP 11: Writing $errorFilePath..."

$esSb = [System.Text.StringBuilder]::new()
[void]$esSb.AppendLine($div60)
[void]$esSb.AppendLine('Search for Filtered Rows and Columns — Error / Warning Log')
[void]$esSb.AppendLine($div60)
[void]$esSb.AppendLine("Generated        : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
[void]$esSb.AppendLine("Run as           : $env:USERDOMAIN\$env:USERNAME")
[void]$esSb.AppendLine("Source workbook  : $spreadsheetAbsPath")
[void]$esSb.AppendLine($div60)
[void]$esSb.AppendLine('')

if ($script:ErrorLog.Count -eq 0) {
    [void]$esSb.AppendLine('No errors or warnings were recorded during this run.')
} else {
    [void]$esSb.AppendLine("$($script:ErrorLog.Count) error/warning(s) recorded:")
    [void]$esSb.AppendLine('')
    foreach ($e in $script:ErrorLog) { [void]$esSb.AppendLine($e) }
}

try {
    [System.IO.File]::WriteAllText($errorFilePath, $esSb.ToString(), [System.Text.Encoding]::UTF8)
    Write-Log "Error log written to: $errorFilePath" 'SUCCESS'
} catch {
    Write-Log "Could not write error log '$errorFilePath': $_" 'WARNING'
}

# ─────────────────────────────────────────────────────────────────────────────
# FINAL SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
Write-Log ('=' * 72)
Write-Log 'COMPLETE' 'SUCCESS'
Write-Log ('=' * 72)
Write-Log "Updated spreadsheet : $updatedFilePath"
Write-Log "Results             : $resultsFilePath"
Write-Log "Error log           : $errorFilePath"
Write-Log ''
Write-Log 'Search Term Counts (filtered rows/columns only):'
foreach ($t in $searchTerms) {
    Write-Log ("  {0,-50} : {1}" -f "'$t'", $termCounts[$t])
}
Write-Log ('=' * 72)
