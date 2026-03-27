#Requires -Version 5.1
<#
.SYNOPSIS
    QC Check: Detects hidden content, external links, formulas, and pivot tables
    in Microsoft spreadsheet files.

.DESCRIPTION
    Scans all Microsoft spreadsheet files in a specified folder and inspects each
    workbook for the following flagged content:

      - Hidden worksheets       (xlSheetHidden)
      - Very Hidden worksheets  (xlSheetVeryHidden - only visible/settable via VBA/COM)
      - Hidden rows             (within each sheet's used range)
      - Hidden columns          (within each sheet's used range)
      - External links          (linked workbooks / external data sources)
      - Formulas                (any cell-level formula on any sheet)
      - Pivot tables            (on any worksheet)

    Each file is copied (not moved) to one of three output sub-directories:

      Has content            - One or more flagged items were detected
      Does not have content  - Inspection completed; no flagged items found
      Review                 - File could not be opened, errored, or result undetermined

    Two reports are written to the Output parent directory:
      logs.txt    - Verbose timestamped processing log for every file and check
      Errors.txt  - List of files that failed processing with reason

    NOTE: Files are NEVER modified. All workbooks are opened read-only with
          alerts and screen-updating suppressed. No Excel window will be shown.

.PARAMETER FolderPath
    Path to the folder containing the spreadsheet files to inspect.
    Supports local paths (e.g. C:\Data) and UNC network paths
    (e.g. \\fileserver\share\documents).

.EXAMPLE
    .\Detect-SpreadsheetHiddenContent.ps1 -FolderPath "C:\Data\Reports"

.EXAMPLE
    .\Detect-SpreadsheetHiddenContent.ps1 -FolderPath "\\fileserver\share\financials"

.NOTES
    Requirements : Microsoft Excel must be installed on the machine running this script.
    Modules      : None (no external modules required).
    Permissions  : The running user must have Read access to the source folder and
                   Write access to the parent directory (for the Output folder).
#>

param(
    [Parameter(Mandatory = $true,
               HelpMessage = "Path to the folder containing spreadsheet files to inspect")]
    [string]$FolderPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"   # Non-terminating errors are logged; script continues

# =============================================================================
# EXCEL COM CONSTANTS
# =============================================================================
$XL_SHEET_VISIBLE     = -1   # xlSheetVisible
$XL_SHEET_HIDDEN      =  0   # xlSheetHidden
$XL_SHEET_VERY_HIDDEN =  2   # xlSheetVeryHidden
$XL_EXCEL_LINKS       =  1   # xlExcelLinks  (LinkSources type)
$XL_CELL_FORMULAS     = -4123 # xlCellTypeFormulas (SpecialCells type)
$XL_WORKSHEET         = -4167 # xlWorksheet (XlSheetType enum)

# Supported Microsoft spreadsheet file extensions
$SUPPORTED_EXTENSIONS = @(
    '.xlsx',  # Excel Workbook
    '.xlsm',  # Excel Macro-Enabled Workbook
    '.xlsb',  # Excel Binary Workbook
    '.xls',   # Excel 97-2003 Workbook
    '.xltx',  # Excel Template
    '.xltm',  # Excel Macro-Enabled Template
    '.xlt',   # Excel 97-2003 Template
    '.xlam',  # Excel Add-In
    '.xla'    # Excel 97-2003 Add-In
)

# =============================================================================
# PATH SETUP
# =============================================================================

# Normalize: strip trailing separators
$FolderPath = $FolderPath.TrimEnd('\', '/')

# Validate source folder exists before doing anything else
if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
    Write-Error "FATAL: The specified folder does not exist or cannot be accessed."
    Write-Error "       Path supplied : '$FolderPath'"
    Write-Error "       Verify the path is correct and the current user has Read permissions."
    exit 1
}

# Resolve to an absolute provider path (handles relative paths and normalises UNC)
$resolvedFolder = (Resolve-Path -LiteralPath $FolderPath).ProviderPath

# Output tree  ->  [parent of scanned folder]\Output\...
$parentDir     = Split-Path -Parent $resolvedFolder
$outputDir     = Join-Path $parentDir "Output"
$dirHasContent = Join-Path $outputDir "Has content"
$dirNoContent  = Join-Path $outputDir "Does not have content"
$dirReview     = Join-Path $outputDir "Review"
$pathLogs      = Join-Path $outputDir "logs.txt"
$pathErrors    = Join-Path $outputDir "Errors.txt"

# Create all output directories (fail early if we cannot write)
foreach ($dir in @($outputDir, $dirHasContent, $dirNoContent, $dirReview)) {
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "FATAL: Cannot create output directory '$dir'."
            Write-Error "       Error: $($_.Exception.Message)"
            Write-Error "       Verify the current user has Write access to '$parentDir'."
            exit 1
        }
    }
}

# =============================================================================
# LOGGING ENGINE
# =============================================================================

$script:LogBuffer  = [System.Text.StringBuilder]::new(65536)
$script:ErrorList  = [System.Collections.Generic.List[string]]::new()
$script:StartTime  = Get-Date

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG','RESULT','SECTION')]
        [string]$Level = 'INFO'
    )

    $ts    = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $entry = "[$ts] [$($Level.PadRight(7))] $Message"

    [void]$script:LogBuffer.AppendLine($entry)

    $color = switch ($Level) {
        'ERROR'   { 'Red'      }
        'WARN'    { 'Yellow'   }
        'RESULT'  { 'Cyan'     }
        'DEBUG'   { 'DarkGray' }
        'SECTION' { 'Magenta'  }
        default   { 'White'    }
    }
    Write-Host $entry -ForegroundColor $color
}

function Write-LogSection {
    param([string]$Title = '')
    $bar = '-' * 80
    [void]$script:LogBuffer.AppendLine('')
    [void]$script:LogBuffer.AppendLine($bar)
    Write-Host ''
    Write-Host $bar -ForegroundColor Magenta
    if ($Title) {
        [void]$script:LogBuffer.AppendLine("  $Title")
        [void]$script:LogBuffer.AppendLine($bar)
        Write-Host "  $Title" -ForegroundColor Magenta
        Write-Host $bar      -ForegroundColor Magenta
    }
}

function Save-OutputFiles {
    # Write verbose log
    try {
        $script:LogBuffer.ToString() | Out-File -FilePath $pathLogs -Encoding UTF8 -Force
    } catch {
        Write-Host "WARNING: Could not write log file '$pathLogs': $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Write error/review report
    try {
        $sb = [System.Text.StringBuilder]::new()
        [void]$sb.AppendLine("SPREADSHEET QC CHECK - ERRORS AND REVIEW FILES")
        [void]$sb.AppendLine("Generated  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$sb.AppendLine("Source     : $resolvedFolder")
        [void]$sb.AppendLine("=" * 70)
        [void]$sb.AppendLine('')

        if ($script:ErrorList.Count -gt 0) {
            [void]$sb.AppendLine("Total entries : $($script:ErrorList.Count)")
            [void]$sb.AppendLine('')
            $n = 1
            foreach ($entry in $script:ErrorList) {
                [void]$sb.AppendLine("  [$n] $entry")
                $n++
            }
        } else {
            [void]$sb.AppendLine("No errors recorded. All files were processed without issue.")
        }

        $sb.ToString() | Out-File -FilePath $pathErrors -Encoding UTF8 -Force
    } catch {
        Write-Host "WARNING: Could not write errors file '$pathErrors': $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# =============================================================================
# SAFE FILE COPY  (copies; never moves; resolves destination name conflicts)
# =============================================================================

function Copy-ToOutput {
    param(
        [string]$SourcePath,
        [string]$DestDir,
        [string]$OriginalName
    )

    $destPath = Join-Path $DestDir $OriginalName

    # If a file with the same name already exists, append a numeric suffix
    if (Test-Path -LiteralPath $destPath) {
        $base    = [System.IO.Path]::GetFileNameWithoutExtension($OriginalName)
        $ext     = [System.IO.Path]::GetExtension($OriginalName)
        $counter = 1
        do {
            $newName  = "${base}_copy${counter}${ext}"
            $destPath = Join-Path $DestDir $newName
            $counter++
        } while (Test-Path -LiteralPath $destPath)
        Write-Log "  Destination name conflict resolved. File will be saved as '$newName'" -Level WARN
    }

    Copy-Item -LiteralPath $SourcePath -Destination $destPath -Force -ErrorAction Stop
    return $destPath
}

# =============================================================================
# EXCEL COM LIFECYCLE
# =============================================================================

function New-ExcelInstance {
    $xl = New-Object -ComObject Excel.Application -ErrorAction Stop
    $xl.Visible          = $false   # No visible window
    $xl.DisplayAlerts    = $false   # Suppress all alert dialogs
    $xl.ScreenUpdating   = $false   # No screen redraws
    $xl.EnableEvents     = $false   # No event macros fire
    $xl.Interactive      = $false   # Prevent user interaction
    $xl.AskToUpdateLinks = $false   # Never prompt to update links
    return $xl
}

function Remove-ExcelInstance {
    param($xl)
    if ($null -eq $xl) { return }
    try   { $xl.Quit() }          catch { }
    try   { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Release-ComObject {
    param($obj)
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

# =============================================================================
# CORE WORKBOOK INSPECTION
# =============================================================================

function Invoke-WorkbookInspection {
    param(
        [string]$FilePath,
        $XlApp
    )

    # Structured result object
    $r = [PSCustomObject]@{
        FileName            = [System.IO.Path]::GetFileName($FilePath)
        FilePath            = $FilePath
        Inspected           = $false          # True only if inspection completed without fatal error

        # Hidden / Very-Hidden sheets
        HasHiddenSheets     = $false
        HiddenSheets        = [System.Collections.Generic.List[string]]::new()
        HasVeryHiddenSheets = $false
        VeryHiddenSheets    = [System.Collections.Generic.List[string]]::new()

        # Hidden rows / columns
        HasHiddenRows       = $false
        HiddenRowDetails    = [System.Collections.Generic.List[string]]::new()
        HasHiddenColumns    = $false
        HiddenColDetails    = [System.Collections.Generic.List[string]]::new()

        # External links
        HasExternalLinks    = $false
        ExternalLinks       = [System.Collections.Generic.List[string]]::new()

        # Formulas
        HasFormulas         = $false
        FormulaSheets       = [System.Collections.Generic.List[string]]::new()

        # Pivot tables
        HasPivotTables      = $false
        PivotDetails        = [System.Collections.Generic.List[string]]::new()

        ErrorMessage        = $null
    }

    $wb = $null

    try {
        Write-Log "  Opening workbook read-only (link-updates and alerts suppressed)..." -Level DEBUG

        # Open read-only; UpdateLinks = 0 (never); suppress password prompts by passing empty strings
        $wb = $XlApp.Workbooks.Open(
            $FilePath,   # Filename
            0,           # UpdateLinks  = 0 (do not update any links)
            $true        # ReadOnly     = true
        )

        $sheetCount = $wb.Sheets.Count
        Write-Log "  Workbook opened successfully. Sheets: $sheetCount  |  Excel version in use: $($XlApp.Version)" -Level DEBUG

        # -----------------------------------------------------------------
        # CHECK 1 : External links (linked workbooks / OLE / DDE)
        # -----------------------------------------------------------------
        Write-Log "  [CHECK 1/5] External links..." -Level DEBUG
        try {
            $links = $wb.LinkSources($XL_EXCEL_LINKS)
            if ($null -ne $links) {
                $r.HasExternalLinks = $true
                foreach ($lnk in $links) {
                    $r.ExternalLinks.Add([string]$lnk)
                    Write-Log "    EXTERNAL LINK FOUND : $lnk" -Level WARN
                }
                Write-Log "  External links detected : $($r.ExternalLinks.Count) link(s)" -Level WARN
            } else {
                Write-Log "  No external links detected." -Level DEBUG
            }
        } catch {
            Write-Log "  Could not check external links (non-fatal, continuing) : $($_.Exception.Message)" -Level WARN
        }

        # -----------------------------------------------------------------
        # CHECK 2 : Workbook-level pivot caches
        # PivotCaches() is the most reliable pivot-table indicator — every
        # pivot table, on any sheet, registers a cache at workbook level.
        # -----------------------------------------------------------------
        Write-Log "  [CHECK 2] Workbook-level pivot caches..." -Level DEBUG
        try {
            $pivotCaches = $wb.PivotCaches()
            if ($null -ne $pivotCaches -and $pivotCaches.Count -gt 0) {
                $r.HasPivotTables = $true
                Write-Log "  Workbook has $($pivotCaches.Count) pivot cache(s) - pivot table(s) confirmed at workbook level" -Level WARN
            } else {
                Write-Log "  No pivot caches found at workbook level." -Level DEBUG
            }
            Release-ComObject $pivotCaches
        } catch {
            Write-Log "  Could not check workbook pivot caches (non-fatal, will check per-sheet) : $($_.Exception.Message)" -Level WARN
        }

        # -----------------------------------------------------------------
        # PER-SHEET INSPECTION
        # -----------------------------------------------------------------
        Write-Log "  [CHECK 3-6] Inspecting $sheetCount sheet(s) for hidden status, pivot tables, formulas, rows/columns..." -Level DEBUG

        foreach ($sheet in $wb.Sheets) {
            $sName = $sheet.Name
            $sVis  = $sheet.Visible

            # --- Sheet visibility -------------------------------------------
            if ($sVis -eq $XL_SHEET_VERY_HIDDEN) {
                $r.HasVeryHiddenSheets = $true
                $r.VeryHiddenSheets.Add($sName)
                Write-Log "    [SHEET VISIBILITY] '$sName'  >>>  VERY HIDDEN (xlSheetVeryHidden = $XL_SHEET_VERY_HIDDEN)" -Level WARN

            } elseif ($sVis -eq $XL_SHEET_HIDDEN) {
                $r.HasHiddenSheets = $true
                $r.HiddenSheets.Add($sName)
                Write-Log "    [SHEET VISIBILITY] '$sName'  >>>  HIDDEN (xlSheetHidden = $XL_SHEET_HIDDEN)" -Level WARN

            } else {
                Write-Log "    [SHEET VISIBILITY] '$sName'  -  Visible (Visibility value: $sVis)" -Level DEBUG
            }

            # Only standard worksheets have rows, columns, formulas, and pivot tables
            $sheetType = $null
            try { $sheetType = $sheet.Type } catch { }

            if ($sheetType -ne $XL_WORKSHEET) {
                Write-Log "    [SHEET TYPE] '$sName' is not a standard worksheet (Type=$sheetType) - skipping cell-level checks" -Level DEBUG
                Release-ComObject $sheet
                continue
            }

            # --- CHECK 3 : Pivot tables (per-sheet detail) ------------------
            Write-Log "    [CHECK 3] '$sName' - pivot tables (per-sheet)..." -Level DEBUG
            try {
                $ptCount = $sheet.PivotTables().Count
                if ($ptCount -gt 0) {
                    $r.HasPivotTables = $true
                    $r.PivotDetails.Add("Sheet '$sName' : $ptCount pivot table(s)")
                    Write-Log "    [PIVOT TABLES] '$sName'  >>>  $ptCount pivot table(s) found" -Level WARN
                } else {
                    Write-Log "    [PIVOT TABLES] '$sName'  -  none" -Level DEBUG
                }
            } catch {
                Write-Log "    [PIVOT TABLES] '$sName'  -  check failed (non-fatal) : $($_.Exception.Message)" -Level WARN
            }

            # --- Obtain the used range (scope all remaining checks to it) ---
            $usedRange = $null
            try { $usedRange = $sheet.UsedRange } catch { }

            if ($null -eq $usedRange) {
                Write-Log "    '$sName' has no used range - skipping formulas/rows/columns checks" -Level DEBUG
                Release-ComObject $sheet
                continue
            }

            $usedRows = $usedRange.Rows.Count
            $usedCols = $usedRange.Columns.Count
            Write-Log "    Used range: $usedRows row(s) x $usedCols column(s)" -Level DEBUG

            # --- CHECK 4 : Formulas -----------------------------------------
            Write-Log "    [CHECK 4] '$sName' - formulas..." -Level DEBUG
            try {
                # SpecialCells throws a COMException (not a PowerShell error) when no
                # matching cells exist - that is the expected "no formulas" signal.
                $fCells = $usedRange.SpecialCells($XL_CELL_FORMULAS)
                if ($null -ne $fCells) {
                    $r.HasFormulas = $true
                    $r.FormulaSheets.Add($sName)
                    Write-Log "    [FORMULAS] '$sName'  >>>  formulas detected in used range" -Level WARN
                }
                Release-ComObject $fCells
            } catch [System.Runtime.InteropServices.COMException] {
                # Normal - no formula cells present
                Write-Log "    [FORMULAS] '$sName'  -  none" -Level DEBUG
            } catch {
                Write-Log "    [FORMULAS] '$sName'  -  check failed (non-fatal) : $($_.Exception.Message)" -Level WARN
            }

            # --- CHECK 5 : Hidden rows (within used range) ------------------
            if ($usedRows -gt 0) {
                Write-Log "    [CHECK 5] '$sName' - scanning $usedRows row(s) for hidden status..." -Level DEBUG
                $hiddenRowCount = 0
                try {
                    for ($ri = 1; $ri -le $usedRows; $ri++) {
                        $rowObj = $usedRange.Rows.Item($ri)
                        if ($rowObj.Hidden) { $hiddenRowCount++ }
                        Release-ComObject $rowObj
                    }
                } catch {
                    Write-Log "    [ROWS] '$sName' - scan interrupted at row index $ri : $($_.Exception.Message)" -Level WARN
                }

                if ($hiddenRowCount -gt 0) {
                    $r.HasHiddenRows = $true
                    $r.HiddenRowDetails.Add("Sheet '$sName' : $hiddenRowCount hidden row(s) within used range ($usedRows total)")
                    Write-Log "    [HIDDEN ROWS] '$sName'  >>>  $hiddenRowCount hidden row(s) found" -Level WARN
                } else {
                    Write-Log "    [HIDDEN ROWS] '$sName'  -  none" -Level DEBUG
                }
            }

            # --- CHECK 6 : Hidden columns (within used range) ---------------
            if ($usedCols -gt 0) {
                Write-Log "    [CHECK 6] '$sName' - scanning $usedCols column(s) for hidden status..." -Level DEBUG
                $hiddenColCount = 0
                try {
                    for ($ci = 1; $ci -le $usedCols; $ci++) {
                        $colObj = $usedRange.Columns.Item($ci)
                        if ($colObj.Hidden) { $hiddenColCount++ }
                        Release-ComObject $colObj
                    }
                } catch {
                    Write-Log "    [COLUMNS] '$sName' - scan interrupted at column index $ci : $($_.Exception.Message)" -Level WARN
                }

                if ($hiddenColCount -gt 0) {
                    $r.HasHiddenColumns = $true
                    $r.HiddenColDetails.Add("Sheet '$sName' : $hiddenColCount hidden column(s) within used range ($usedCols total)")
                    Write-Log "    [HIDDEN COLS] '$sName'  >>>  $hiddenColCount hidden column(s) found" -Level WARN
                } else {
                    Write-Log "    [HIDDEN COLS] '$sName'  -  none" -Level DEBUG
                }
            }

            Release-ComObject $usedRange
            Release-ComObject $sheet

        } # end foreach sheet

        $r.Inspected = $true

    } catch {
        $r.ErrorMessage = $_.Exception.Message
        Write-Log "  INSPECTION ERROR : $($_.Exception.Message)" -Level ERROR
        Write-Log "  Script stack     : $($_.ScriptStackTrace)" -Level DEBUG

    } finally {
        # Always close the workbook without saving, regardless of outcome
        if ($null -ne $wb) {
            try {
                $wb.Close($false)         # $false = do not save changes
                Release-ComObject $wb
            } catch {
                Write-Log "  Warning: Problem releasing workbook COM object : $($_.Exception.Message)" -Level WARN
            }
        }
    }

    return $r
}

# =============================================================================
# MAIN SCRIPT BODY
# =============================================================================

Write-LogSection "SPREADSHEET QC CHECK - HIDDEN CONTENT DETECTOR"
Write-Log "Script version   : 1.0" -Level INFO
Write-Log "Start time       : $($script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level INFO
Write-Log "Running as user  : $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)" -Level INFO
Write-Log "PowerShell ver   : $($PSVersionTable.PSVersion)" -Level INFO
Write-Log "OS               : $([System.Environment]::OSVersion.VersionString)" -Level INFO
Write-Log "Machine name     : $env:COMPUTERNAME" -Level INFO
Write-Log "Source folder    : $resolvedFolder" -Level INFO
Write-Log "Parent directory : $parentDir" -Level INFO
Write-Log "Output directory : $outputDir" -Level INFO
Write-Log "  Has content    : $dirHasContent" -Level INFO
Write-Log "  No content     : $dirNoContent" -Level INFO
Write-Log "  Review         : $dirReview" -Level INFO
Write-Log "Log file         : $pathLogs" -Level INFO
Write-Log "Error file       : $pathErrors" -Level INFO
Write-Log "Supported formats: $($SUPPORTED_EXTENSIONS -join ', ')" -Level INFO

# =============================================================================
# ACCESS VERIFICATION
# =============================================================================

Write-LogSection "ACCESS VERIFICATION"

# --- Read access on source folder ---
Write-Log "Verifying READ access to source folder: $resolvedFolder" -Level INFO
try {
    $null = Get-ChildItem -LiteralPath $resolvedFolder -Force -ErrorAction Stop
    Write-Log "Read access CONFIRMED for source folder." -Level INFO
} catch {
    Write-Log "FATAL: Cannot read source folder '$resolvedFolder'." -Level ERROR
    Write-Log "       Error: $($_.Exception.Message)" -Level ERROR
    Write-Log "       Ensure the current user ($([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)) has Read permissions." -Level ERROR
    Save-OutputFiles
    exit 1
}

# --- Write access on output directory ---
Write-Log "Verifying WRITE access to output directory: $outputDir" -Level INFO
try {
    $probeFile = Join-Path $outputDir ".write_probe_$(Get-Random).tmp"
    [System.IO.File]::WriteAllText($probeFile, 'probe')
    Remove-Item -LiteralPath $probeFile -Force
    Write-Log "Write access CONFIRMED for output directory." -Level INFO
} catch {
    Write-Log "FATAL: Cannot write to output directory '$outputDir'." -Level ERROR
    Write-Log "       Error: $($_.Exception.Message)" -Level ERROR
    Write-Log "       Ensure the current user has Write permissions on '$parentDir'." -Level ERROR
    Save-OutputFiles
    exit 1
}

# =============================================================================
# FILE DISCOVERY
# =============================================================================

Write-LogSection "FILE DISCOVERY"
Write-Log "Enumerating spreadsheet files in: $resolvedFolder" -Level INFO

$spreadsheetFiles = @(
    Get-ChildItem -LiteralPath $resolvedFolder -File -Force |
    Where-Object { $SUPPORTED_EXTENSIONS -contains $_.Extension.ToLower() }
)

$totalFiles = $spreadsheetFiles.Count
Write-Log "Total spreadsheet files found: $totalFiles" -Level INFO

if ($totalFiles -eq 0) {
    Write-Log "No spreadsheet files found matching extensions: $($SUPPORTED_EXTENSIONS -join ', ')" -Level WARN
    Write-Log "Verify the folder path is correct and files have supported extensions." -Level WARN
    Save-OutputFiles
    exit 0
}

# List each discovered file
$idx = 0
foreach ($f in $spreadsheetFiles) {
    $idx++
    $szStr = if   ($f.Length -ge 1MB) { "$([math]::Round($f.Length / 1MB, 2)) MB" }
             elseif ($f.Length -ge 1KB) { "$([math]::Round($f.Length / 1KB, 2)) KB" }
             else                       { "$($f.Length) bytes" }
    Write-Log "  [$idx/$totalFiles]  $($f.Name)  ($szStr)  Last modified: $($f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level INFO
}

# =============================================================================
# INITIALISE EXCEL COM
# =============================================================================

Write-LogSection "INITIALISING EXCEL COM"
$xlApp = $null

try {
    Write-Log "Creating hidden Excel COM application instance..." -Level INFO
    $xlApp = New-ExcelInstance
    Write-Log "Excel COM ready. Excel version: $($xlApp.Version)" -Level INFO
} catch {
    Write-Log "FATAL: Could not initialise Excel COM application." -Level ERROR
    Write-Log "       Error: $($_.Exception.Message)" -Level ERROR
    Write-Log "       Microsoft Excel must be installed on this machine." -Level ERROR
    Write-Log "       All $totalFiles file(s) will be placed in the Review folder." -Level ERROR

    foreach ($f in $spreadsheetFiles) {
        $script:ErrorList.Add("$($f.Name)  |  REASON: Excel COM unavailable - $($_.Exception.Message)")
        try {
            Copy-ToOutput -SourcePath $f.FullName -DestDir $dirReview -OriginalName $f.Name | Out-Null
            Write-Log "  Copied to Review: $($f.Name)" -Level WARN
        } catch {
            Write-Log "  FAILED to copy '$($f.Name)' to Review: $($_.Exception.Message)" -Level ERROR
        }
    }
    Save-OutputFiles
    exit 1
}

# =============================================================================
# PROCESS EACH FILE
# =============================================================================

$countHasContent = 0
$countNoContent  = 0
$countReview     = 0
$fileIndex       = 0

foreach ($file in $spreadsheetFiles) {
    $fileIndex++
    $fileName = $file.Name
    $filePath = $file.FullName

    $szStr = if   ($file.Length -ge 1MB) { "$([math]::Round($file.Length / 1MB, 2)) MB" }
             elseif ($file.Length -ge 1KB) { "$([math]::Round($file.Length / 1KB, 2)) KB" }
             else                          { "$($file.Length) bytes" }

    Write-LogSection "FILE $fileIndex of $totalFiles : $fileName"
    Write-Log "Full path    : $filePath" -Level INFO
    Write-Log "Size         : $szStr" -Level INFO
    Write-Log "Extension    : $($file.Extension.ToLower())" -Level INFO
    Write-Log "Last modified: $($file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level INFO
    Write-Log "Created      : $($file.CreationTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level INFO

    # --- Confirm the file still exists (could have been removed during a long run) ---
    if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
        Write-Log "ERROR: File no longer found at '$filePath'." -Level ERROR
        Write-Log "       It may have been moved, renamed, or deleted during processing." -Level ERROR
        $script:ErrorList.Add("$fileName  |  REASON: File not found at path '$filePath' during processing")
        $countReview++
        continue   # No file to copy - skip
    }

    # --- Check the file is not locked by another process ---
    Write-Log "Checking file lock status..." -Level DEBUG
    try {
        $fs = [System.IO.File]::Open($filePath, 'Open', 'Read', 'ReadWrite')
        $fs.Close()
        $fs.Dispose()
        Write-Log "File lock check passed - file is accessible." -Level DEBUG
    } catch {
        $lockMsg = $_.Exception.Message
        Write-Log "ERROR: '$fileName' is locked or inaccessible." -Level ERROR
        Write-Log "       $lockMsg" -Level ERROR
        Write-Log "       The file may be open in another application. Close it and re-run." -Level ERROR
        $script:ErrorList.Add("$fileName  |  REASON: File locked / inaccessible - $lockMsg")
        try {
            $dest = Copy-ToOutput -SourcePath $filePath -DestDir $dirReview -OriginalName $fileName
            Write-Log "Copied to Review: $dest" -Level WARN
        } catch {
            Write-Log "FAILED to copy locked file to Review: $($_.Exception.Message)" -Level ERROR
        }
        $countReview++
        continue
    }

    # --- Inspect workbook ---
    Write-Log "Starting workbook inspection..." -Level INFO
    $result = Invoke-WorkbookInspection -FilePath $filePath -XlApp $xlApp

    # --- Handle inspection failure ---
    if (-not $result.Inspected) {
        Write-Log "INSPECTION FAILED for '$fileName'." -Level ERROR
        Write-Log "  Error detail: $($result.ErrorMessage)" -Level ERROR
        Write-Log "  File will be placed in Review for manual examination." -Level WARN
        $script:ErrorList.Add("$fileName  |  REASON: Inspection failed - $($result.ErrorMessage)")
        try {
            $dest = Copy-ToOutput -SourcePath $filePath -DestDir $dirReview -OriginalName $fileName
            Write-Log "Copied to Review: $dest" -Level WARN
        } catch {
            Write-Log "FAILED to copy to Review: $($_.Exception.Message)" -Level ERROR
        }
        $countReview++
        continue
    }

    # --- Compile findings list ---
    $findings = [System.Collections.Generic.List[string]]::new()

    if ($result.HasVeryHiddenSheets) {
        $result.VeryHiddenSheets | ForEach-Object {
            $findings.Add("VERY HIDDEN SHEET    : '$_'")
        }
    }
    if ($result.HasHiddenSheets) {
        $result.HiddenSheets | ForEach-Object {
            $findings.Add("HIDDEN SHEET         : '$_'")
        }
    }
    if ($result.HasHiddenRows) {
        $result.HiddenRowDetails | ForEach-Object {
            $findings.Add("HIDDEN ROWS          : $_")
        }
    }
    if ($result.HasHiddenColumns) {
        $result.HiddenColDetails | ForEach-Object {
            $findings.Add("HIDDEN COLUMNS       : $_")
        }
    }
    if ($result.HasExternalLinks) {
        $result.ExternalLinks | ForEach-Object {
            $findings.Add("EXTERNAL LINK        : $_")
        }
    }
    if ($result.HasFormulas) {
        $findings.Add("FORMULAS             : found on sheet(s): $($result.FormulaSheets -join ', ')")
    }
    if ($result.HasPivotTables) {
        $result.PivotDetails | ForEach-Object {
            $findings.Add("PIVOT TABLE          : $_")
        }
    }

    # --- Route to output directory ---
    $hasFlaggedContent = (
        $result.HasVeryHiddenSheets -or
        $result.HasHiddenSheets     -or
        $result.HasHiddenRows       -or
        $result.HasHiddenColumns    -or
        $result.HasExternalLinks    -or
        $result.HasFormulas         -or
        $result.HasPivotTables
    )

    if ($hasFlaggedContent) {
        Write-Log "RESULT: '$fileName'  >>>  HAS FLAGGED CONTENT  ($($findings.Count) finding(s))" -Level RESULT
        foreach ($finding in $findings) {
            Write-Log "   >> $finding" -Level RESULT
        }
        try {
            $dest = Copy-ToOutput -SourcePath $filePath -DestDir $dirHasContent -OriginalName $fileName
            Write-Log "File copied to 'Has content' : $dest" -Level INFO
        } catch {
            Write-Log "ERROR: Failed to copy '$fileName' to 'Has content' : $($_.Exception.Message)" -Level ERROR
            $script:ErrorList.Add("$fileName  |  REASON: Failed to copy to 'Has content' - $($_.Exception.Message)")
        }
        $countHasContent++

    } else {
        Write-Log "RESULT: '$fileName'  >>>  NO FLAGGED CONTENT detected" -Level RESULT
        try {
            $dest = Copy-ToOutput -SourcePath $filePath -DestDir $dirNoContent -OriginalName $fileName
            Write-Log "File copied to 'Does not have content' : $dest" -Level INFO
        } catch {
            Write-Log "ERROR: Failed to copy '$fileName' to 'Does not have content' : $($_.Exception.Message)" -Level ERROR
            $script:ErrorList.Add("$fileName  |  REASON: Failed to copy to 'Does not have content' - $($_.Exception.Message)")
        }
        $countNoContent++
    }

} # end foreach file

# =============================================================================
# CLEANUP
# =============================================================================

Write-LogSection "CLEANUP"
Write-Log "Releasing Excel COM application..." -Level INFO
Remove-ExcelInstance -xl $xlApp
Write-Log "Excel COM application released successfully." -Level INFO

# =============================================================================
# FINAL SUMMARY
# =============================================================================

$endTime  = Get-Date
$duration = $endTime - $script:StartTime

Write-LogSection "PROCESSING COMPLETE - SUMMARY"
Write-Log "Finished at        : $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level INFO
Write-Log "Total duration     : $($duration.ToString('hh\:mm\:ss\.fff'))" -Level INFO
Write-Log "Files scanned      : $totalFiles" -Level INFO
Write-Log "Has content        : $countHasContent file(s)  ->  $dirHasContent"  -Level RESULT
Write-Log "No content         : $countNoContent file(s)  ->  $dirNoContent"    -Level RESULT
Write-Log "Review required    : $countReview file(s)  ->  $dirReview"          -Level RESULT
Write-Log "Errors logged      : $($script:ErrorList.Count)"                    -Level INFO
Write-Log "Log saved to       : $pathLogs"                                      -Level INFO
Write-Log "Error file saved to: $pathErrors"                                    -Level INFO

Save-OutputFiles

Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Green
Write-Host "  QC CHECK COMPLETE" -ForegroundColor Green
Write-Host "  Output : $outputDir" -ForegroundColor Green
Write-Host ("=" * 80) -ForegroundColor Green
