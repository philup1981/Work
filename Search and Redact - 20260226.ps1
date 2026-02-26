#Requires -Version 5.0
<#
.SYNOPSIS
    Search and Redact Tool - Replaces specified search terms in spreadsheets with defined replacement text.

.DESCRIPTION
    This script reads search terms from a text file (one term per line) and replaces ALL occurrences
    in the specified spreadsheet across ALL worksheets. It supports Excel formats (.xlsx, .xlsm, .xls,
    .xlsb, etc.) via Excel COM automation and CSV/text files natively. No external modules required.

    Output files are created in an "Output" directory located one level above the spreadsheet's folder:
      - Updated_<original filename>  : The modified spreadsheet
      - Results.txt                  : Replacement counts per search term
      - Error.txt                    : Any errors or warnings encountered

.NOTES
    File Name      : Search and Redact - 20260226.ps1
    Requires       : PowerShell 5.0+, Microsoft Excel (for Excel file formats only)
    Compatibility  : Windows 10 / Windows Server 2019 and later
    Network Paths  : Fully supported (UNC paths e.g. \\server\share\file.xlsx)
    Permissions    : Runs under the current user account
    Created        : 2026-02-26
#>

# ==============================================================================
# FUNCTIONS
# ==============================================================================

function Write-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "=================================================================" -ForegroundColor Cyan
    Write-Host "   Search and Redact Tool  -  20260226" -ForegroundColor Cyan
    Write-Host "=================================================================" -ForegroundColor Cyan
    Write-Host "   Replaces search terms in spreadsheets with defined text" -ForegroundColor Gray
    Write-Host "   Supports: .xlsx .xlsm .xls .xlsb .csv .tsv and more" -ForegroundColor Gray
    Write-Host "   Network paths (UNC) supported" -ForegroundColor Gray
    Write-Host "=================================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Level) {
        "SUCCESS" { Write-Host "[$ts] [SUCCESS] $Message" -ForegroundColor Green }
        "WARN"    { Write-Host "[$ts] [WARN]    $Message" -ForegroundColor Yellow }
        "ERROR"   { Write-Host "[$ts] [ERROR]   $Message" -ForegroundColor Red }
        default   { Write-Host "[$ts] [INFO]    $Message" -ForegroundColor White }
    }
}

function Write-ErrorEntry {
    <#
    .SYNOPSIS Writes a timestamped entry to the error log file AND the console. #>
    param(
        [string]$ErrorLogPath,
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "ERROR"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$ts] [$Level] $Message"
    try {
        Add-Content -Path $ErrorLogPath -Value $entry -ErrorAction Stop
    }
    catch {
        Write-Host "[$ts] [ERROR] Could not write to error log. Original message follows:" -ForegroundColor Red
        Write-Host "[$ts] [ERROR] $Message" -ForegroundColor Red
    }
    Write-Log -Message $Message -Level $Level
}

function Test-FileAccessible {
    <#
    .SYNOPSIS
        Verifies that a file exists, is a file (not a directory), and is readable.
        Returns a PSCustomObject with Success (bool) and Message (string).
    #>
    param(
        [string]$FilePath,
        [string]$FileLabel
    )

    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        return [PSCustomObject]@{
            Success = $false
            Message = "The $FileLabel path is empty. A valid full file path must be provided."
        }
    }

    # Test existence
    $existsAsLeaf      = Test-Path -LiteralPath $FilePath -PathType Leaf
    $existsAsContainer = Test-Path -LiteralPath $FilePath -PathType Container

    if ($existsAsContainer) {
        return [PSCustomObject]@{
            Success = $false
            Message = "The $FileLabel path points to a FOLDER, not a file: '$FilePath'. Please provide the full path including the filename and extension."
        }
    }

    if (-not $existsAsLeaf) {
        return [PSCustomObject]@{
            Success = $false
            Message = "The $FileLabel file was NOT FOUND: '$FilePath'. Please verify the path and filename are correct. If this is a network path (\\server\share\...), ensure you are connected to the network and have access to the share."
        }
    }

    # Test read access by attempting to open the file
    try {
        $stream = [System.IO.File]::Open(
            $FilePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )
        $stream.Close()
        $stream.Dispose()
        return [PSCustomObject]@{
            Success = $true
            Message = "File is accessible: '$FilePath'"
        }
    }
    catch [System.UnauthorizedAccessException] {
        return [PSCustomObject]@{
            Success = $false
            Message = "ACCESS DENIED to $FileLabel file: '$FilePath'. The current user account '$($env:USERDOMAIN)\$($env:USERNAME)' does not have READ permission. Please check file/share permissions and try again."
        }
    }
    catch [System.IO.IOException] {
        return [PSCustomObject]@{
            Success = $false
            Message = "I/O ERROR accessing $FileLabel file: '$FilePath'. The file may be locked by another application, or the network path may be temporarily unavailable. Error detail: $($_.Exception.Message)"
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            Message = "UNEXPECTED ERROR accessing $FileLabel file: '$FilePath'. Error detail: $($_.Exception.Message)"
        }
    }
}

function Get-SearchTermsFromFile {
    <#
    .SYNOPSIS Reads the search terms text file. Returns a PSCustomObject with Success, Message, and Terms[]. #>
    param([string]$FilePath)

    try {
        $lines = Get-Content -LiteralPath $FilePath -ErrorAction Stop
        $terms = @(
            $lines |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() }
        )

        if ($terms.Count -eq 0) {
            return [PSCustomObject]@{
                Success = $false
                Message = "The search terms file contains no valid terms. The file may be empty or contain only blank lines: '$FilePath'"
                Terms   = @()
            }
        }

        return [PSCustomObject]@{
            Success = $true
            Message = "Loaded $($terms.Count) search term(s) from '$FilePath'"
            Terms   = $terms
        }
    }
    catch [System.UnauthorizedAccessException] {
        return [PSCustomObject]@{
            Success = $false
            Message = "ACCESS DENIED reading search terms file: '$FilePath'. Current user: '$($env:USERDOMAIN)\$($env:USERNAME)'"
            Terms   = @()
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            Message = "FAILED to read search terms file: '$FilePath'. Error: $($_.Exception.Message)"
            Terms   = @()
        }
    }
}

function Get-SafeRegexReplacement {
    <#
    .SYNOPSIS
        Escapes a literal replacement string for use with [regex]::Replace().
        In .NET regex replacement strings, '$' has special meaning (e.g. $1 = capture group 1).
        Replacing each '$' with '$$' makes it a literal dollar sign in the output.
    #>
    param([string]$Text)
    # '$' -> '$$' : each input '$' produces a literal '$' in the regex replacement output
    return $Text.Replace('$', '$$')
}

function Test-ExcelInstalled {
    <# .SYNOPSIS Returns $true if Microsoft Excel COM automation is available on this machine. #>
    try {
        $type = [System.Type]::GetTypeFromProgID("Excel.Application")
        return ($null -ne $type)
    }
    catch {
        return $false
    }
}

function Get-ExcelFileFormat {
    <#
    .SYNOPSIS Returns the Excel FileFormat constant for a given file extension.
             Defaults to 51 (xlOpenXMLWorkbook / .xlsx) for unknown extensions.
    #>
    param([string]$Extension)
    $ext = $Extension.ToLower().TrimStart('.')
    $formatMap = @{
        'xlsx' = 51  # xlOpenXMLWorkbook
        'xlsm' = 52  # xlOpenXMLWorkbookMacroEnabled
        'xls'  = 56  # xlExcel8
        'xlsb' = 50  # xlExcel12 (Binary)
        'xlam' = 55  # xlOpenXMLAddIn
        'xltx' = 54  # xlOpenXMLTemplate
        'xltm' = 53  # xlOpenXMLTemplateMacroEnabled
        'xlt'  = 17  # xlTemplate
        'xlw'  = 35  # xlWorkspace
    }
    if ($formatMap.ContainsKey($ext)) {
        return $formatMap[$ext]
    }
    return 51  # Default to xlsx
}

function Process-ExcelWorkbook {
    <#
    .SYNOPSIS
        Opens an Excel workbook via COM automation, searches all worksheets for each search term,
        replaces occurrences in non-formula text cells, and saves the result to OutputPath.
        Formula cells containing a search term are flagged as warnings (not auto-replaced).
        Returns a PSCustomObject with Success, Counts (ordered hashtable), and FormulaWarnings (list).
    #>
    param(
        [string]   $InputPath,
        [string]   $OutputPath,
        [string[]] $SearchTerms,
        [string]   $ReplacementText,
        [string]   $ErrorLogPath
    )

    # Initialise per-term replacement counts
    $counts = [ordered]@{}
    foreach ($term in $SearchTerms) {
        if (-not $counts.Contains($term)) {
            $counts[$term] = 0
        }
    }
    $formulaWarnings = [System.Collections.Generic.List[string]]::new()
    $safeReplacement = Get-SafeRegexReplacement -Text $ReplacementText
    $fileFormat      = Get-ExcelFileFormat -Extension ([System.IO.Path]::GetExtension($OutputPath))

    $excel    = $null
    $workbook = $null
    $success  = $false

    try {
        # -- Create Excel COM instance -----------------------------------------
        Write-Log "Initialising Microsoft Excel COM automation..."
        try {
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        }
        catch {
            Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "ERROR" `
                -Message "CRITICAL: Cannot create Excel COM object. Microsoft Excel may not be installed or may be blocked by a policy. Error: $($_.Exception.Message)"
            return [PSCustomObject]@{ Success = $false; Counts = $counts; FormulaWarnings = $formulaWarnings }
        }

        $excel.Visible        = $false
        $excel.DisplayAlerts  = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents   = $false
        # Note: Calculation is set after workbook open - some Excel versions
        # throw 0x800A03EC if set on the Application before any workbook is loaded.

        # -- Open workbook -----------------------------------------------------
        Write-Log "Opening workbook: '$InputPath'"
        try {
            # Open(Filename, UpdateLinks=0, ReadOnly=false)
            $workbook = $excel.Workbooks.Open($InputPath, 0, $false)
        }
        catch {
            Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "ERROR" `
                -Message "CRITICAL: Failed to open spreadsheet '$InputPath'. The file may be password-protected, corrupted, or in an unsupported format. Error: $($_.Exception.Message)"
            return [PSCustomObject]@{ Success = $false; Counts = $counts; FormulaWarnings = $formulaWarnings }
        }

        $excel.Calculation = -4135   # xlCalculationManual - disables auto-recalc for speed
        $totalSheets = $workbook.Worksheets.Count
        Write-Log "Workbook opened. Found $totalSheets worksheet(s)."

        # -- Process each worksheet --------------------------------------------
        $sheetIndex = 0
        foreach ($worksheet in $workbook.Worksheets) {
            $sheetIndex++
            $sheetName = $worksheet.Name
            Write-Log "  Processing worksheet $sheetIndex of $totalSheets : '$sheetName'"

            $usedRange = $null
            try {
                $usedRange = $worksheet.UsedRange

                if ($null -eq $usedRange) {
                    Write-Log "    Worksheet '$sheetName' has no used range - skipping."
                    continue
                }

                $totalRows = $usedRange.Rows.Count
                $totalCols = $usedRange.Columns.Count
                Write-Log "    Used range: $totalRows row(s) x $totalCols column(s)"

                $cellsModified = 0

                for ($row = 1; $row -le $totalRows; $row++) {

                    # Progress update every 500 rows to keep output readable
                    if ($row -gt 1 -and ($row % 500 -eq 1)) {
                        Write-Log "    Progress: $($row - 1) / $totalRows rows completed in '$sheetName'..."
                    }

                    for ($col = 1; $col -le $totalCols; $col++) {
                        $cell = $null
                        try {
                            $cell = $usedRange.Cells.Item($row, $col)

                            # -- Formula cells: warn, do not replace ----------
                            if ($cell.HasFormula) {
                                $displayedText = $cell.Text
                                if (-not [string]::IsNullOrEmpty($displayedText)) {
                                    foreach ($term in $SearchTerms) {
                                        if ([string]::IsNullOrEmpty($term)) { continue }
                                        $pattern = [regex]::Escape($term)
                                        if ([regex]::IsMatch($displayedText, $pattern,
                                                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                                            $absRow = $usedRange.Row    + $row - 1
                                            $absCol = $usedRange.Column + $col - 1
                                            $addr   = try { $cell.Address } catch { "R${absRow}C${absCol}" }
                                            $warnMsg = "Search term '$term' found in FORMULA CELL - Sheet='$sheetName', Cell=$addr (Row=$absRow, Col=$absCol). Formula cells are not automatically replaced to avoid breaking formulas. Manual review required."
                                            $formulaWarnings.Add($warnMsg)
                                        }
                                    }
                                }
                                continue   # skip to next cell
                            }

                            # -- Non-formula cells: get value -----------------
                            $cellValue = $cell.Value2

                            # Only process non-null string values
                            if ($null -eq $cellValue)               { continue }
                            if (-not ($cellValue -is [string]))     { continue }
                            if ([string]::IsNullOrEmpty($cellValue)) { continue }

                            $newValue = $cellValue

                            # Apply each search term
                            foreach ($term in $SearchTerms) {
                                if ([string]::IsNullOrEmpty($term)) { continue }

                                $pattern     = [regex]::Escape($term)
                                $regexOpts   = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
                                $matchCount  = [regex]::Matches($newValue, $pattern, $regexOpts).Count

                                if ($matchCount -gt 0) {
                                    $counts[$term] += $matchCount
                                    $newValue = [regex]::Replace($newValue, $pattern, $safeReplacement, $regexOpts)
                                }
                            }

                            # Write back only if the value changed
                            if ($newValue -ne $cellValue) {
                                $cell.Value2 = $newValue
                                $cellsModified++
                            }
                        }
                        catch {
                            $addr = if ($null -ne $cell) { try { $cell.Address } catch { "Unknown" } } else { "Unknown" }
                            Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "ERROR" `
                                -Message "Error processing cell at Sheet='$sheetName', Cell=$addr : $($_.Exception.Message)"
                        }
                        finally {
                            if ($null -ne $cell) {
                                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null
                                $cell = $null
                            }
                        }
                    } # end col loop
                } # end row loop

                Write-Log "    Completed '$sheetName': $cellsModified cell(s) modified."
            }
            catch {
                Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "ERROR" `
                    -Message "Error processing worksheet '$sheetName': $($_.Exception.Message)"
            }
            finally {
                if ($null -ne $usedRange) {
                    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null
                    $usedRange = $null
                }
            }
        } # end worksheet loop

        # Log formula warnings to error log
        if ($formulaWarnings.Count -gt 0) {
            Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "WARN" `
                -Message "-------- FORMULA CELL WARNINGS: $($formulaWarnings.Count) cell(s) require manual review --------"
            foreach ($w in $formulaWarnings) {
                Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "WARN" -Message $w
            }
            Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "WARN" `
                -Message "-------- End of formula cell warnings --------"
        }

        # Re-enable auto-calculation before saving
        $excel.Calculation = -4105   # xlCalculationAutomatic

        # -- Save workbook to output path --------------------------------------
        Write-Log "Saving updated workbook to: '$OutputPath'"
        try {
            # Remove existing output file to avoid format-mismatch prompts
            if (Test-Path -LiteralPath $OutputPath) {
                Remove-Item -LiteralPath $OutputPath -Force -ErrorAction Stop
            }
            $workbook.SaveAs($OutputPath, $fileFormat)
            Write-Log "Workbook saved successfully." -Level "SUCCESS"
            $success = $true
        }
        catch {
            Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "ERROR" `
                -Message "CRITICAL: Failed to save workbook to '$OutputPath'. The output location may not be writable, or another application has the file open. Error: $($_.Exception.Message)"
        }
    }
    catch {
        Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "ERROR" `
            -Message "CRITICAL: Unexpected error during Excel processing: $($_.Exception.Message)"
    }
    finally {
        # -- Release all COM objects -------------------------------------------
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            $workbook = $null
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            $excel = $null
        }
        # Force garbage collection to prevent orphaned Excel processes
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Write-Log "Excel COM objects released."
    }

    return [PSCustomObject]@{
        Success         = $success
        Counts          = $counts
        FormulaWarnings = $formulaWarnings
    }
}

function Process-CsvFile {
    <#
    .SYNOPSIS
        Processes a CSV or plain-text file by performing search-and-replace on the raw content.
        Preserves the original encoding (reads/writes as UTF-8).
        Returns a PSCustomObject with Success, Counts (ordered hashtable), and FormulaWarnings (empty list).
    #>
    param(
        [string]   $InputPath,
        [string]   $OutputPath,
        [string[]] $SearchTerms,
        [string]   $ReplacementText,
        [string]   $ErrorLogPath
    )

    $counts = [ordered]@{}
    foreach ($term in $SearchTerms) {
        if (-not $counts.Contains($term)) {
            $counts[$term] = 0
        }
    }
    $safeReplacement = Get-SafeRegexReplacement -Text $ReplacementText

    try {
        Write-Log "Reading file: '$InputPath'"
        $rawContent = [System.IO.File]::ReadAllText($InputPath, [System.Text.Encoding]::UTF8)

        if ([string]::IsNullOrEmpty($rawContent)) {
            Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "WARN" `
                -Message "The file is empty: '$InputPath'. An empty output file will be created."
            [System.IO.File]::WriteAllText($OutputPath, '', [System.Text.Encoding]::UTF8)
            return [PSCustomObject]@{ Success = $true; Counts = $counts; FormulaWarnings = @() }
        }

        $modifiedContent = $rawContent

        foreach ($term in $SearchTerms) {
            if ([string]::IsNullOrEmpty($term)) { continue }
            $pattern    = [regex]::Escape($term)
            $regexOpts  = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            $matchCount = [regex]::Matches($modifiedContent, $pattern, $regexOpts).Count
            if ($matchCount -gt 0) {
                $counts[$term]   += $matchCount
                $modifiedContent  = [regex]::Replace($modifiedContent, $pattern, $safeReplacement, $regexOpts)
            }
        }

        [System.IO.File]::WriteAllText($OutputPath, $modifiedContent, [System.Text.Encoding]::UTF8)
        Write-Log "File saved to: '$OutputPath'" -Level "SUCCESS"

        return [PSCustomObject]@{ Success = $true; Counts = $counts; FormulaWarnings = @() }
    }
    catch [System.UnauthorizedAccessException] {
        Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "ERROR" `
            -Message "ACCESS DENIED reading/writing file. Input='$InputPath', Output='$OutputPath'. Current user: '$($env:USERDOMAIN)\$($env:USERNAME)'. Error: $($_.Exception.Message)"
        return [PSCustomObject]@{ Success = $false; Counts = $counts; FormulaWarnings = @() }
    }
    catch {
        Write-ErrorEntry -ErrorLogPath $ErrorLogPath -Level "ERROR" `
            -Message "FAILED to process file '$InputPath'. Error: $($_.Exception.Message)"
        return [PSCustomObject]@{ Success = $false; Counts = $counts; FormulaWarnings = @() }
    }
}

# ==============================================================================
# MAIN SCRIPT
# ==============================================================================

Write-Banner
$scriptStartTime = Get-Date

# ------------------------------------------------------------------------------
# STEP 1 - Collect user inputs
# ------------------------------------------------------------------------------

Write-Host "Please provide the following information." -ForegroundColor Yellow
Write-Host "Tip: You can paste paths directly, including network paths (e.g. \\\\server\\share\\file.xlsx)." -ForegroundColor Gray
Write-Host ""

# -- Spreadsheet path ----------------------------------------------------------
do {
    Write-Host "SPREADSHEET" -ForegroundColor Cyan
    $spreadsheetPath = (Read-Host "  Enter the full path and filename of the spreadsheet").Trim().Trim('"').Trim("'")
    if ([string]::IsNullOrWhiteSpace($spreadsheetPath)) {
        Write-Host "  [ERROR] Path cannot be empty. Please try again." -ForegroundColor Red
        Write-Host ""
    }
} while ([string]::IsNullOrWhiteSpace($spreadsheetPath))

Write-Host ""

# -- Search terms file path ----------------------------------------------------
do {
    Write-Host "SEARCH TERMS FILE" -ForegroundColor Cyan
    Write-Host "  (Plain text file, one search term per line)" -ForegroundColor Gray
    $searchTermsPath = (Read-Host "  Enter the full path and filename of the search terms file").Trim().Trim('"').Trim("'")
    if ([string]::IsNullOrWhiteSpace($searchTermsPath)) {
        Write-Host "  [ERROR] Path cannot be empty. Please try again." -ForegroundColor Red
        Write-Host ""
    }
} while ([string]::IsNullOrWhiteSpace($searchTermsPath))

Write-Host ""

# -- Replacement text ----------------------------------------------------------
Write-Host "REPLACEMENT TEXT" -ForegroundColor Cyan
Write-Host "  This text replaces every matched search term in the spreadsheet." -ForegroundColor Gray
Write-Host "  Square brackets [ ] are fully supported (e.g. [REDACTED], [REMOVED])." -ForegroundColor Gray
Write-Host "  Maximum recommended length: 125 characters." -ForegroundColor Gray
Write-Host ""

$replacementText         = ""
$replacementTextAccepted = $false

do {
    $replacementText = Read-Host "  Enter replacement text"

    if ([string]::IsNullOrEmpty($replacementText)) {
        Write-Host ""
        Write-Host "  [WARNING] You entered EMPTY replacement text." -ForegroundColor Yellow
        Write-Host "            Every matched search term will be DELETED (replaced with nothing)." -ForegroundColor Yellow
        $confirm = Read-Host "  Are you sure you want empty replacement text? (Y/N)"
        if ($confirm -match '^[Yy]') {
            $replacementTextAccepted = $true
        }
        else {
            Write-Host "  Please enter your replacement text again." -ForegroundColor Gray
            Write-Host ""
        }
    }
    elseif ($replacementText.Length -gt 125) {
        Write-Host ""
        Write-Host "  [WARNING] Replacement text is $($replacementText.Length) characters long." -ForegroundColor Yellow
        Write-Host "            This exceeds the recommended maximum of 125 characters." -ForegroundColor Yellow
        Write-Host "            Very long replacement text may exceed Excel cell character limits (32,767 chars)." -ForegroundColor Yellow
        Write-Host "  Text entered: '$replacementText'" -ForegroundColor Yellow
        Write-Host ""
        $confirm = Read-Host "  Do you want to continue with this replacement text anyway? (Y/N)"
        if ($confirm -match '^[Yy]') {
            $replacementTextAccepted = $true
        }
        else {
            Write-Host "  Please enter a shorter replacement text." -ForegroundColor Gray
            Write-Host ""
        }
    }
    else {
        $replacementTextAccepted = $true
    }
} while (-not $replacementTextAccepted)

Write-Host ""
Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "  Configuration Summary" -ForegroundColor White
Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "  Spreadsheet      : $spreadsheetPath" -ForegroundColor White
Write-Host "  Search Terms     : $searchTermsPath" -ForegroundColor White
Write-Host "  Replacement Text : '$replacementText'" -ForegroundColor White
Write-Host "  Replacement Len  : $($replacementText.Length) character(s)" -ForegroundColor White
Write-Host "  Running As       : $($env:USERDOMAIN)\$($env:USERNAME)" -ForegroundColor White
Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

# ------------------------------------------------------------------------------
# STEP 2 - Validate input files (must exist and be readable BEFORE proceeding)
# ------------------------------------------------------------------------------

Write-Log "Validating input files..."

# Spreadsheet
$ssValidation = Test-FileAccessible -FilePath $spreadsheetPath -FileLabel "spreadsheet"
if (-not $ssValidation.Success) {
    Write-Host ""
    Write-Host "[FATAL ERROR] Cannot access the spreadsheet file:" -ForegroundColor Red
    Write-Host "  $($ssValidation.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting checklist:" -ForegroundColor Yellow
    Write-Host "  1. Is the path and filename correct (including the file extension)?" -ForegroundColor Yellow
    Write-Host "  2. Does the file exist at that location?" -ForegroundColor Yellow
    Write-Host "  3. Is the file open in another application (e.g. Excel)?" -ForegroundColor Yellow
    Write-Host "  4. For network paths: are you connected and do you have share/NTFS read rights?" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Log "Spreadsheet file OK." -Level "SUCCESS"

# Search terms file
$stValidation = Test-FileAccessible -FilePath $searchTermsPath -FileLabel "search terms"
if (-not $stValidation.Success) {
    Write-Host ""
    Write-Host "[FATAL ERROR] Cannot access the search terms file:" -ForegroundColor Red
    Write-Host "  $($stValidation.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshooting checklist:" -ForegroundColor Yellow
    Write-Host "  1. Is the path and filename correct (including the .txt extension)?" -ForegroundColor Yellow
    Write-Host "  2. Does the file exist at that location?" -ForegroundColor Yellow
    Write-Host "  3. For network paths: are you connected and do you have share/NTFS read rights?" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Log "Search terms file OK." -Level "SUCCESS"

# ------------------------------------------------------------------------------
# STEP 3 - Load search terms
# ------------------------------------------------------------------------------

Write-Log "Loading search terms..."
$termsResult = Get-SearchTermsFromFile -FilePath $searchTermsPath

if (-not $termsResult.Success) {
    Write-Host ""
    Write-Host "[FATAL ERROR] $($termsResult.Message)" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

$searchTerms = $termsResult.Terms
Write-Log "Loaded $($searchTerms.Count) search term(s):" -Level "SUCCESS"
for ($i = 0; $i -lt $searchTerms.Count; $i++) {
    Write-Log "  [$($i + 1)]  '$($searchTerms[$i])'"
}
Write-Host ""

# ------------------------------------------------------------------------------
# STEP 4 - Determine and create the Output directory
# ------------------------------------------------------------------------------

Write-Log "Configuring output directory..."

# Get the directory containing the spreadsheet
$ssItem = Get-Item -LiteralPath $spreadsheetPath -ErrorAction SilentlyContinue
$spreadsheetDir = if ($null -ne $ssItem) {
    $ssItem.DirectoryName
}
else {
    [System.IO.Path]::GetDirectoryName($spreadsheetPath)
}

if ([string]::IsNullOrEmpty($spreadsheetDir)) {
    Write-Host "[FATAL ERROR] Cannot determine the parent directory of: '$spreadsheetPath'" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# One level above the spreadsheet's folder
$parentDir = [System.IO.Path]::GetDirectoryName($spreadsheetDir)
if ([string]::IsNullOrEmpty($parentDir)) {
    # Edge case: spreadsheet is in the root of a drive or UNC share root
    Write-Log "Note: Spreadsheet is in a root-level folder. Output directory will be created alongside it." -Level "WARN"
    $parentDir = $spreadsheetDir
}

$outputDir             = [System.IO.Path]::Combine($parentDir, "Output")
$originalFileName      = [System.IO.Path]::GetFileName($spreadsheetPath)
$outputSpreadsheetPath = [System.IO.Path]::Combine($outputDir, "Updated_$originalFileName")
$errorLogPath          = [System.IO.Path]::Combine($outputDir, "Error.txt")
$resultsLogPath        = [System.IO.Path]::Combine($outputDir, "Results.txt")

Write-Log "Output directory    : '$outputDir'"
Write-Log "Updated spreadsheet : '$outputSpreadsheetPath'"
Write-Log "Results file        : '$resultsLogPath'"
Write-Log "Error log           : '$errorLogPath'"

# Create Output directory
try {
    if (-not (Test-Path -LiteralPath $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force -ErrorAction Stop | Out-Null
        Write-Log "Output directory created." -Level "SUCCESS"
    }
    else {
        Write-Log "Output directory already exists - files will be overwritten if present."
    }
}
catch [System.UnauthorizedAccessException] {
    Write-Host ""
    Write-Host "[FATAL ERROR] ACCESS DENIED: Cannot create output directory '$outputDir'." -ForegroundColor Red
    Write-Host "  The current user '$($env:USERDOMAIN)\$($env:USERNAME)' does not have WRITE permission to '$parentDir'." -ForegroundColor Red
    Write-Host "  Please choose a spreadsheet location where you have write access, or contact your administrator." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}
catch {
    Write-Host ""
    Write-Host "[FATAL ERROR] Failed to create output directory '$outputDir'. Error: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Initialise error log
try {
    @(
        "=================================================================",
        "  Search and Redact - Error Log",
        "  Script started : $($scriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))",
        "=================================================================",
        "  Spreadsheet     : $spreadsheetPath",
        "  Search Terms    : $searchTermsPath",
        "  Replacement Text: '$replacementText'",
        "  Running As      : $($env:USERDOMAIN)\$($env:USERNAME)",
        "=================================================================",
        ""
    ) | Set-Content -Path $errorLogPath -ErrorAction Stop
    Write-Log "Error log initialised: '$errorLogPath'"
}
catch {
    Write-Host "[WARNING] Could not initialise error log at '$errorLogPath'. Errors will be shown on screen only. Error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# ------------------------------------------------------------------------------
# STEP 5 - Process the spreadsheet
# ------------------------------------------------------------------------------

Write-Host ""
Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
Write-Log "Starting spreadsheet processing (accuracy-focused, cell-by-cell)..."
Write-Host "-----------------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

$fileExtension   = [System.IO.Path]::GetExtension($spreadsheetPath).ToLower().TrimStart('.')
$excelExtensions = @('xlsx','xlsm','xls','xlsb','xlam','xltx','xltm','xlt','xlw')
$csvExtensions   = @('csv','tsv','tab')
$textExtensions  = @('txt')

$processingResult = $null

if ($excelExtensions -contains $fileExtension) {
    # -- Excel workbook --------------------------------------------------------
    Write-Log "File type: Excel workbook (.$fileExtension)"

    if (-not (Test-ExcelInstalled)) {
        $msg = "CRITICAL: Microsoft Excel is not installed or not accessible on this machine. Excel COM automation is required to process '.$fileExtension' files. Please install Microsoft Excel 2019 or later on this machine, or convert the file to CSV format first."
        Write-ErrorEntry -ErrorLogPath $errorLogPath -Level "ERROR" -Message $msg
        Write-Host ""
        Write-Host "[FATAL ERROR] $msg" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }

    Write-Log "Microsoft Excel detected. Using COM automation."
    $processingResult = Process-ExcelWorkbook `
        -InputPath       $spreadsheetPath `
        -OutputPath      $outputSpreadsheetPath `
        -SearchTerms     $searchTerms `
        -ReplacementText $replacementText `
        -ErrorLogPath    $errorLogPath
}
elseif (($csvExtensions + $textExtensions) -contains $fileExtension) {
    # -- CSV / plain text ------------------------------------------------------
    Write-Log "File type: CSV/text file (.$fileExtension)"
    $processingResult = Process-CsvFile `
        -InputPath       $spreadsheetPath `
        -OutputPath      $outputSpreadsheetPath `
        -SearchTerms     $searchTerms `
        -ReplacementText $replacementText `
        -ErrorLogPath    $errorLogPath
}
else {
    # -- Unknown extension: try Excel COM first, then fail gracefully ----------
    Write-Log "File extension '.$fileExtension' is not in the standard list." -Level "WARN"
    Write-ErrorEntry -ErrorLogPath $errorLogPath -Level "WARN" `
        -Message "Unrecognised file extension '.$fileExtension'. Attempting Excel COM automation as a fallback."

    if (Test-ExcelInstalled) {
        Write-Log "Attempting to process with Excel COM automation..."
        $processingResult = Process-ExcelWorkbook `
            -InputPath       $spreadsheetPath `
            -OutputPath      $outputSpreadsheetPath `
            -SearchTerms     $searchTerms `
            -ReplacementText $replacementText `
            -ErrorLogPath    $errorLogPath
    }
    else {
        $msg = "CRITICAL: Unrecognised file extension '.$fileExtension' and Microsoft Excel is not available as a fallback. Supported Excel formats: .$($excelExtensions -join ', .') | Supported CSV/text formats: .$($csvExtensions + $textExtensions -join ', .'). Please convert your file to a supported format and try again."
        Write-ErrorEntry -ErrorLogPath $errorLogPath -Level "ERROR" -Message $msg
        Write-Host "[FATAL ERROR] $msg" -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# ------------------------------------------------------------------------------
# STEP 6 - Write Results.txt
# ------------------------------------------------------------------------------

$scriptEndTime = Get-Date
$duration      = $scriptEndTime - $scriptStartTime

Write-Host ""
Write-Log "Writing results file..."

$totalReplacements = 0
$resultsLines      = [System.Collections.Generic.List[string]]::new()

$resultsLines.Add("=================================================================")
$resultsLines.Add("  Search and Redact - Results")
$resultsLines.Add("  Generated  : $($scriptEndTime.ToString('yyyy-MM-dd HH:mm:ss'))")
$resultsLines.Add("  Duration   : $($duration.ToString('hh\:mm\:ss\.fff'))")
$resultsLines.Add("=================================================================")
$resultsLines.Add("  Spreadsheet     : $spreadsheetPath")
$resultsLines.Add("  Search Terms    : $searchTermsPath")
$resultsLines.Add("  Replacement Text: '$replacementText'")
$resultsLines.Add("  Running As      : $($env:USERDOMAIN)\$($env:USERNAME)")
$resultsLines.Add("=================================================================")
$resultsLines.Add("")
$resultsLines.Add("REPLACEMENT COUNTS PER SEARCH TERM")
$resultsLines.Add("-----------------------------------------------------------------")

foreach ($term in $searchTerms) {
    $count = 0
    if ($null -ne $processingResult -and
        $null -ne $processingResult.Counts -and
        $processingResult.Counts.Contains($term)) {
        $count = $processingResult.Counts[$term]
    }
    $totalReplacements += $count
    $label = if ($count -eq 0) { "0  (NOT FOUND - no replacements made)" } else { "$count replacement(s)" }
    $resultsLines.Add("  '$term'  :  $label")
}

$resultsLines.Add("")
$resultsLines.Add("-----------------------------------------------------------------")
$resultsLines.Add("  TOTAL REPLACEMENTS MADE : $totalReplacements")
$resultsLines.Add("-----------------------------------------------------------------")
$resultsLines.Add("")

# Formula warnings summary
if ($null -ne $processingResult -and
    $null -ne $processingResult.FormulaWarnings -and
    $processingResult.FormulaWarnings.Count -gt 0) {
    $resultsLines.Add("FORMULA CELL WARNINGS  ($($processingResult.FormulaWarnings.Count) cell(s) require manual review)")
    $resultsLines.Add("-----------------------------------------------------------------")
    $resultsLines.Add("  The following cells contain a search term in their DISPLAYED value,")
    $resultsLines.Add("  but were NOT automatically replaced because they contain formulas.")
    $resultsLines.Add("  Please review and update these cells manually:")
    $resultsLines.Add("")
    foreach ($w in $processingResult.FormulaWarnings) {
        $resultsLines.Add("  $w")
    }
    $resultsLines.Add("")
}

if ($null -ne $processingResult -and $processingResult.Success) {
    $resultsLines.Add("STATUS : COMPLETED SUCCESSFULLY")
    $resultsLines.Add("OUTPUT : $outputSpreadsheetPath")
}
else {
    $resultsLines.Add("STATUS : COMPLETED WITH ERRORS")
    $resultsLines.Add("         Please review the error log for details: $errorLogPath")
}

try {
    $resultsLines | Set-Content -Path $resultsLogPath -ErrorAction Stop
    Write-Log "Results written to: '$resultsLogPath'" -Level "SUCCESS"
}
catch {
    Write-Host "[ERROR] Failed to write Results.txt to '$resultsLogPath'. Error: $($_.Exception.Message)" -ForegroundColor Red
    try { Add-Content -Path $errorLogPath -Value "[$((Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))] [ERROR] Failed to write Results.txt: $($_.Exception.Message)" } catch { }
}

# Finalise error log
try {
    Add-Content -Path $errorLogPath -Value ""
    Add-Content -Path $errorLogPath -Value "================================================================="
    Add-Content -Path $errorLogPath -Value "  Script completed : $($scriptEndTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    Add-Content -Path $errorLogPath -Value "  Total duration   : $($duration.ToString('hh\:mm\:ss\.fff'))"
    Add-Content -Path $errorLogPath -Value "================================================================="
}
catch { }

# ------------------------------------------------------------------------------
# STEP 7 - Final summary on screen
# ------------------------------------------------------------------------------

Write-Host ""
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "  Processing Complete" -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Output directory  : $outputDir" -ForegroundColor White
$updatedColor = if ($processingResult -and $processingResult.Success) { "Green" } else { "Yellow" }
Write-Host "  Updated file      : $outputSpreadsheetPath" -ForegroundColor $updatedColor
Write-Host "  Results file      : $resultsLogPath" -ForegroundColor White
Write-Host "  Error log         : $errorLogPath" -ForegroundColor White
Write-Host ""
Write-Host "  Search Term Results:" -ForegroundColor Cyan

foreach ($term in $searchTerms) {
    $count = 0
    if ($null -ne $processingResult -and
        $null -ne $processingResult.Counts -and
        $processingResult.Counts.Contains($term)) {
        $count = $processingResult.Counts[$term]
    }
    if ($count -eq 0) {
        Write-Host "    [  0 hits  ]  '$term'" -ForegroundColor Gray
    }
    else {
        Write-Host "    [ $count replacement(s) ]  '$term'" -ForegroundColor Green
    }
}

Write-Host ""
$totalColor = if ($totalReplacements -gt 0) { "Green" } else { "Yellow" }
Write-Host "  Total replacements made : $totalReplacements" -ForegroundColor $totalColor
Write-Host "  Duration                : $($duration.ToString('hh\:mm\:ss\.fff'))" -ForegroundColor Gray
Write-Host ""

if ($null -ne $processingResult -and $processingResult.Success) {
    Write-Host "  STATUS : SUCCESS" -ForegroundColor Green
}
else {
    Write-Host "  STATUS : COMPLETED WITH ERRORS" -ForegroundColor Yellow
    Write-Host "           Review error log: $errorLogPath" -ForegroundColor Yellow
}

if ($null -ne $processingResult -and
    $null -ne $processingResult.FormulaWarnings -and
    $processingResult.FormulaWarnings.Count -gt 0) {
    Write-Host ""
    Write-Host "  ATTENTION: $($processingResult.FormulaWarnings.Count) formula cell(s) contained a search term" -ForegroundColor Yellow
    Write-Host "             but were NOT auto-replaced (formulas preserved)." -ForegroundColor Yellow
    Write-Host "             See Results.txt and Error.txt for cell locations." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
