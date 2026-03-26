#Requires -Version 5.1
<#
.SYNOPSIS
    Contains Charts or Images - 20260326.ps1

.DESCRIPTION
    Scans a folder of Microsoft Excel spreadsheets to detect charts and images.
    Categorises files into: "Has content", "Does not have content", "review".
    Runs a second QC pass. Supports local and UNC network paths.
    No external modules. Does NOT modify source files.

.PARAMETER FolderPath
    Path to the folder containing Excel spreadsheets to scan.

.EXAMPLE
    .\Contains Charts or Images - 20260326.ps1 -FolderPath "C:\Reports"
    .\Contains Charts or Images - 20260326.ps1 -FolderPath "\\server\share\Reports"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$FolderPath,

    # Maximum parallel analysis threads.
    # 0 = auto (half the logical CPU count, minimum 2, maximum 8).
    [Parameter(Mandatory = $false)]
    [ValidateRange(0,32)]
    [int]$MaxThreads = 0
)

# ============================================================
#  INTERACTIVE PATH PROMPT
# ============================================================
if ([string]::IsNullOrWhiteSpace($FolderPath)) {
    Write-Host ''
    Write-Host '============================================================' -ForegroundColor White
    Write-Host '  Contains Charts or Images - 20260326' -ForegroundColor White
    Write-Host '============================================================' -ForegroundColor White
    Write-Host ''
    Write-Host '  Enter the path to the folder containing your spreadsheets.' -ForegroundColor Cyan
    Write-Host '  Local example : C:\Finance\Reports' -ForegroundColor DarkGray
    Write-Host '  Network example: \\server\share\Reports' -ForegroundColor DarkGray
    Write-Host ''
    $FolderPath = (Read-Host '  Folder path').Trim()
    Write-Host ''
}

Set-StrictMode -Off
$ErrorActionPreference = 'Continue'

# ============================================================
#  CONSTANTS & INITIALISATION
# ============================================================
$ScriptName      = 'Contains Charts or Images - 20260326'
$ScriptVersion   = '1.0'
$ScriptStart     = Get-Date
$RunStamp        = $ScriptStart.ToString('yyyyMMdd_HHmmss')

$OoxmlExts  = @('.xlsx','.xlsm','.xlsb','.xltx','.xltm','.xlam')
$BinaryExts = @('.xls','.xlt','.xlw')
$AllExts    = $OoxmlExts + $BinaryExts

# Load ZIP assembly once
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression          -ErrorAction Stop
} catch {
    Write-Error "FATAL: Cannot load System.IO.Compression: $($_.Exception.Message)"
    exit 1
}

# Log / error collections
$Log    = [System.Text.StringBuilder]::new()
$Errors = [System.Collections.Generic.List[string]]::new()

# Counters
$cTotal   = 0
$cHas     = 0
$cNo      = 0
$cReview  = 0

# ============================================================
#  LOGGING HELPER
# ============================================================
function Write-Log {
    param(
        [string]$Msg,
        [ValidateSet('INFO','WARN','ERROR','OK','DBG','HEAD','SEP')]
        [string]$L = 'INFO'
    )
    if ($L -eq 'SEP') {
        $line = '-' * 70
        [void]$Log.AppendLine($line)
        Write-Host $line -ForegroundColor DarkGray
        return
    }
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $pad  = $L.PadRight(5)
    $line = "[$ts][$pad] $Msg"
    [void]$Log.AppendLine($line)
    $col  = switch ($L) {
        'INFO' { 'Cyan'     }
        'WARN' { 'Yellow'   }
        'ERROR'{ 'Red'      }
        'OK'   { 'Green'    }
        'DBG'  { 'DarkGray' }
        'HEAD' { 'White'    }
    }
    Write-Host $line -ForegroundColor $col
}

function Log-Error {
    param([string]$Id, [string]$Cat, [string]$Msg)
    $entry = "[$Cat] $Id | $Msg"
    $Errors.Add($entry)
    Write-Log "ERROR RECORDED: $entry" -L ERROR
}

# ============================================================
#  HEADER
# ============================================================
# Resolve MaxThreads before first log line so it appears in header
if ($MaxThreads -eq 0) {
    $MaxThreads = [Math]::Max(2, [Math]::Min(8, [int][Math]::Floor([Environment]::ProcessorCount / 2)))
}

Write-Log ('=' * 70) -L HEAD
Write-Log "  $ScriptName  v$ScriptVersion" -L HEAD
Write-Log "  Started  : $($ScriptStart.ToString('yyyy-MM-dd HH:mm:ss'))" -L HEAD
Write-Log "  User     : $($env:USERDOMAIN)\$($env:USERNAME)  on  $($env:COMPUTERNAME)" -L HEAD
Write-Log "  CPU cores: $([Environment]::ProcessorCount)  |  Analysis threads: $MaxThreads" -L HEAD
Write-Log ('=' * 70) -L HEAD
Write-Log ''

# ============================================================
#  VALIDATE INPUT FOLDER
# ============================================================
Write-Log "Validating folder: '$FolderPath'" -L INFO

if ([string]::IsNullOrWhiteSpace($FolderPath)) {
    Write-Log "FATAL: FolderPath is empty." -L ERROR; exit 1
}

if (-not (Test-Path -LiteralPath $FolderPath -PathType Container -ErrorAction SilentlyContinue)) {
    Write-Log "FATAL: Folder not found or not accessible: '$FolderPath'" -L ERROR
    Write-Log "  Check: spelling, network connectivity, and read permissions." -L ERROR
    exit 1
}

try   { $FolderPath = (Resolve-Path -LiteralPath $FolderPath -ErrorAction Stop).Path }
catch { Write-Log "WARNING: Could not resolve full path, using as-is. ($($_.Exception.Message))" -L WARN }

Write-Log "Resolved path : '$FolderPath'" -L INFO

# ============================================================
#  BUILD OUTPUT STRUCTURE
# ============================================================
$ParentDir  = Split-Path -Parent $FolderPath
$OutRoot    = Join-Path $ParentDir 'Output'
$DirHas     = Join-Path $OutRoot  'Has content'
$DirNo      = Join-Path $OutRoot  'Does not have content'
$DirReview  = Join-Path $OutRoot  'review'
$LogFile    = Join-Path $OutRoot  "logs_$RunStamp.txt"
$ErrFile    = Join-Path $OutRoot  "Errors_$RunStamp.txt"

Write-Log "Output root : '$OutRoot'" -L INFO

foreach ($d in @($OutRoot, $DirHas, $DirNo, $DirReview)) {
    if (-not (Test-Path -LiteralPath $d -PathType Container)) {
        try {
            New-Item -ItemType Directory -Path $d -Force -ErrorAction Stop | Out-Null
            Write-Log "  [CREATED] $d" -L OK
        } catch {
            Write-Log "FATAL: Cannot create '$d': $($_.Exception.Message)" -L ERROR
            exit 1
        }
    } else {
        Write-Log "  [EXISTS]  $d" -L DBG
    }
}

# ============================================================
#  PERMISSION CHECK ON SOURCE FOLDER
# ============================================================
Write-Log '' ; Write-Log "Verifying read access to source folder..." -L INFO
try {
    $acl = Get-Acl -LiteralPath $FolderPath -ErrorAction Stop
    Write-Log "  ACL owner: $($acl.Owner)" -L DBG
    Write-Log "  Read access confirmed." -L OK
} catch {
    Write-Log "  WARNING: Could not read folder ACL: $($_.Exception.Message)" -L WARN
    Write-Log "  Continuing - per-file errors will be captured individually." -L WARN
}

# ============================================================
#  DETECTION: OOXML (ZIP-based) formats
# ============================================================
function Get-OoxmlContent {
    param([System.IO.FileInfo]$File)

    $r = [PSCustomObject]@{
        HasChart    = $false; HasImage   = $false; HasSmartArt = $false; HasShape = $false
        ChartCount  = 0;      ImageCount = 0;      DrawingCount = 0;    ShapeCount = 0
        SmartArtCnt = 0;      Details    = [System.Collections.Generic.List[string]]::new()
        Error       = $null
    }
    try {
        $stream = [System.IO.File]::Open($File.FullName,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite)
        $zip = [System.IO.Compression.ZipArchive]::new($stream,
                    [System.IO.Compression.ZipArchiveMode]::Read, $false)
    } catch {
        $r.Error = "Cannot open as ZIP: $($_.Exception.Message)"
        if ($stream) { try { $stream.Dispose() } catch {} }
        return $r
    }

    try {
        $entries = @($zip.Entries | Select-Object -ExpandProperty FullName)

        # -- Charts: xl/charts/
        $chartFiles = @($entries | Where-Object { $_ -match '^xl[/\\]charts[/\\][^/\\]+\.xml$' })
        if ($chartFiles.Count -gt 0) {
            $r.HasChart   = $true
            $r.ChartCount = $chartFiles.Count
            foreach ($c in $chartFiles) { $r.Details.Add("Chart file: $c") }
        }

        # -- Chart sheets: xl/chartsheets/
        $csFiles = @($entries | Where-Object { $_ -match '^xl[/\\]chartsheets[/\\][^/\\]+\.xml$' -and $_ -notmatch '_rels' })
        if ($csFiles.Count -gt 0) {
            $r.HasChart    = $true
            $r.ChartCount += $csFiles.Count
            foreach ($c in $csFiles) { $r.Details.Add("Chart sheet: $c") }
        }

        # -- Images: xl/media/
        $imgExts = @('.png','.jpg','.jpeg','.gif','.bmp','.tif','.tiff',
                     '.wmf','.emf','.svg','.ico','.webp','.dib','.eps','.wdp','.hdp','.pcz','.pct')
        $mediaFiles = @($entries | Where-Object {
            if ($_ -match '^xl[/\\]media[/\\](.+)$') {
                $ext = [System.IO.Path]::GetExtension($Matches[1]).ToLower()
                return $imgExts -contains $ext
            }; $false
        })
        if ($mediaFiles.Count -gt 0) {
            $r.HasImage    = $true
            $r.ImageCount  = $mediaFiles.Count
            foreach ($m in $mediaFiles) { $r.Details.Add("Image file: $m") }
        }

        # -- SmartArt: xl/diagrams/
        $diagFiles = @($entries | Where-Object { $_ -match '^xl[/\\]diagrams[/\\][^/\\]+\.xml$' -and $_ -notmatch '_rels' })
        if ($diagFiles.Count -gt 0) {
            $r.HasSmartArt  = $true
            $r.SmartArtCnt  = $diagFiles.Count
            foreach ($d in $diagFiles) { $r.Details.Add("SmartArt diagram: $d") }
        }

        # -- Drawings XML: check for picture/chart elements inside
        $drawFiles = @($entries | Where-Object { $_ -match '^xl[/\\]drawings[/\\]drawing[^/\\]*\.xml$' })
        $r.DrawingCount = $drawFiles.Count
        foreach ($df in $drawFiles) {
            $entry = $zip.GetEntry($df)
            if (-not $entry) { continue }
            try {
                $rdr  = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
                $xml  = $rdr.ReadToEnd(); $rdr.Close()
                if ($xml -match 'blipFill|<xdr:pic[\s>]|<a:blip\s|<pic:pic[\s>]') {
                    $r.HasImage = $true
                    $r.Details.Add("Picture element in: $df")
                }
                if ($xml -match '<c:chart[\s/]|graphicFrame|c14:chart|cx:chart') {
                    $r.HasChart = $true
                    $r.Details.Add("Chart reference in: $df")
                }
                # Shapes: <xdr:sp> = basic shape (rect/arrow/callout/text box etc.)
                #         <xdr:cxnSp> = connector/line shape
                if ($xml -match '<xdr:sp[\s>]|<xdr:cxnSp[\s>]') {
                    $r.HasShape  = $true
                    $spCnt  = ([regex]::Matches($xml, '<xdr:sp[\s>]')).Count
                    $cxnCnt = ([regex]::Matches($xml, '<xdr:cxnSp[\s>]')).Count
                    $r.ShapeCount += $spCnt + $cxnCnt
                    $r.Details.Add("Shapes in $df`: sp=$spCnt connectors=$cxnCnt")
                }
            } catch {
                $r.Details.Add("WARNING: Could not read drawing '$df': $($_.Exception.Message)")
            }
        }

        # -- Relationship files: final sweep for chart/image refs
        $relsFiles = @($entries | Where-Object { $_ -match '\.rels$' })
        foreach ($rf in $relsFiles) {
            $entry = $zip.GetEntry($rf)
            if (-not $entry) { continue }
            try {
                $rdr = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
                $xml = $rdr.ReadToEnd(); $rdr.Close()
                if ($xml -match 'relationships/chart"') {
                    $r.HasChart = $true
                    $r.Details.Add("Chart relationship in: $rf")
                }
                if ($xml -match 'relationships/image"') {
                    $r.HasImage = $true
                    $r.Details.Add("Image relationship in: $rf")
                }
                if ($xml -match 'relationships/diagram') {
                    $r.HasSmartArt = $true
                    $r.Details.Add("Diagram relationship in: $rf")
                }
            } catch {
                $r.Details.Add("WARNING: Could not read rels '$rf': $($_.Exception.Message)")
            }
        }

    } catch {
        $r.Error = "ZIP analysis error: $($_.Exception.Message)"
    } finally {
        if ($zip)    { try { $zip.Dispose()    } catch {} }
        if ($stream) { try { $stream.Dispose() } catch {} }
    }
    return $r
}

# ============================================================
#  DETECTION: Binary BIFF8 (.xls / .xlt / .xlw)
# ============================================================
function Get-BinaryXlsContent {
    param([System.IO.FileInfo]$File)

    $r = [PSCustomObject]@{
        HasChart    = $false; HasImage   = $false; HasSmartArt = $false; HasShape = $false
        ChartCount  = 0;      ImageCount = 0;      DrawingCount = 0;    ShapeCount = 0
        SmartArtCnt = 0;      Details    = [System.Collections.Generic.List[string]]::new()
        Error       = $null
    }

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($File.FullName,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite)

        $fileLen = $stream.Length
        if ($fileLen -lt 8) {
            $r.Error = "File too small to be valid ($fileLen bytes)."
            return $r
        }

        # Check OLE2 magic from first 8 bytes (no full file load)
        $hdr = New-Object byte[] 8
        [void]$stream.Read($hdr, 0, 8)
        $stream.Position = 0
        $magic = [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
        $isOle = $true
        for ($m = 0; $m -lt 8; $m++) { if ($hdr[$m] -ne $magic[$m]) { $isOle = $false; break } }
        if ($isOle) {
            $r.Details.Add("OLE2 Compound Document confirmed.")
        } else {
            $r.Details.Add("WARNING: OLE2 magic not found - may be older BIFF or corrupted. Scanning anyway.")
        }

        # Streaming scan in 256 KB chunks with 8-byte overlap to catch cross-boundary patterns.
        # This avoids loading the entire file into memory.
        $CHUNK   = 256 * 1024
        $OVERLAP = 8
        $buf     = New-Object byte[] ($CHUNK + $OVERLAP)
        $prev    = New-Object byte[] $OVERLAP
        $gOffset = [long]0
        $first   = $true
        $msoDrw = 0; $msoDrwGrp = 0; $chartBof = 0; $imdata = 0

        while ($true) {
            if (-not $first) { [Array]::Copy($prev, 0, $buf, 0, $OVERLAP) }
            $readStart = if ($first) { 0 } else { $OVERLAP }
            $bytesRead = $stream.Read($buf, $readStart, $CHUNK)
            if ($bytesRead -eq 0) { break }

            $scanEnd = $readStart + $bytesRead
            $tailSrc = $scanEnd - $OVERLAP
            if ($tailSrc -ge 0) { [Array]::Copy($buf, $tailSrc, $prev, 0, $OVERLAP) }

            for ($i = 0; $i -lt ($scanEnd - 1); $i++) {
                $b0 = $buf[$i]; $b1 = $buf[$i + 1]
                if      ($b0 -eq 0xEC -and $b1 -eq 0x00) { $msoDrw++ }
                elseif  ($b0 -eq 0xEB -and $b1 -eq 0x00) { $msoDrwGrp++ }
                elseif  ($b0 -eq 0x7F -and $b1 -eq 0x00) {
                    $imdata++; $r.HasImage = $true
                    $r.Details.Add("IMDATA record at offset ~$($gOffset + $i)")
                }
                elseif ($b0 -eq 0x09 -and $b1 -eq 0x08 -and ($i + 7) -lt $scanEnd) {
                    $recLen = [BitConverter]::ToUInt16($buf, $i + 2)
                    if ($recLen -ge 4 -and ($i + 4 + $recLen) -lt $scanEnd) {
                        $bofType = [BitConverter]::ToUInt16($buf, $i + 6)
                        if ($bofType -eq 0x0020) {
                            $chartBof++; $r.HasChart = $true
                            $r.Details.Add("Chart BOF at offset ~$($gOffset + $i)")
                        }
                    }
                }
            }
            $gOffset += $bytesRead
            $first    = $false
        }

        if ($msoDrw -gt 0) {
            $r.DrawingCount = $msoDrw
            $r.Details.Add("MSODRAWING records: $msoDrw (covers charts, images and shapes in binary format)")
            # IMDATA records confirm actual embedded image data
            if ($imdata -gt 0) { $r.HasImage = $true }
            # Chart BOF confirms embedded charts
            # Any remaining MSODRAWING records beyond confirmed charts/images are shapes
            $remainingDrw = $msoDrw - $chartBof
            if ($remainingDrw -gt 0) {
                # In BIFF8 we cannot reliably distinguish images from shapes without
                # full OLE2 sector parsing, so flag both as potentially present
                $r.HasShape = $true
                $r.ShapeCount = $remainingDrw
                $r.Details.Add("Drawing objects (shapes/images) inferred from MSODRAWING: $remainingDrw")
                # If IMDATA is also present, images are confirmed; otherwise classify as shapes
                if ($imdata -eq 0) {
                    $r.Details.Add("No IMDATA records found - drawing objects classified as shapes (may include images)")
                }
            }
        }
        if ($msoDrwGrp -gt 0) { $r.Details.Add("MSODRAWINGGROUP records: $msoDrwGrp") }

        $r.ChartCount  = $chartBof
        $r.ImageCount  = $imdata
        $r.Details.Add("Binary scan totals - MSODRAWING:$msoDrw | MSODRAWINGGROUP:$msoDrwGrp | ChartBOF:$chartBof | IMDATA:$imdata")

    } catch {
        $r.Error = "Binary analysis failed: $($_.Exception.Message)"
    } finally {
        if ($stream) { try { $stream.Dispose() } catch {} }
    }
    return $r
}

# ============================================================
#  DISPATCHER
# ============================================================
function Invoke-Analysis {
    param([System.IO.FileInfo]$File)
    $ext = $File.Extension.ToLower()
    if ($OoxmlExts -contains $ext)  { return Get-OoxmlContent     -File $File }
    if ($BinaryExts -contains $ext) { return Get-BinaryXlsContent -File $File }
    return [PSCustomObject]@{
        HasChart=$false; HasImage=$false; HasSmartArt=$false
        ChartCount=0; ImageCount=0; DrawingCount=0; SmartArtCnt=0
        Details=[System.Collections.Generic.List[string]]::new()
        Error="Unrecognised extension: $ext"
    }
}

# ============================================================
#  SAFE COPY (handles name conflicts)
# ============================================================
function Copy-Safe {
    param([string]$Src, [string]$DestDir, [string]$Name)
    $dest = Join-Path $DestDir $Name
    if (Test-Path -LiteralPath $dest) {
        Write-Log "  Already exists in destination, skipping copy: '$Name'" -L WARN
        return $dest   # return existing path; do NOT create a second copy
    }
    Copy-Item -LiteralPath $Src -Destination $dest -Force -ErrorAction Stop
    return $dest
}

# ============================================================
#  FILE DISCOVERY
# ============================================================
Write-Log '' ; Write-Log -L SEP
Write-Log "FILE DISCOVERY  -  scanning '$FolderPath'" -L HEAD
Write-Log -L SEP

try {
    $allFiles = @(Get-ChildItem -LiteralPath $FolderPath -File -Recurse -ErrorAction Stop |
                  Where-Object { $AllExts -contains $_.Extension.ToLower() })
} catch {
    Write-Log "FATAL: Cannot enumerate files: $($_.Exception.Message)" -L ERROR; exit 1
}

$cTotal = $allFiles.Count
Write-Log "Found $cTotal spreadsheet file(s)." -L INFO

$allFiles | Group-Object { $_.Extension.ToLower() } | Sort-Object Name |
    ForEach-Object { Write-Log "  $($_.Name.PadRight(8)) : $($_.Count)" -L INFO }

if ($cTotal -eq 0) {
    Write-Log "Nothing to process. Exiting." -L WARN
    $Log.ToString() | Set-Content -LiteralPath $LogFile -Encoding UTF8
    "No files processed." | Set-Content -LiteralPath $ErrFile -Encoding UTF8
    exit 0
}

# ============================================================
#  FIRST PASS
# ============================================================
Write-Log '' ; Write-Log -L SEP
Write-Log "FIRST PASS  -  $cTotal file(s)" -L HEAD
Write-Log -L SEP

$Results = [System.Collections.Generic.List[PSCustomObject]]::new()

# ── Build a RunspacePool so analysis runs on up to $MaxThreads threads.
#    File I/O checks (exist + read access) are done sequentially first —
#    they are fast and avoid burning threads on files that cannot be read.
#    Only the CPU-bound analysis work is parallelised.
#    Copying files to output folders stays sequential (safe filesystem writes).

$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new(
    'Get-OoxmlContent',     ${function:Get-OoxmlContent}.ToString()))
$iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new(
    'Get-BinaryXlsContent', ${function:Get-BinaryXlsContent}.ToString()))

$pool = [RunspaceFactory]::CreateRunspacePool(1, $MaxThreads, $iss, $Host)
$pool.ApartmentState = 'MTA'
$pool.Open()

# Script block executed inside each runspace (functions pre-loaded via ISS above)
$WorkerScript = {
    param([string]$FilePath, [string]$Ext, [string[]]$OoxmlExts, [string[]]$BinaryExts)
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression            -ErrorAction SilentlyContinue
    $f = Get-Item -LiteralPath $FilePath -ErrorAction Stop
    if ($OoxmlExts  -contains $Ext) { return Get-OoxmlContent     -File $f }
    if ($BinaryExts -contains $Ext) { return Get-BinaryXlsContent -File $f }
    return [PSCustomObject]@{
        HasChart=$false; HasImage=$false; HasSmartArt=$false; HasShape=$false
        ChartCount=0; ImageCount=0; DrawingCount=0; ShapeCount=0; SmartArtCnt=0
        Details=[System.Collections.Generic.List[string]]::new()
        Error="Unsupported extension: $Ext"
    }
}

# Phase 1: pre-flight checks + submit to pool
$pending = [System.Collections.Generic.List[hashtable]]::new()

for ($i = 0; $i -lt $cTotal; $i++) {
    $file = $allFiles[$i]
    $idx  = $i + 1

    Write-Log ''
    Write-Log "[$idx/$cTotal] Pre-flight: $($file.Name)" -L HEAD
    Write-Log "  Path : $($file.FullName)" -L INFO
    Write-Log "  Size : $([Math]::Round($file.Length/1KB,2)) KB  |  Modified: $($file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))" -L INFO

    # Existence check
    if (-not (Test-Path -LiteralPath $file.FullName -PathType Leaf)) {
        $err = "File no longer exists at path."
        Log-Error "[$idx] $($file.FullName)" 'FILE_NOT_FOUND' $err
        $Results.Add([PSCustomObject]@{
            Index=$idx; File=$file; Category='review'
            HasChart=$false; HasImage=$false; HasShape=$false; ChartCount=0; ImageCount=0; ShapeCount=0; DrawingCount=0
            Details=@(); Error=$err; CopiedTo=$null
        })
        $cReview++; continue
    }

    # Read-access check
    $ts = $null
    try {
        $ts = [System.IO.File]::Open($file.FullName, [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        Write-Log "  Access OK." -L DBG
    } catch {
        $err = "Cannot open for reading: $($_.Exception.Message)"
        Log-Error "[$idx] $($file.FullName)" 'ACCESS_DENIED' $err
        $Results.Add([PSCustomObject]@{
            Index=$idx; File=$file; Category='review'
            HasChart=$false; HasImage=$false; HasShape=$false; ChartCount=0; ImageCount=0; ShapeCount=0; DrawingCount=0
            Details=@(); Error=$err; CopiedTo=$null
        })
        $cReview++; continue
    } finally {
        if ($ts) { try { $ts.Close(); $ts.Dispose() } catch {} }
    }

    # Submit analysis to pool
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($WorkerScript)
    [void]$ps.AddArgument($file.FullName)
    [void]$ps.AddArgument($file.Extension.ToLower())
    [void]$ps.AddArgument($OoxmlExts)
    [void]$ps.AddArgument($BinaryExts)
    $pending.Add(@{ PS=$ps; Handle=$ps.BeginInvoke(); File=$file; Index=$idx })
    Write-Log "  Submitted to analysis pool." -L DBG
}

Write-Log ''
Write-Log "Pre-flight done. Collecting $($pending.Count) analysis result(s)..." -L INFO

# Phase 2: collect results as threads complete, then copy (sequential)
$doneCount = 0
while ($pending.Count -gt 0) {
    $finished = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($job in $pending) {
        if ($job.Handle.IsCompleted) { $finished.Add($job) }
    }

    foreach ($job in $finished) {
        [void]$pending.Remove($job)
        $doneCount++

        $raw = $null
        try   { $raw = $job.PS.EndInvoke($job.Handle) }
        catch {
            $errMsg = "Runspace execution error: $($_.Exception.Message)"
            Log-Error "[$($job.Index)] $($job.File.Name)" 'RUNSPACE_ERROR' $errMsg
        }
        # Log any non-terminating errors from the runspace stream
        if ($job.PS.HadErrors) {
            foreach ($se in $job.PS.Streams.Error) {
                Write-Log "  Runspace stream error ($($job.File.Name)): $($se.Exception.Message)" -L WARN
            }
        }
        $job.PS.Dispose()

        # EndInvoke returns a PSDataCollection; unwrap the first element
        $a = if ($raw -and $raw.Count -gt 0) { $raw[0] } else { $null }

        Write-Log ''
        Write-Log "[$($job.Index)/$cTotal] Result ($doneCount of $($pending.Count + $doneCount) done): $($job.File.Name)" -L HEAD

        $rec = [PSCustomObject]@{
            Index      = $job.Index
            File       = $job.File
            Category   = 'review'
            HasChart   = $false; HasImage = $false; HasShape = $false
            ChartCount = 0;      ImageCount = 0;    ShapeCount = 0; DrawingCount = 0
            Details    = @()
            Error      = $null
            CopiedTo   = $null
        }

        if ($null -eq $a) {
            $rec.Error = "Analysis returned no result."
            $Results.Add($rec); $cReview++; continue
        }

        if ($a.Error) {
            $rec.Error = $a.Error
            Log-Error "[$($job.Index)] $($job.File.FullName)" 'ANALYSIS_ERROR' $a.Error
            $rec.Details = @($a.Details)
            $Results.Add($rec); $cReview++; continue
        }

        $rec.HasChart     = $a.HasChart
        $rec.HasImage     = $a.HasImage
        $rec.HasShape     = $a.HasShape
        $rec.ChartCount   = $a.ChartCount
        $rec.ImageCount   = $a.ImageCount
        $rec.ShapeCount   = $a.ShapeCount
        $rec.DrawingCount = $a.DrawingCount
        $rec.Details      = @($a.Details)

        $hasContent   = $a.HasChart -or $a.HasImage -or $a.HasSmartArt -or $a.HasShape
        $rec.Category = if ($hasContent) { 'Has content' } else { 'Does not have content' }

        Write-Log "  Charts   : $($a.ChartCount)"   -L INFO
        Write-Log "  Images   : $($a.ImageCount)"   -L INFO
        Write-Log "  Shapes   : $($a.ShapeCount)"   -L INFO
        Write-Log "  Drawings : $($a.DrawingCount)" -L INFO
        Write-Log "  SmartArt : $($a.SmartArtCnt)"  -L INFO
        $_lvl = if ($hasContent) { 'OK' } else { 'INFO' }
        Write-Log "  RESULT   : $($rec.Category)"   -L $_lvl
        foreach ($d in $a.Details) { Write-Log "    >> $d" -L DBG }

        # Copy to output (sequential — safe filesystem write)
        $targetDir = if ($hasContent) { $DirHas } else { $DirNo }
        try {
            $rec.CopiedTo = Copy-Safe -Src $job.File.FullName -DestDir $targetDir -Name $job.File.Name
            Write-Log "  Copied -> $($rec.CopiedTo)" -L OK
            if ($hasContent) { $cHas++ } else { $cNo++ }
        } catch {
            $rec.Error    = "Copy failed: $_"
            $rec.Category = 'review'
            Log-Error "[$($job.Index)] $($job.File.FullName)" 'COPY_FAILED' $rec.Error
            try {
                $rec.CopiedTo = Copy-Safe -Src $job.File.FullName -DestDir $DirReview -Name $job.File.Name
                Write-Log "  Fallback copy -> review" -L WARN
            } catch {
                Write-Log "  CRITICAL: Could not copy to any output: $_" -L ERROR
            }
            $cReview++
        }
        $Results.Add($rec)
    }

    if ($pending.Count -gt 0) { Start-Sleep -Milliseconds 150 }
}

$pool.Close()
$pool.Dispose()

# ============================================================
#  FIRST PASS SUMMARY
# ============================================================
Write-Log '' ; Write-Log -L SEP
Write-Log "FIRST PASS COMPLETE" -L HEAD
Write-Log "  Total              : $cTotal" -L INFO
Write-Log "  Has content        : $cHas"   -L OK
Write-Log "  Does not have content: $cNo"  -L INFO
$_lvl = if ($cReview -gt 0) { 'WARN' } else { 'INFO' }
Write-Log "  Review (errors)    : $cReview" -L $_lvl
Write-Log -L SEP

# ============================================================
#  SECOND PASS - QC VERIFICATION
# ============================================================
Write-Log '' ; Write-Log -L SEP
Write-Log "SECOND PASS  -  QC VERIFICATION" -L HEAD
Write-Log -L SEP

$qcFiles    = @($Results | Where-Object { $_.Category -ne 'review' -and $null -ne $_.CopiedTo })
$qcTotal    = $qcFiles.Count
$qcVerified = 0; $qcMismatch = 0; $qcErr = 0

Write-Log "QC files to check: $qcTotal" -L INFO

for ($q = 0; $q -lt $qcTotal; $q++) {
    $rec  = $qcFiles[$q]
    $qIdx = $q + 1
    Write-Log ''
    Write-Log "[QC $qIdx/$qcTotal] $($rec.File.Name)" -L HEAD
    Write-Log "  Original category: $($rec.Category)" -L INFO

    if (-not (Test-Path -LiteralPath $rec.CopiedTo -PathType Leaf)) {
        Write-Log "  QC ERROR: Copied file missing at '$($rec.CopiedTo)'" -L ERROR
        Log-Error "[QC] $($rec.File.Name)" 'QC_FILE_MISSING' "Not found: $($rec.CopiedTo)"
        $qcErr++; continue
    }

    $qcFile = Get-Item -LiteralPath $rec.CopiedTo -ErrorAction SilentlyContinue
    $qa     = $null
    try   { $qa = Invoke-Analysis -File $qcFile }
    catch {
        Write-Log "  QC ERROR: Analysis exception: $($_.Exception.Message)" -L ERROR
        Log-Error "[QC] $($rec.File.Name)" 'QC_ANALYSIS_EXCEPTION' $_.Exception.Message
        $qcErr++; continue
    }

    if ($qa.Error) {
        Write-Log "  QC ERROR: $($qa.Error)" -L ERROR
        Log-Error "[QC] $($rec.File.Name)" 'QC_ANALYSIS_ERROR' $qa.Error
        $qcErr++; continue
    }

    $qaHas    = $qa.HasChart -or $qa.HasImage -or $qa.HasSmartArt -or $qa.HasShape
    $origHas  = ($rec.Category -eq 'Has content')
    $qaLabel  = if ($qaHas) { 'Has content' } else { 'Does not have content' }

    Write-Log "  QC Charts  : $($qa.ChartCount)" -L INFO
    Write-Log "  QC Images  : $($qa.ImageCount)" -L INFO
    Write-Log "  QC Shapes  : $($qa.ShapeCount)" -L INFO
    Write-Log "  QC Result  : $qaLabel"          -L INFO

    if ($qaHas -ne $origHas) {
        $qcMismatch++
        Write-Log "  !! MISMATCH: was '$($rec.Category)' -> QC says '$qaLabel'" -L WARN
        Write-Log "  Moving to 'review'..." -L WARN
        Log-Error "[QC_MISMATCH] $($rec.File.Name)" 'QC_MISMATCH' `
            "Pass1='$($rec.Category)' QC='$qaLabel' Charts:$($rec.ChartCount)->$($qa.ChartCount) Images:$($rec.ImageCount)->$($qa.ImageCount) Shapes:$($rec.ShapeCount)->$($qa.ShapeCount)"

        $dest = $null
        try {
            $dest = Copy-Safe -Src $rec.CopiedTo -DestDir $DirReview -Name $rec.File.Name
            Remove-Item -LiteralPath $rec.CopiedTo -Force -ErrorAction SilentlyContinue
            Write-Log "  Moved to: $dest" -L WARN
            if ($origHas) { $cHas-- } else { $cNo-- }
            $cReview++
        } catch {
            Write-Log "  ERROR moving to review: $_" -L ERROR
        }
    } else {
        $qcVerified++
        Write-Log "  QC CONFIRMED: '$($rec.Category)'" -L OK
    }
}

# ============================================================
#  THIRD PASS - FOLDER DEDUPLICATION QC
#  Ensures no spreadsheet exists in more than one output folder.
#  Rule: if a file is in BOTH "Has content" AND "Does not have
#        content", remove it from "Does not have content".
# ============================================================
Write-Log '' ; Write-Log -L SEP
Write-Log "THIRD PASS  -  FOLDER DEDUPLICATION QC" -L HEAD
Write-Log -L SEP

function Get-FolderIndex {
    param([string]$Dir)
    $idx = @{}
    if (Test-Path -LiteralPath $Dir) {
        Get-ChildItem -LiteralPath $Dir -File | ForEach-Object { $idx[$_.Name.ToLower()] = $_.FullName }
    }
    return $idx
}

$idxHas    = Get-FolderIndex -Dir $DirHas
$idxNo     = Get-FolderIndex -Dir $DirNo
$idxReview = Get-FolderIndex -Dir $DirReview

$dedupRemoved = 0
$dedupClean   = 0

# Check every unique filename across all three folders
$allNames = ($idxHas.Keys + $idxNo.Keys + $idxReview.Keys) | Sort-Object -Unique

foreach ($name in $allNames) {
    $inHas    = $idxHas.ContainsKey($name)
    $inNo     = $idxNo.ContainsKey($name)
    $inReview = $idxReview.ContainsKey($name)

    $locations = @()
    if ($inHas)    { $locations += 'Has content' }
    if ($inNo)     { $locations += 'Does not have content' }
    if ($inReview) { $locations += 'review' }

    if ($locations.Count -gt 1) {
        Write-Log "  DUPLICATE DETECTED: '$name' found in: $($locations -join ', ')" -L WARN

        # Rule: if in both Has content and Does not have content -> remove from Does not have content
        if ($inHas -and $inNo) {
            try {
                Remove-Item -LiteralPath $idxNo[$name] -Force -ErrorAction Stop
                Write-Log "    Removed from 'Does not have content' (kept in 'Has content'): '$name'" -L WARN
                Log-Error "[DEDUP] $name" 'FOLDER_DUPLICATE' "Found in 'Has content' and 'Does not have content' - removed from 'Does not have content'"
                $cNo--
                $dedupRemoved++
            } catch {
                Write-Log "    ERROR removing '$name' from 'Does not have content': $($_.Exception.Message)" -L ERROR
                Log-Error "[DEDUP] $name" 'DEDUP_REMOVE_FAILED' $_.Exception.Message
            }
        }

        # If in Has content (or Does not have content) AND review, log it but leave review copy
        if ($inReview -and ($inHas -or $inNo)) {
            Write-Log "    NOTE: '$name' also exists in 'review' - review copy retained for inspection." -L WARN
        }
    } else {
        $dedupClean++
    }
}

Write-Log "  Deduplication complete. Clean: $dedupClean | Duplicates removed: $dedupRemoved" -L INFO
Write-Log -L SEP

# ============================================================
#  FINAL SUMMARY
# ============================================================
$ScriptEnd = Get-Date
$Duration  = $ScriptEnd - $ScriptStart

$tCharts   = ($Results | Where-Object { $_.Category -ne 'review' } | Measure-Object -Property ChartCount  -Sum).Sum
$tImages   = ($Results | Where-Object { $_.Category -ne 'review' } | Measure-Object -Property ImageCount  -Sum).Sum
$tShapes   = ($Results | Where-Object { $_.Category -ne 'review' } | Measure-Object -Property ShapeCount  -Sum).Sum
$tDrawings = ($Results | Where-Object { $_.Category -ne 'review' } | Measure-Object -Property DrawingCount -Sum).Sum

Write-Log '' ; Write-Log ('=' * 70) -L HEAD
Write-Log "FINAL SUMMARY" -L HEAD
Write-Log ('=' * 70) -L HEAD
Write-Log "  Completed   : $($ScriptEnd.ToString('yyyy-MM-dd HH:mm:ss'))" -L INFO
Write-Log "  Duration    : $($Duration.ToString('hh\:mm\:ss\.fff'))" -L INFO
Write-Log "  Run by      : $($env:USERDOMAIN)\$($env:USERNAME)" -L INFO
Write-Log ''
Write-Log "  --- Classification ---" -L HEAD
Write-Log "  Total files              : $cTotal"   -L INFO
Write-Log "  Has content              : $cHas"     -L OK
Write-Log "  Does not have content    : $cNo"      -L INFO
$_lvl = if ($cReview -gt 0) { 'WARN' } else { 'INFO' }
Write-Log "  Review (errors / QC fail): $cReview" -L $_lvl
Write-Log ''
Write-Log "  --- Content Totals (Pass 1) ---" -L HEAD
Write-Log "  Charts found             : $tCharts"   -L INFO
Write-Log "  Images found             : $tImages"   -L INFO
Write-Log "  Shapes found             : $tShapes"   -L INFO
Write-Log "  Drawing objects (total)  : $tDrawings" -L INFO
Write-Log ''
Write-Log "  --- QC Results ---" -L HEAD
Write-Log "  Files QC-checked         : $qcTotal"    -L INFO
Write-Log "  Confirmed correct        : $qcVerified" -L OK
$_lvl = if ($qcMismatch -gt 0) { 'WARN' } else { 'INFO' }
Write-Log "  Mismatches (-> review)   : $qcMismatch" -L $_lvl
$_lvl = if ($qcErr -gt 0) { 'WARN' } else { 'INFO' }
Write-Log "  QC errors                : $qcErr"      -L $_lvl
Write-Log ''
Write-Log "  --- Error Log ---" -L HEAD
$_lvl = if ($Errors.Count -gt 0) { 'WARN' } else { 'INFO' }
Write-Log "  Total errors recorded    : $($Errors.Count)" -L $_lvl
Write-Log ''
Write-Log "  --- Output ---" -L HEAD
Write-Log "  Has content dir          : $DirHas"    -L INFO
Write-Log "  No content dir           : $DirNo"     -L INFO
Write-Log "  Review dir               : $DirReview" -L INFO
Write-Log "  Log file                 : $LogFile"   -L INFO
Write-Log "  Errors file              : $ErrFile"   -L INFO
Write-Log ('=' * 70) -L HEAD

# ============================================================
#  WRITE OUTPUT FILES
# ============================================================
try {
    $Log.ToString() | Set-Content -LiteralPath $LogFile -Encoding UTF8 -ErrorAction Stop
    Write-Host "Log written  : $LogFile" -ForegroundColor Cyan
} catch {
    Write-Warning "Could not write log file: $($_.Exception.Message)"
}

try {
    if ($Errors.Count -gt 0) {
        $Errors | Set-Content -LiteralPath $ErrFile -Encoding UTF8 -ErrorAction Stop
    } else {
        'No errors recorded.' | Set-Content -LiteralPath $ErrFile -Encoding UTF8
    }
    Write-Host "Errors file  : $ErrFile" -ForegroundColor Cyan
} catch {
    Write-Warning "Could not write errors file: $($_.Exception.Message)"
}
