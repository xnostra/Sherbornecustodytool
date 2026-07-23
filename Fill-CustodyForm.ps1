<#
.SYNOPSIS
    IT Asset Custody Form Automation Tool
    Automatically extracts hardware specs and populates Excel tracking forms.

.DESCRIPTION
    This script extracts detailed hardware specifications from the local machine
    (brand, model, CPU, OS, RAM, GPU, screen size, storage) and automatically
    populates a formatted Excel (.xlsx) custody/tracking form.

    ZERO-DEPENDENCY / FULLY PORTABLE:
    - No external modules required (no ImportExcel, no Excel install)
    - Uses only built-in System.IO.Compression classes (Windows PowerShell 5.1+)
    - Copy entire folder to USB and run from any Windows computer
    - Admin elevation handled automatically

.PARAMETER TemplatePath
    Path to the template Excel file. Defaults to "custody form.xlsx" in script directory.

.PARAMETER OutputFolder
    Directory where completed forms are saved. Defaults to "Filled\" in script directory.

.EXAMPLE
    .\Fill-CustodyForm.ps1

    # Or specify custom paths:
    .\Fill-CustodyForm.ps1 -TemplatePath "C:\Templates\form.xlsx" -OutputFolder "C:\Output"

.NOTES
    Author: Sherborne Custody Tool Team
    Version: 3.9
    Requires: Windows 10/11, PowerShell 5.1+
    LastModified: 2026-07-23
    Dependencies: None (zero external dependencies)

.LINK
    https://github.com/xnostra/Sherbornecustodytool

#>

[CmdletBinding()]
param(
    [string]$TemplatePath = "",
    [string]$OutputFolder = "",
    [switch]$EmailForm  # if set, also uploads the finished form to OneDrive so the email-relay flow picks it up
)

$ErrorActionPreference = "Stop"

# Handle PSScriptRoot/PSCommandPath being empty when run via Invoke-Expression (e.g. the .bat
# launcher pipes this file's text into PowerShell instead of running it as a file, so that it
# still works on computers where running .ps1 FILES is disabled by policy).
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$scriptSelf = if ($PSCommandPath) { $PSCommandPath } else { Join-Path $scriptRoot 'Fill-CustodyForm.ps1' }

# Set defaults if not provided
if (-not $TemplatePath) {
    $TemplatePath = Join-Path $scriptRoot "custody form.xlsx"
}
if (-not $OutputFolder) {
    $OutputFolder = Join-Path $scriptRoot "Filled"
}
$ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
$GiBtoGB = 1.073741824

$ramStandards     = 1,2,3,4,6,8,12,16,24,32,48,64,96,128,192,256,384,512
$storageStandards = 16,32,64,120,128,240,250,256,320,480,500,512,640,750,960,1000,1024,1500,1920,2000,2048,3000,3840,4000,4096,6000,8000,10000,16000

# Every run's messages are also written to a log file next to the script, so if the window closes
# before you can read/copy an error, the full text is still sitting there afterward. On any error,
# the log auto-opens in Notepad so you don't have to go hunting for it.
$script:LogPath = Join-Path $scriptRoot 'CustodyForm.log'
$script:LogMaxBytes = 2MB  # after months of daily use this would otherwise grow forever

# If the log has grown past the cap, trim it down to roughly its last quarter before this run adds
# more - keeps recent history for troubleshooting without letting the file grow unbounded.
try {
    $existingLog = Get-Item -LiteralPath $script:LogPath -ErrorAction Stop
    if ($existingLog.Length -gt $script:LogMaxBytes) {
        # @() forces an array even if the file happens to be a single line, so the range index
        # below always works against a real line collection rather than being treated as one string.
        $allLines = @(Get-Content -LiteralPath $script:LogPath -ErrorAction Stop)
        if ($allLines.Count -gt 4) {
            $keepFrom = [Math]::Max(0, $allLines.Count - [Math]::Floor($allLines.Count / 4))
            $trimmed = $allLines[$keepFrom..($allLines.Count - 1)]
            Set-Content -LiteralPath $script:LogPath -Value $trimmed -ErrorAction Stop
            Add-Content -LiteralPath $script:LogPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  [Info] Log file exceeded $($script:LogMaxBytes / 1MB)MB - older entries were trimmed to keep it from growing unbounded." -ErrorAction Stop
        }
    }
} catch { }

function Write-Log {
    param([string]$Message)
    try { Add-Content -LiteralPath $script:LogPath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message" -ErrorAction Stop } catch { }
}

function Write-Status {
    param([string]$Message, [ValidateSet('Info', 'Success', 'Warning', 'Error')] [string]$Level = 'Info')
    $colors = @{ 'Info' = 'Cyan'; 'Success' = 'Green'; 'Warning' = 'Yellow'; 'Error' = 'Red' }
    Write-Host $Message -ForegroundColor $colors[$Level]
    Write-Log "[$Level] $Message"
}

function Open-LogOnError {
    try { Start-Process notepad.exe -ArgumentList $script:LogPath -ErrorAction Stop } catch { }
}

trap {
    $msg = $_.Exception.Message
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($msg -match 'Access.*denied|UnauthorizedAccessException|requires elevation' -and -not $isAdmin) {
        Write-Status "Admin rights needed - requesting elevation..." -Level Warning
        # Encoded command instead of "-File $scriptSelf" so the elevated relaunch still works even
        # if running script FILES is disabled by policy on this machine. Calling a .ps1 file with
        # "&" is still subject to Execution Policy even from inside an -EncodedCommand, so instead
        # the inner command reads this file's own text and runs it as a scriptblock (not a file) -
        # that is never subject to Execution Policy, and normal param() binding still works.
        # Preserve -EmailForm across the elevation relaunch - otherwise a run started with emailing
        # requested would silently lose that after elevating and finish without emailing at all.
        $elevEmailArg = if ($EmailForm) { ' -EmailForm' } else { '' }
        $elevInner   = "& ([scriptblock]::Create((Get-Content -Raw -LiteralPath '$scriptSelf'))) -TemplatePath '$TemplatePath' -OutputFolder '$OutputFolder'$elevEmailArg"
        $elevEncoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($elevInner))
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $elevEncoded"
        exit 0
    }
    Write-Status "`nERROR:" -Level Error
    Write-Status $_.Exception.Message -Level Error
    Write-Status $_.InvocationInfo.PositionMessage -Level Warning
    Open-LogOnError
    Read-Host "Press Enter to close"
    exit 1
}

# Auto-allow running this tool's scripts: try CurrentUser first (no admin needed), then Process
# scope as a fallback. This only relaxes the policy for future runs / this session - it does not
# touch machine-wide policy. If a GPO enforces a stricter policy above these scopes, Set-ExecutionPolicy
# throws a SecurityException that -ErrorAction can't suppress (it's a terminating error, not the
# kind -ErrorAction SilentlyContinue catches) - so each call is wrapped in try/catch instead. Either
# way this is a no-op on a GPO-locked machine, which is exactly why the .bat launcher pipes the
# script instead of running it as a file (that path works regardless of what this can or can't change).
try { Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop } catch { }
try { Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop } catch { }
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

if (-not (Test-Path $TemplatePath)) {
    Write-Status "ERROR: Template not found at: $TemplatePath" -Level Error
    Read-Host "Press Enter to close"
    exit 1
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

function ColLetterToNum {
    param([string]$col)
    $n = 0
    foreach ($ch in $col.ToUpper().ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
    return $n
}

function Set-CellInlineValue {
    param($RowNode, $XmlDoc, $NsMgr, [string]$CellRef, [string]$Value)
    $col = ($CellRef -replace '\d', '')
    $cell = $RowNode.SelectSingleNode("s:c[@r='$CellRef']", $NsMgr)
    if (-not $cell) {
        $cell = $XmlDoc.CreateElement('c', $ns)
        $cell.SetAttribute('r', $CellRef)
        $inserted = $false
        foreach ($existing in $RowNode.SelectNodes('s:c', $NsMgr)) {
            $existingCol = ($existing.GetAttribute('r') -replace '\d', '')
            if ((ColLetterToNum $existingCol) -gt (ColLetterToNum $col)) {
                $RowNode.InsertBefore($cell, $existing) | Out-Null
                $inserted = $true
                break
            }
        }
        if (-not $inserted) { $RowNode.AppendChild($cell) | Out-Null }
    } else {
        $cell.RemoveAttribute('t')
        while ($cell.HasChildNodes) { $cell.RemoveChild($cell.FirstChild) | Out-Null }
    }
    $cell.SetAttribute('t', 'inlineStr')
    $isNode = $XmlDoc.CreateElement('is', $ns)
    $tNode = $XmlDoc.CreateElement('t', $ns)
    $tNode.SetAttribute('xml:space', 'preserve')
    $tNode.InnerText = $Value
    $isNode.AppendChild($tNode) | Out-Null
    $cell.AppendChild($isNode) | Out-Null
}

function Get-StyleVariantIndex {
    param($CellXfsNode, $StylesXml, $NsMgr, [int]$OriginalIndex, [string]$Mode, [hashtable]$Cache)
    $cacheKey = "$OriginalIndex-$Mode"
    if ($Cache.ContainsKey($cacheKey)) { return $Cache[$cacheKey] }
    $xfNodes = $CellXfsNode.SelectNodes('s:xf', $NsMgr)
    if ($OriginalIndex -lt $xfNodes.Count) { $newXf = $xfNodes.Item($OriginalIndex).CloneNode($true) } else {
        $newXf = $StylesXml.CreateElement('xf', $ns)
        $newXf.SetAttribute('numFmtId', '0'); $newXf.SetAttribute('fontId', '0'); $newXf.SetAttribute('fillId', '0'); $newXf.SetAttribute('borderId', '0'); $newXf.SetAttribute('xfId', '0')
    }
    $newXf.SetAttribute('applyAlignment', '1')
    $align = $newXf.SelectSingleNode('s:alignment', $NsMgr)
    if (-not $align) { $align = $StylesXml.CreateElement('alignment', $ns); $newXf.AppendChild($align) | Out-Null }
    $align.SetAttribute('horizontal', 'left'); $align.SetAttribute('vertical', 'center')
    if ($Mode -eq 'wrap') { $align.SetAttribute('wrapText', '1') } elseif ($Mode -eq 'shrink') { $align.SetAttribute('shrinkToFit', '1') }
    $newIndex = $xfNodes.Count
    $CellXfsNode.AppendChild($newXf) | Out-Null
    $CellXfsNode.SetAttribute('count', ([int]$CellXfsNode.GetAttribute('count') + 1).ToString())
    $Cache[$cacheKey] = $newIndex
    return $newIndex
}

function Read-ZipXml {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    $xmlDoc = New-Object System.Xml.XmlDocument
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $xmlDoc.LoadXml($reader.ReadToEnd()); $reader.Close()
    return $xmlDoc
}

function Write-ZipXml {
    param($Zip, [string]$EntryName, $XmlDoc)
    $Zip.GetEntry($EntryName).Delete()
    $newEntry = $Zip.CreateEntry($EntryName); $stream = $newEntry.Open()
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $writerSettings = New-Object System.Xml.XmlWriterSettings; $writerSettings.Encoding = $utf8NoBom
    $writer = [System.Xml.XmlWriter]::Create($stream, $writerSettings)
    $XmlDoc.Save($writer); $writer.Close(); $stream.Close()
}

function Set-PrintLayout {
    param($SheetXml, $NsMgr, $RootNode)
    $order = @('sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','drawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst')
    function Get-OrAddElement { param([string]$TagName); $node = $RootNode.SelectSingleNode("s:$TagName", $NsMgr); if (-not $node) { $node = $SheetXml.CreateElement($TagName, $ns); $tagIndex = [array]::IndexOf($order, $TagName); $target = $null; foreach ($child in $RootNode.ChildNodes) { if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }; $childIndex = [array]::IndexOf($order, $child.LocalName); if ($childIndex -ge 0 -and $childIndex -gt $tagIndex) { $target = $child; break } }; if ($target) { $RootNode.InsertBefore($node, $target) | Out-Null } else { $RootNode.AppendChild($node) | Out-Null } }; return $node }
    $sheetPr = Get-OrAddElement 'sheetPr'
    $pageSetUpPr = $sheetPr.SelectSingleNode('s:pageSetUpPr', $NsMgr)
    if (-not $pageSetUpPr) { $pageSetUpPr = $SheetXml.CreateElement('pageSetUpPr', $ns); $sheetPr.AppendChild($pageSetUpPr) | Out-Null }
    $pageSetUpPr.SetAttribute('fitToPage', '1')
    $printOptions = Get-OrAddElement 'printOptions'; $printOptions.SetAttribute('horizontalCentered', '1')
    $pageMargins = Get-OrAddElement 'pageMargins'
    $pageMargins.SetAttribute('left', '0.5'); $pageMargins.SetAttribute('right', '0.5'); $pageMargins.SetAttribute('top', '0.75'); $pageMargins.SetAttribute('bottom', '0.75'); $pageMargins.SetAttribute('header', '0.3'); $pageMargins.SetAttribute('footer', '0.3')
    $pageSetup = Get-OrAddElement 'pageSetup'
    $pageSetup.SetAttribute('fitToWidth', '1'); $pageSetup.SetAttribute('orientation', 'portrait')
}

function Get-NearestStandardSize {
    param([double]$ActualGB, [double[]]$Standards, [string]$Mode = 'ceiling')
    if ($Mode -eq 'nearest') { $best = $Standards[0]; $bestDiff = [Math]::Abs($Standards[0] - $ActualGB); foreach ($s in $Standards) { $diff = [Math]::Abs($s - $ActualGB); if ($diff -lt $bestDiff) { $bestDiff = $diff; $best = $s } }; return $best } else { foreach ($s in $Standards) { if ($ActualGB -le ($s * 1.04)) { return $s } }; return $Standards[-1] }
}

function Format-StorageLabel { param([double]$SizeGB); if ($SizeGB -ge 1000 -and ($SizeGB % 1000) -eq 0) { return "$([int]($SizeGB/1000))TB" } elseif ($SizeGB -ge 1000) { return "$([Math]::Round($SizeGB/1000,1))TB" } else { return "${SizeGB}GB" } }

function Fill-Workbook {
    param([string]$Path, [hashtable]$Values, [string]$SheetName, [string[]]$WrapCells = @(), [string[]]$ShrinkCells = @(), [hashtable]$RowHeights = @{})
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $wbXml = Read-ZipXml -Zip $zip -EntryName "xl/workbook.xml"
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
        $nsMgr.AddNamespace('s', $ns); $nsMgr.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $sheetNode = $wbXml.SelectSingleNode("//s:sheets/s:sheet[@name='$SheetName']", $nsMgr)
        if (-not $sheetNode) { throw "Sheet '$SheetName' not found" }
        $rId = $sheetNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $relsXml = Read-ZipXml -Zip $zip -EntryName "xl/_rels/workbook.xml.rels"
        $relNode = $relsXml.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rId']")
        $target = $relNode.GetAttribute('Target') -replace '^/', ''
        if ($target -notmatch '^xl/') { $target = "xl/$target" }
        $sheetXml = Read-ZipXml -Zip $zip -EntryName $target
        $sheetNsMgr = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable); $sheetNsMgr.AddNamespace('s', $ns)
        Set-PrintLayout -SheetXml $sheetXml -NsMgr $sheetNsMgr -RootNode $sheetXml.DocumentElement
        $stylesXml = Read-ZipXml -Zip $zip -EntryName "xl/styles.xml"
        $stylesNsMgr = New-Object System.Xml.XmlNamespaceManager($stylesXml.NameTable); $stylesNsMgr.AddNamespace('s', $ns)
        $cellXfsNode = $stylesXml.SelectSingleNode("//s:styleSheet/s:cellXfs", $stylesNsMgr)
        $styleCache = @{}
        foreach ($cellRef in $Values.Keys) {
            $rowNum = ($cellRef -replace '\D', '')
            $row = $sheetXml.SelectSingleNode("//s:sheetData/s:row[@r='$rowNum']", $sheetNsMgr)
            if (-not $row) { continue }
            Set-CellInlineValue -RowNode $row -XmlDoc $sheetXml -NsMgr $sheetNsMgr -CellRef $cellRef -Value $Values[$cellRef]
        }
        foreach ($cellRef in $Values.Keys) {
            $rowNum = ($cellRef -replace '\D', '')
            $row = $sheetXml.SelectSingleNode("//s:sheetData/s:row[@r='$rowNum']", $sheetNsMgr)
            if (-not $row) { continue }
            $cell = $row.SelectSingleNode("s:c[@r='$cellRef']", $sheetNsMgr)
            if (-not $cell) { continue }
            $origIdx = if ($cell.GetAttribute('s')) { [int]$cell.GetAttribute('s') } else { 0 }
            $mode = if ($WrapCells -contains $cellRef) { 'wrap' } elseif ($ShrinkCells -contains $cellRef) { 'shrink' } else { 'align' }
            $newIdx = Get-StyleVariantIndex -CellXfsNode $cellXfsNode -StylesXml $stylesXml -NsMgr $stylesNsMgr -OriginalIndex $origIdx -Mode $mode -Cache $styleCache
            $cell.SetAttribute('s', $newIdx)
        }
        foreach ($rowNum in $RowHeights.Keys) {
            $row = $sheetXml.SelectSingleNode("//s:sheetData/s:row[@r='$rowNum']", $sheetNsMgr)
            if ($row) { $row.SetAttribute('ht', [string]$RowHeights[$rowNum]); $row.SetAttribute('customHeight', '1') }
        }
        Write-ZipXml -Zip $zip -EntryName $target -XmlDoc $sheetXml
        Write-ZipXml -Zip $zip -EntryName "xl/styles.xml" -XmlDoc $stylesXml
    } finally { $zip.Dispose() }
}

function Add-SignatureImage {
    # Embeds a drawn signature (PNG bytes) into a single cell of the given sheet, as a floating
    # picture anchored to that cell - same overall zip/XML-editing approach as Fill-Workbook, just
    # touching the drawing/media parts instead of the sheet's cell values. Cell-anchor math (col/row
    # are 0-based in DrawingML, vs 1-based/letter-based in normal cell refs) is the only tricky part.
    param([string]$Path, [byte[]]$PngBytes, [string]$SheetName, [string]$CellRef)
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $wbXml = Read-ZipXml -Zip $zip -EntryName "xl/workbook.xml"
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
        $nsMgr.AddNamespace('s', $ns); $nsMgr.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $sheetNode = $wbXml.SelectSingleNode("//s:sheets/s:sheet[@name='$SheetName']", $nsMgr)
        if (-not $sheetNode) { throw "Sheet '$SheetName' not found" }
        $rId = $sheetNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $relsXml = Read-ZipXml -Zip $zip -EntryName "xl/_rels/workbook.xml.rels"
        $relNode = $relsXml.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rId']")
        $sheetTarget = $relNode.GetAttribute('Target') -replace '^/', ''
        if ($sheetTarget -notmatch '^xl/') { $sheetTarget = "xl/$sheetTarget" }
        $sheetFileName = [System.IO.Path]::GetFileName($sheetTarget)

        # Find (or fail gracefully - shouldn't happen with the known template) the sheet's own
        # drawing relationship, which points at the drawingN.xml that holds the logo etc.
        $sheetRelsPath = "xl/worksheets/_rels/$sheetFileName.rels"
        $sheetRelsEntry = $zip.GetEntry($sheetRelsPath)
        if (-not $sheetRelsEntry) { throw "No drawing relationship found for $sheetFileName - can't embed signature" }
        $sheetRelsXml = Read-ZipXml -Zip $zip -EntryName $sheetRelsPath
        $drawingRel = $sheetRelsXml.SelectSingleNode("//*[local-name()='Relationship' and contains(@Type,'/drawing')]")
        if (-not $drawingRel) { throw "No drawing part linked from $sheetFileName - can't embed signature" }
        $drawingTarget = $drawingRel.GetAttribute('Target') -replace '^\.\./', 'xl/' -replace '^/', ''
        if ($drawingTarget -notmatch '^xl/') { $drawingTarget = "xl/$drawingTarget" }
        $drawingFileName = [System.IO.Path]::GetFileName($drawingTarget)

        # Add the PNG itself as a new media part with a name that won't collide with existing ones.
        $mediaIndex = 1
        while ($zip.GetEntry("xl/media/signature$mediaIndex.png")) { $mediaIndex++ }
        $mediaEntryName = "xl/media/signature$mediaIndex.png"
        $mediaEntry = $zip.CreateEntry($mediaEntryName)
        $mediaStream = $mediaEntry.Open()
        $mediaStream.Write($PngBytes, 0, $PngBytes.Length)
        $mediaStream.Close()

        # Register the new image in the drawing's own rels file, picking an rId that isn't taken.
        $drawingRelsPath = "xl/drawings/_rels/$drawingFileName.rels"
        $drawingRelsXml = $null
        $drawingRelsIsNew = $false
        if ($zip.GetEntry($drawingRelsPath)) {
            $drawingRelsXml = Read-ZipXml -Zip $zip -EntryName $drawingRelsPath
        } else {
            $drawingRelsIsNew = $true
            $drawingRelsXml = New-Object System.Xml.XmlDocument
            $drawingRelsXml.LoadXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
        }
        $existingIds = @($drawingRelsXml.SelectNodes("//*[local-name()='Relationship']") | ForEach-Object { $_.GetAttribute('Id') })
        $n = 1
        while ($existingIds -contains "rId$n") { $n++ }
        $newRelId = "rId$n"
        $relRoot = $drawingRelsXml.DocumentElement
        $newRel = $drawingRelsXml.CreateElement('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships')
        $newRel.SetAttribute('Id', $newRelId)
        $newRel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
        $newRel.SetAttribute('Target', "../media/signature$mediaIndex.png")
        $relRoot.AppendChild($newRel) | Out-Null
        if ($drawingRelsIsNew) {
            $newEntry = $zip.CreateEntry($drawingRelsPath)
            $stream = $newEntry.Open()
            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
            $writerSettings = New-Object System.Xml.XmlWriterSettings; $writerSettings.Encoding = $utf8NoBom
            $writer = [System.Xml.XmlWriter]::Create($stream, $writerSettings)
            $drawingRelsXml.Save($writer); $writer.Close(); $stream.Close()
        } else {
            Write-ZipXml -Zip $zip -EntryName $drawingRelsPath -XmlDoc $drawingRelsXml
        }

        # Anchor the picture to the target cell (e.g. "C35") - DrawingML columns/rows are 0-based,
        # so a normal 1-based cell ref needs converting; "to" is one column/row past "from" so the
        # picture fills exactly that single cell.
        $colLetters = ($CellRef -replace '\d', '')
        $rowNum = [int]($CellRef -replace '\D', '')
        $colIndex0 = (ColLetterToNum $colLetters) - 1
        $rowIndex0 = $rowNum - 1

        $xdrNs = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
        $aNs = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        $drawingXml = Read-ZipXml -Zip $zip -EntryName $drawingTarget
        $drawingNsMgr = New-Object System.Xml.XmlNamespaceManager($drawingXml.NameTable)
        $drawingNsMgr.AddNamespace('xdr', $xdrNs); $drawingNsMgr.AddNamespace('a', $aNs)
        $wsDr = $drawingXml.DocumentElement

        # If a signature picture was already embedded here from a previous run of this function on
        # the same file (shouldn't normally happen, but avoids ever double-stacking two signatures
        # on top of each other if this were ever called twice), remove the old one first.
        $existingSig = $wsDr.SelectSingleNode("xdr:twoCellAnchor[xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name='Signature Image']", $drawingNsMgr)
        if ($existingSig) { $wsDr.RemoveChild($existingSig) | Out-Null }

        $anchor = $drawingXml.CreateElement('xdr', 'twoCellAnchor', $xdrNs)
        $anchor.SetAttribute('editAs', 'oneCell')

        $from = $drawingXml.CreateElement('xdr', 'from', $xdrNs)
        $fromCol = $drawingXml.CreateElement('xdr', 'col', $xdrNs); $fromCol.InnerText = "$colIndex0"
        $fromColOff = $drawingXml.CreateElement('xdr', 'colOff', $xdrNs); $fromColOff.InnerText = '19050'
        $fromRow = $drawingXml.CreateElement('xdr', 'row', $xdrNs); $fromRow.InnerText = "$rowIndex0"
        $fromRowOff = $drawingXml.CreateElement('xdr', 'rowOff', $xdrNs); $fromRowOff.InnerText = '9525'
        $from.AppendChild($fromCol) | Out-Null; $from.AppendChild($fromColOff) | Out-Null
        $from.AppendChild($fromRow) | Out-Null; $from.AppendChild($fromRowOff) | Out-Null

        $to = $drawingXml.CreateElement('xdr', 'to', $xdrNs)
        $toCol = $drawingXml.CreateElement('xdr', 'col', $xdrNs); $toCol.InnerText = "$($colIndex0 + 1)"
        $toColOff = $drawingXml.CreateElement('xdr', 'colOff', $xdrNs); $toColOff.InnerText = '-19050'
        $toRow = $drawingXml.CreateElement('xdr', 'row', $xdrNs); $toRow.InnerText = "$($rowIndex0 + 1)"
        $toRowOff = $drawingXml.CreateElement('xdr', 'rowOff', $xdrNs); $toRowOff.InnerText = '-9525'
        $to.AppendChild($toCol) | Out-Null; $to.AppendChild($toColOff) | Out-Null
        $to.AppendChild($toRow) | Out-Null; $to.AppendChild($toRowOff) | Out-Null

        $pic = $drawingXml.CreateElement('xdr', 'pic', $xdrNs)
        $nvPicPr = $drawingXml.CreateElement('xdr', 'nvPicPr', $xdrNs)
        $cNvPr = $drawingXml.CreateElement('xdr', 'cNvPr', $xdrNs)
        $cNvPr.SetAttribute('id', '200'); $cNvPr.SetAttribute('name', 'Signature Image')
        $cNvPicPr = $drawingXml.CreateElement('xdr', 'cNvPicPr', $xdrNs)
        $picLocks = $drawingXml.CreateElement('a', 'picLocks', $aNs); $picLocks.SetAttribute('noChangeAspect', '1')
        $cNvPicPr.AppendChild($picLocks) | Out-Null
        $nvPicPr.AppendChild($cNvPr) | Out-Null; $nvPicPr.AppendChild($cNvPicPr) | Out-Null

        $blipFill = $drawingXml.CreateElement('xdr', 'blipFill', $xdrNs)
        $blip = $drawingXml.CreateElement('a', 'blip', $aNs)
        $blip.SetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', $newRelId)
        $stretch = $drawingXml.CreateElement('a', 'stretch', $aNs)
        $fillRect = $drawingXml.CreateElement('a', 'fillRect', $aNs)
        $stretch.AppendChild($fillRect) | Out-Null
        $blipFill.AppendChild($blip) | Out-Null; $blipFill.AppendChild($stretch) | Out-Null

        $spPr = $drawingXml.CreateElement('xdr', 'spPr', $xdrNs)
        $prstGeom = $drawingXml.CreateElement('a', 'prstGeom', $aNs); $prstGeom.SetAttribute('prst', 'rect')
        $avLst = $drawingXml.CreateElement('a', 'avLst', $aNs)
        $prstGeom.AppendChild($avLst) | Out-Null
        $spPr.AppendChild($prstGeom) | Out-Null

        $pic.AppendChild($nvPicPr) | Out-Null; $pic.AppendChild($blipFill) | Out-Null; $pic.AppendChild($spPr) | Out-Null

        $clientData = $drawingXml.CreateElement('xdr', 'clientData', $xdrNs)

        $anchor.AppendChild($from) | Out-Null; $anchor.AppendChild($to) | Out-Null
        $anchor.AppendChild($pic) | Out-Null; $anchor.AppendChild($clientData) | Out-Null
        $wsDr.AppendChild($anchor) | Out-Null

        Write-ZipXml -Zip $zip -EntryName $drawingTarget -XmlDoc $drawingXml
    } finally { $zip.Dispose() }
}

$defaultUser = $env:USERNAME
$dateIssued  = Get-Date -Format "dd-MMM-yyyy"
# Safe defaults in case this is NOT a computer/laptop (skips auto-detection entirely) and the
# category/description boxes in the popup form are left blank - never leave these undefined.
$itemCategory = "Other Equipment"
$itemDescription = ""

# Ask first whether this is even a computer/laptop - hardware auto-detection only makes sense when
# the tool is running ON the device being handed over. For other items (clickers, projectors,
# screens, etc.) you're filling this out from your OWN computer, so everything must be typed in
# by hand instead of auto-detected from this machine's own hardware.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$isComputerResult = [System.Windows.Forms.MessageBox]::Show(
    "Is this a computer or laptop (the tool is running ON the device being handed over)?`n`nChoose No for other items - projectors, clickers, screens, etc. - where you'll type the details in yourself.",
    'IT Asset Custody Form Tool',
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
)
$isThisAComputer = ($isComputerResult -eq [System.Windows.Forms.DialogResult]::Yes)

if ($isThisAComputer) {

Write-Status "Detecting hardware..." -Level Info
$cs   = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$enc  = Get-CimInstance Win32_SystemEnclosure
$cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
$os   = Get-CimInstance Win32_OperatingSystem

$laptopTypes = 8,9,10,11,12,14,18,21
$isLaptop = ($enc.ChassisTypes | Where-Object { $laptopTypes -contains $_ }).Count -gt 0
$itemCategory = if ($isLaptop) { "Laptop/Notebook" } else { "Desktop Computer" }

$modelName = $cs.Model; $sysSku = $cs.SystemSKUNumber
$rawManufacturer = $cs.Manufacturer.Trim()

# Placeholder/junk values seen on generic, misconfigured, or white-box motherboards - never show
# these as the "model" (or, further below, as the "manufacturer").
$junkPatterns = 'System Product Name|To Be Filled|Default string|O\.?E\.?M\.?|Not Applicable|None|N/A|Not Specified|Unknown'

# Manufacturer field often reports the full legal entity name rather than the brand name everyone
# actually recognizes (e.g. "Dell Inc." vs "Dell", "ASUSTeK COMPUTER INC." vs "ASUS") - normalize
# the common ones to their everyday brand name so the result reads naturally.
$manufacturer = switch -Regex ($rawManufacturer) {
    'LENOVO'                      { 'Lenovo'; break }
    'Hewlett-Packard|^HP$'        { 'HP'; break }
    'Dell'                        { 'Dell'; break }
    'ASUSTeK'                     { 'ASUS'; break }
    'Acer'                        { 'Acer'; break }
    'Microsoft'                   { 'Microsoft'; break }
    'Apple'                       { 'Apple'; break }
    'Toshiba|Dynabook'            { 'Toshiba'; break }
    'Samsung'                     { 'Samsung'; break }
    'Gigabyte'                    { 'Gigabyte'; break }
    'MSI|Micro-Star'              { 'MSI'; break }
    default                       { $rawManufacturer }
}
if ($manufacturer -match $junkPatterns) { $manufacturer = 'Unknown brand' }

# Build a clean, human-friendly "Brand Model" name that works across brands, then keep the exact
# system model number/SKU in parentheses afterward for asset records - not just for Lenovo.
#
# Different brands expose the friendly name differently:
#  - Lenovo: Model is a bare code (e.g. "21SX001TGR"); the friendly name is embedded in
#    SystemSKUNumber instead (e.g. "LENOVO_MT_21SX_BU_THINK_FM_THINKPAD E14 GEN 7" -> everything
#    after "_FM_").
#  - Dell, HP, Microsoft Surface, Acer, Asus, etc.: Model is usually already friendly on its own
#    (e.g. "Latitude 5420", "HP EliteBook 840 G8", "Surface Laptop 4") - just needs light cleanup
#    (removing a duplicated brand name, trimming placeholder junk).
#  - Generic/white-box/OEM builds: Model is often meaningless placeholder text (e.g. "System
#    Product Name", "To Be Filled By O.E.M.") - detect and fall back to something honest instead
#    of showing garbage.
$friendlyModel = $null
$modelWasDeduped = $false  # true when $friendlyModel came from Model with a duplicated brand prefix stripped off

if ($rawManufacturer -match 'LENOVO' -and $sysSku -match '_FM_(.+)$') {
    $friendlyModel = $Matches[1].Trim()
    # Title-case it (e.g. "THINKPAD E14 GEN 7" -> "ThinkPad E14 Gen 7") without breaking existing
    # mixed-case product codes - only reformat words that are all-uppercase in the source string.
    $friendlyModel = ($friendlyModel -split ' ' | ForEach-Object {
        if ($_ -cmatch '^[A-Z0-9]+$' -and $_.Length -gt 1) {
            if ($_ -eq 'THINKPAD') { 'ThinkPad' }
            elseif ($_ -eq 'THINKCENTRE') { 'ThinkCentre' }
            elseif ($_ -eq 'THINKBOOK') { 'ThinkBook' }
            elseif ($_ -eq 'IDEAPAD') { 'IdeaPad' }
            elseif ($_ -eq 'GEN') { 'Gen' }
            elseif ($_ -match '^\d') { $_ }  # keep model numbers like "E14" or "21SX" as-is
            else { (Get-Culture).TextInfo.ToTitleCase($_.ToLower()) }
        } else { $_ }
    }) -join ' '
} elseif ($modelName -and $modelName.Trim() -ne '' -and $modelName.Trim() -notmatch $junkPatterns) {
    $cleanModel = $modelName.Trim()
    # Some vendors (notably HP, ASUS) duplicate the brand name inside Model itself, e.g. Manufacturer
    # "HP" + Model "HP EliteBook 840 G8" - strip the leading duplicate so it doesn't show twice.
    # Checked against both the normalized brand name and the raw manufacturer string, since the
    # duplication in Model could use either form.
    foreach ($prefix in @($manufacturer, $rawManufacturer) | Select-Object -Unique) {
        if ($prefix -and $cleanModel -match "^$([regex]::Escape($prefix))\s+") {
            $cleanModel = $cleanModel -replace "^$([regex]::Escape($prefix))\s+", ''
            $modelWasDeduped = $true
            break
        }
    }
    $friendlyModel = $cleanModel
}

$brandModel = if ($friendlyModel) { "$manufacturer $friendlyModel".Trim() } else { "$manufacturer (model not reported by this device)".Trim() }
# Keep the exact system model number available for asset records, distinct from the friendly name -
# skip it if it's junk, if it's the same text already shown, or if $friendlyModel came from Model
# itself with a duplicated brand prefix stripped off (that's the SAME underlying value, just with
# "HP "/"ASUS " trimmed - showing it again in parentheses would just repeat the model, e.g.
# "HP EliteBook 840 G8 (HP EliteBook 840 G8)").
if ($modelName -and $modelName.Trim() -ne '' -and $modelName.Trim() -notmatch $junkPatterns -and $modelName.Trim() -ne $friendlyModel -and -not $modelWasDeduped) {
    $brandModel += " ($($modelName.Trim()))"
}

$cpuName = $cpu.Name -replace '\s+', ' '
$cpuGen = $null
if ($cpuName -match 'i[3579]-(\d{4,5})') { $num = $Matches[1]; $cpuGen = if ($num.Length -eq 5) { $num.Substring(0,2) } else { $num.Substring(0,1) } }
$cpuFull = if ($cpuGen) { "$cpuName (Gen $cpuGen)" } else { $cpuName }

$osInfo = "$($os.Caption) $($os.OSArchitecture)" -replace 'Microsoft ', ''
$rawRamGB = $cs.TotalPhysicalMemory / 1GB
$ramGB = Get-NearestStandardSize -ActualGB $rawRamGB -Standards $ramStandards -Mode 'ceiling'

$gpuNames = Get-CimInstance Win32_VideoController | Where-Object { $_.Name } | Select-Object -ExpandProperty Name
$gpuInfo = if ($gpuNames) { ($gpuNames -join '; ') } else { "Unknown" }

$screenSize = "Unknown"
try {
    $mon = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams -ErrorAction Stop | Select-Object -First 1
    if ($mon -and $mon.MaxHorizontalImageSize -gt 0 -and $mon.MaxVerticalImageSize -gt 0) {
        $diagCm = [Math]::Sqrt([Math]::Pow($mon.MaxHorizontalImageSize,2) + [Math]::Pow($mon.MaxVerticalImageSize,2))
        $screenSize = "$([Math]::Round($diagCm / 2.54, 1))in"
    }
} catch { }

$storageInfo = "Unknown"
try {
    $osDisk = Get-Disk -ErrorAction Stop | Where-Object { $_.IsBoot -or $_.IsSystem } | Select-Object -First 1
    if (-not $osDisk) { $osDisk = Get-Disk -ErrorAction Stop | Sort-Object Number | Select-Object -First 1 }
    $physDisk = Get-PhysicalDisk -ErrorAction Stop | Where-Object { $_.DeviceId -eq $osDisk.Number.ToString() } | Select-Object -First 1
    if (-not $physDisk) { $physDisk = Get-PhysicalDisk -ErrorAction Stop | Where-Object { $_.BusType -ne 'USB' } | Select-Object -First 1 }
    if ($physDisk) {
        $rawStorageGiB = $physDisk.Size / 1GB
        $storageEstimateGB = $rawStorageGiB * $GiBtoGB
        $storageGB = Get-NearestStandardSize -ActualGB $storageEstimateGB -Standards $storageStandards -Mode 'nearest'
        $storageType = if ($physDisk.BusType -eq 'NVMe') { "NVMe SSD" } elseif ($physDisk.MediaType -eq 'SSD') { "SSD" } elseif ($physDisk.MediaType -eq 'HDD') { "HDD" } else { "$($physDisk.MediaType)" }
        $storageInfo = "$(Format-StorageLabel $storageGB) $storageType"
    }
} catch {
    try {
        $mainDrive = Get-CimInstance Win32_DiskDrive -ErrorAction Stop | Where-Object { $_.MediaType -match 'Fixed' -or $_.InterfaceType -ne 'USB' } | Select-Object -First 1
        if ($mainDrive) {
            $rawStorageGiB = $mainDrive.Size / 1GB
            $storageEstimateGB = $rawStorageGiB * $GiBtoGB
            $storageGB = Get-NearestStandardSize -ActualGB $storageEstimateGB -Standards $storageStandards -Mode 'nearest'
            $storageInfo = "$(Format-StorageLabel $storageGB)"
        }
    } catch { }
}

$deviceSerial = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { "Unknown" }
$itemDescription = "Model: $brandModel`nDevice Serial: $deviceSerial`nCPU: $cpuFull`nOS: $osInfo`nRAM: ${ramGB}GB`nGPU: $gpuInfo`nScreen: $screenSize`nStorage: $storageInfo"

Write-Status "`nDetected specifications:" -Level Success
Write-Host "  Model    : $brandModel`n  Category : $itemCategory`n  Serial   : $deviceSerial`n  CPU      : $cpuFull`n  OS       : $osInfo`n  RAM      : ${ramGB}GB`n  GPU      : $gpuInfo`n  Screen   : $screenSize`n  Storage  : $storageInfo`n  Username : $defaultUser`n"

} else {
    # Not a computer/laptop - skip all hardware auto-detection, everything gets typed in manually
    # via the popup form below (an extra "Item Category" and "Item Description" box appears in
    # that case, since there's nothing to auto-fill them from).
    Write-Status "Not a computer/laptop - you'll enter all the item details manually." -Level Info
}

# Popup form instead of typing answers into the console - clearer about what's expected, and
# Location is a dropdown (fixed list) instead of free text so it can't be mistyped/inconsistent.

$locationOptions = @('MALL OF QATAR', 'BANI HAJER', 'BOYS SCHOOL', 'GIRLS SCHOOL')

$formHeight = if ($isThisAComputer) { 340 } else { 470 }
if ($EmailForm) { $formHeight += 45 }  # extra row for the "Send email to:" recipient dropdown
$formHeight += 190  # extra room for the signature box + its label/buttons
$form = New-Object System.Windows.Forms.Form
$form.Text = 'IT Asset Custody Form - Details'
$form.Size = New-Object System.Drawing.Size(420, $formHeight)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$form.TopMost = $true
$form.AutoScroll = $true

$y = 20
$labelLocation = New-Object System.Windows.Forms.Label
$labelLocation.Text = 'Location:'
$labelLocation.Location = New-Object System.Drawing.Point(20, $y)
$labelLocation.AutoSize = $true
$form.Controls.Add($labelLocation)

$comboLocation = New-Object System.Windows.Forms.ComboBox
$comboLocation.Location = New-Object System.Drawing.Point(160, ($y - 3))
$comboLocation.Size = New-Object System.Drawing.Size(220, 24)
$comboLocation.DropDownStyle = 'DropDownList'  # choose only from the list, can't type a new value
$comboLocation.Items.AddRange($locationOptions)
$comboLocation.SelectedIndex = 0
$form.Controls.Add($comboLocation)

$y += 45
$labelDept = New-Object System.Windows.Forms.Label
$labelDept.Text = 'Department:'
$labelDept.Location = New-Object System.Drawing.Point(20, $y)
$labelDept.AutoSize = $true
$form.Controls.Add($labelDept)

$textDept = New-Object System.Windows.Forms.TextBox
$textDept.Location = New-Object System.Drawing.Point(160, ($y - 3))
$textDept.Size = New-Object System.Drawing.Size(220, 24)
$form.Controls.Add($textDept)

$y += 45
$labelName = New-Object System.Windows.Forms.Label
$labelName.Text = "Custodian name`n(leave blank for `"$defaultUser`"):"
$labelName.Location = New-Object System.Drawing.Point(20, $y)
$labelName.AutoSize = $true
$form.Controls.Add($labelName)

$textName = New-Object System.Windows.Forms.TextBox
$textName.Location = New-Object System.Drawing.Point(160, ($y - 3))
$textName.Size = New-Object System.Drawing.Size(220, 24)
$form.Controls.Add($textName)

$y += 60
$labelAsset = New-Object System.Windows.Forms.Label
$labelAsset.Text = 'Asset Tag Number:'
$labelAsset.Location = New-Object System.Drawing.Point(20, $y)
$labelAsset.AutoSize = $true
$form.Controls.Add($labelAsset)

$textAsset = New-Object System.Windows.Forms.TextBox
$textAsset.Location = New-Object System.Drawing.Point(160, ($y - 3))
$textAsset.Size = New-Object System.Drawing.Size(220, 24)
$form.Controls.Add($textAsset)

# Signature box - lets whoever is receiving the item sign right here with a mouse or trackpad
# instead of printing, signing on paper, and re-scanning. Drawn onto an in-memory Bitmap that
# tracks the mouse while the left button is held; the finished Bitmap gets saved as a PNG and
# embedded into the "Signature" cell (C35) of the generated form further down. If left blank,
# the form is still generated fine - the signature cell is just left empty, same as before.
$y += 55
$labelSignature = New-Object System.Windows.Forms.Label
$labelSignature.Text = 'Signature (draw with mouse/trackpad below):'
$labelSignature.Location = New-Object System.Drawing.Point(20, $y)
$labelSignature.AutoSize = $true
$form.Controls.Add($labelSignature)

$y += 20
$sigWidth = 340
$sigHeight = 110
$panelSignature = New-Object System.Windows.Forms.Panel
$panelSignature.Location = New-Object System.Drawing.Point(20, $y)
$panelSignature.Size = New-Object System.Drawing.Size($sigWidth, $sigHeight)
$panelSignature.BackColor = [System.Drawing.Color]::White
$panelSignature.BorderStyle = 'FixedSingle'
$panelSignature.Cursor = [System.Windows.Forms.Cursors]::Cross
$form.Controls.Add($panelSignature)

# Backing bitmap the panel paints from - drawn to directly on mouse move, then blitted to the
# panel's own Paint event so it survives being covered/uncovered (e.g. minimizing the dialog).
$sigBitmap = New-Object System.Drawing.Bitmap($sigWidth, $sigHeight)
$sigGraphics = [System.Drawing.Graphics]::FromImage($sigBitmap)
$sigGraphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$sigGraphics.Clear([System.Drawing.Color]::White)
$sigPen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 2)
$sigPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
$sigPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
$script:sigDrawing = $false
$script:sigLastPoint = New-Object System.Drawing.Point(0, 0)
$script:sigHasStrokes = $false

$panelSignature.Add_MouseDown({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $script:sigDrawing = $true
        $script:sigLastPoint = $e.Location
    }
})
$panelSignature.Add_MouseMove({
    param($s, $e)
    if ($script:sigDrawing) {
        $sigGraphics.DrawLine($sigPen, $script:sigLastPoint, $e.Location)
        $script:sigLastPoint = $e.Location
        $script:sigHasStrokes = $true
        $panelSignature.Invalidate()
    }
})
$panelSignature.Add_MouseUp({
    param($s, $e)
    $script:sigDrawing = $false
})
$panelSignature.Add_Paint({
    param($s, $e)
    $e.Graphics.DrawImage($sigBitmap, 0, 0)
})
$form.Controls.Add($panelSignature)

$y += $sigHeight + 8
$clearSigButton = New-Object System.Windows.Forms.Button
$clearSigButton.Text = 'Clear signature'
$clearSigButton.Location = New-Object System.Drawing.Point(20, $y)
$clearSigButton.Size = New-Object System.Drawing.Size(120, 26)
$clearSigButton.Add_Click({
    $sigGraphics.Clear([System.Drawing.Color]::White)
    $script:sigHasStrokes = $false
    $panelSignature.Invalidate()
})
$form.Controls.Add($clearSigButton)

# Only shown when auto-email is on (-EmailForm) - lets whoever is running the tool choose whether
# THIS form goes to the default inbox or to the group IT inbox instead. Both destinations use the
# exact same OneDrive folder/upload/sign-in - the recipient choice is only encoded into the
# filename (an "[ITN]" tag), and a Condition step in the Power Automate flow branches on that tag
# to decide which address to send to. Nothing else about the upload changes.
$comboRecipient = $null
if ($EmailForm) {
    $y += 45
    $labelRecipient = New-Object System.Windows.Forms.Label
    $labelRecipient.Text = 'Send email to:'
    $labelRecipient.Location = New-Object System.Drawing.Point(20, $y)
    $labelRecipient.AutoSize = $true
    $form.Controls.Add($labelRecipient)

    $comboRecipient = New-Object System.Windows.Forms.ComboBox
    $comboRecipient.Location = New-Object System.Drawing.Point(160, ($y - 3))
    $comboRecipient.Size = New-Object System.Drawing.Size(220, 24)
    $comboRecipient.DropDownStyle = 'DropDownList'
    $comboRecipient.Items.AddRange(@('Default (jcarlos@sherborneqatar.org)', 'IT Group (itn@sherborneqatar.org)'))
    $comboRecipient.SelectedIndex = 0
    $form.Controls.Add($comboRecipient)
}

# Only shown for non-computer items (clicker, projector, screen, etc.) - nothing was auto-detected
# for these, so the category and description have to be typed in by hand.
$textCategory = $null
$textDescription = $null
if (-not $isThisAComputer) {
    $y += 45
    $labelCategory = New-Object System.Windows.Forms.Label
    $labelCategory.Text = "Item Category`n(e.g. Projector, Clicker):"
    $labelCategory.Location = New-Object System.Drawing.Point(20, $y)
    $labelCategory.AutoSize = $true
    $form.Controls.Add($labelCategory)

    $textCategory = New-Object System.Windows.Forms.TextBox
    $textCategory.Location = New-Object System.Drawing.Point(160, ($y - 3))
    $textCategory.Size = New-Object System.Drawing.Size(220, 24)
    $form.Controls.Add($textCategory)

    $y += 60
    $labelDescription = New-Object System.Windows.Forms.Label
    $labelDescription.Text = "Item Description`n(brand/model/serial, etc.):"
    $labelDescription.Location = New-Object System.Drawing.Point(20, $y)
    $labelDescription.AutoSize = $true
    $form.Controls.Add($labelDescription)

    $textDescription = New-Object System.Windows.Forms.TextBox
    $textDescription.Location = New-Object System.Drawing.Point(160, ($y - 3))
    $textDescription.Size = New-Object System.Drawing.Size(220, 60)
    $textDescription.Multiline = $true
    $form.Controls.Add($textDescription)
}

$y += 55
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = 'Generate Form'
$okButton.Location = New-Object System.Drawing.Point(160, $y)
$okButton.Size = New-Object System.Drawing.Size(120, 32)
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.Controls.Add($okButton)
$form.AcceptButton = $okButton

# Validate required fields before letting the dialog close - Department and Asset Tag always
# required; for non-computer items, Category and Description are required too since there's
# nothing auto-detected to fall back on.
$okButton.Add_Click({
    $missing = @()
    if (-not $textDept.Text.Trim())  { $missing += 'Department' }
    if (-not $textAsset.Text.Trim()) { $missing += 'Asset Tag Number' }
    if (-not $isThisAComputer) {
        if (-not $textCategory.Text.Trim())    { $missing += 'Item Category' }
        if (-not $textDescription.Text.Trim()) { $missing += 'Item Description' }
    }
    if ($missing.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please fill in: $($missing -join ', ')", 'Missing information', 'OK', 'Warning') | Out-Null
        $form.DialogResult = [System.Windows.Forms.DialogResult]::None
    }
})

$dialogResult = $form.ShowDialog()
if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Status "Cancelled - no form was generated." -Level Warning
    exit 0
}

$location   = $comboLocation.SelectedItem.ToString()
$company    = "Sherborne $location".Trim()
$department = $textDept.Text.Trim()
$staffName  = $textName.Text.Trim()
if ([string]::IsNullOrWhiteSpace($staffName)) { $staffName = $defaultUser }
$serialTag  = $textAsset.Text.Trim()

# Grab the drawn signature as PNG bytes (if anything was actually drawn) before the form/bitmap
# get disposed - left $null when the box was left blank, which just means the signature cell on
# the generated form stays empty (same as before this feature existed).
$signaturePngBytes = $null
if ($script:sigHasStrokes) {
    $sigStream = New-Object System.IO.MemoryStream
    $sigBitmap.Save($sigStream, [System.Drawing.Imaging.ImageFormat]::Png)
    $signaturePngBytes = $sigStream.ToArray()
    $sigStream.Close()
}

# Which inbox this form's emailed copy should go to - only asked when -EmailForm is on. Both
# options use the exact same OneDrive folder and upload process; the choice only changes a tag
# added to the UPLOADED copy's filename (see $uploadFileName below), which the Power Automate flow
# reads to decide which address to send to. The locally-saved file's name is unaffected.
$sendToGroup = $false
if ($EmailForm -and $comboRecipient) {
    $sendToGroup = ($comboRecipient.SelectedIndex -eq 1)
}

if (-not $isThisAComputer) {
    if ($textCategory.Text.Trim())    { $itemCategory    = $textCategory.Text.Trim() }
    if ($textDescription.Text.Trim()) { $itemDescription = $textDescription.Text.Trim() }
}

Write-Status "Generating form..." -Level Info

# Organize saved forms into a per-location subfolder (e.g. "Filled\MALL OF QATAR\") purely for
# tidiness - so forms from different sites don't all pile up loose in one folder. $location was
# already picked from the fixed dropdown above, so this can't produce a stray/mistyped folder name.
$locationFolder = Join-Path $OutputFolder $location
if (-not (Test-Path -LiteralPath $locationFolder)) {
    New-Item -ItemType Directory -Path $locationFolder -Force | Out-Null
}

$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
$safeName = ($staffName -replace "[$([regex]::Escape($invalidChars))]", "").Trim()
$fileDate = Get-Date -Format "dd-MM-yyyy"
$newFileName = "$safeName custody $fileDate.xlsx"
$newFormPath = Join-Path $locationFolder $newFileName

# Never silently overwrite an existing form - e.g. two custody forms for the same person on the
# same day (different assets, or a second staff member using a shared "Filled" location) would
# otherwise collide on the exact same filename and the first one would be lost with no warning.
# Add "(2)", "(3)", etc. until the name is free.
if (Test-Path -LiteralPath $newFormPath) {
    $copyNum = 2
    do {
        $newFileName = "$safeName custody $fileDate ($copyNum).xlsx"
        $newFormPath = Join-Path $locationFolder $newFileName
        $copyNum++
    } while (Test-Path -LiteralPath $newFormPath)
    Write-Status "A form for $staffName on $fileDate already existed - saving this one as `"$newFileName`" instead so nothing gets overwritten." -Level Warning
}

Copy-Item -Path $TemplatePath -Destination $newFormPath -Force

$ackText = "1. I,$staffName, holder of QID________________acknowledge that I have received the equipment listed on this document.   "

$values = @{
    "D3"  = $company.ToUpper()
    "D4"  = $department.ToUpper()
    "D5"  = $staffName.ToUpper()
    "B9"  = $itemCategory.ToUpper()
    "C9"  = $itemDescription.ToUpper()
    "E9"  = $serialTag.ToUpper()
    "F9"  = $dateIssued.ToUpper()
    "B19" = $ackText.ToUpper()
    "C34" = $staffName.ToUpper()
    "C36" = $dateIssued.ToUpper()
}

# Row 9 height is fixed (rather than auto-calculated from line count) to match the form's print layout.
$row9Height = 172.8

Fill-Workbook -Path $newFormPath -Values $values -SheetName "ITAssetTrackForm" -WrapCells @("C9") -ShrinkCells @("E9") -RowHeights @{ "9" = $row9Height }

# Drop the drawn signature (if any was drawn) into the "Signature" cell (C35) of the form -
# skipped entirely, leaving the cell blank as before, if the box was left empty.
if ($signaturePngBytes) {
    try {
        Add-SignatureImage -Path $newFormPath -PngBytes $signaturePngBytes -SheetName "ITAssetTrackForm" -CellRef "C35"
    } catch {
        Write-Status "Could not embed the signature image - the rest of the form is unaffected. ($($_.Exception.Message))" -Level Warning
    }
}

Write-Status "`nForm completed!" -Level Success
Write-Host "Saved: $newFormPath`n"
Write-Status "Ready to print." -Level Info

# Read the finished form's bytes into memory BEFORE opening it in Excel - once Excel has it open,
# Windows locks the file and a later read (e.g. for emailing) fails with "being used by another
# process". Reading it now means the upload step below never has to touch the file on disk again.
$formBytesForEmail = $null
try { $formBytesForEmail = [System.IO.File]::ReadAllBytes($newFormPath) } catch { }

# Auto-open the finished form so it works the same way whether launched via the .bat, the
# one-liner, or directly - previously only the one-liner's wrapper opened the file.
try {
    Write-Status "Opening file..." -Level Info
    Start-Process -FilePath $newFormPath -ErrorAction Stop
} catch {
    Write-Status "Could not auto-open the file - open it manually from: $newFormPath" -Level Warning
}

# Optional: email the finished form by uploading it to a OneDrive folder that a Power Automate
# flow watches ("When a file is created" -> "Send an email (V2)" to jcarlos@sherborneqatar.org).
# No mail credential or app secret lives in this script - it only ever gets a short-lived Graph
# sign-in token for whoever is running the tool, via the interactive "device code" flow (the same
# kind of "go to microsoft.com/devicelogin and enter this code" sign-in used by many CLI tools).
# No external module needed - just plain REST calls, consistent with this tool's zero-dependency design.
#
# TargetDriveId is jcarlos@sherborneqatar.org's OneDrive - uploads always go HERE regardless of
# which staff account signs in during the device-code step below (the CustodyFormsToEmail folder
# is shared org-wide with edit access, so any signed-in staff account is allowed to write to it).
# Without this, "/me/drive/..." would upload to whichever account just signed in's OWN OneDrive
# instead, which would never reach jcarlos's folder or trigger the email flow.
$targetDriveId = 'b!bjuGe2fvJ0KHNeTiVjoiPjhIOFedjfpMny1O22T8AlyqM2Vcc-37RZKqA8hym8r7'
if ($EmailForm) {
    # Quick connectivity check first - on a computer with no internet (common right after imaging,
    # before it's joined to Wi-Fi), skip straight to a clear message instead of sitting through the
    # device-code request's own connection timeout before failing anyway.
    $hasInternet = Test-Connection -ComputerName 'login.microsoftonline.com' -Count 1 -Quiet -ErrorAction SilentlyContinue
    if (-not $hasInternet) {
        Write-Status "No internet connection detected - skipping auto-email. The form is still saved at: $newFormPath" -Level Warning
    } else {
    try {
        Write-Status "`nSigning in to email the form (uses your Microsoft 365 account)..." -Level Info

        # Uses the well-known Microsoft Graph PowerShell/CLI public client ID (no app registration
        # needed on your end) with the device code flow: safe to use in a public script because it
        # never handles or stores a password - only a short-lived, single-purpose sign-in code.
        $clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
        $tenant   = 'common'
        # .All (not just Files.ReadWrite) is required here because the upload target is jcarlos's
        # OneDrive folder, not necessarily the signed-in person's own - .All covers writing to any
        # file/folder the signed-in account has been granted access to via sharing, not just its own.
        $scope    = 'Files.ReadWrite.All'

        $deviceResp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode" -Body @{ client_id = $clientId; scope = $scope } -ErrorAction Stop

        # Microsoft's device-code sign-in page does NOT auto-fill the code (this is a Microsoft
        # limitation, not something this script can bypass) - something has to type/paste it in.
        # The code is copied to the clipboard, the browser is opened to the sign-in page, and then
        # (best-effort) the code is auto-pasted into the page's own code box below, so normally
        # nothing needs to be typed at all before the password screen.
        #
        # IMPORTANT / hard boundary: this automation stops the instant the code is submitted. It
        # never touches, reads, sees, or interacts with the password field or any screen after the
        # code page - password entry always stays 100% manual, on purpose, with no exception.
        Write-Status "`n=========================================================" -Level Warning
        Write-Status "  Sign-in code: $($deviceResp.user_code)" -Level Warning
        Write-Status "  (auto-pasting this into the browser - just enter your password when asked)" -Level Warning
        Write-Status "=========================================================`n" -Level Warning
        try { Set-Clipboard -Value $deviceResp.user_code -ErrorAction Stop } catch { }

        $browserOpened = $false
        try {
            Start-Process $deviceResp.verification_uri -ErrorAction Stop
            $browserOpened = $true
        } catch {
            Write-Status "Could not auto-open the browser - go to $($deviceResp.verification_uri) manually and paste the code (already on your clipboard)." -Level Warning
        }

        if ($browserOpened) {
            # Best-effort auto-paste of the CODE ONLY. This brings the just-opened browser window to
            # the foreground and sends Ctrl+V then Enter - nothing more. It deliberately does not
            # attempt to detect or wait for the password field, does not send any further keystrokes,
            # and does nothing at all if the browser window can't be found - it just silently falls
            # back to manual paste, which always still works since the code is on the clipboard either way.
            try {
                Add-Type -AssemblyName System.Windows.Forms
                Add-Type -Name Win32Foreground -Namespace CustodyTool -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
[DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
'@ -ErrorAction SilentlyContinue

                # Give the browser a moment to launch and the sign-in page to load and focus the code box.
                Start-Sleep -Seconds 3

                # Find the browser process that just opened (best-effort - matches common browsers).
                # If none is found, or focusing fails, we just skip the auto-paste and the code stays
                # on the clipboard for a normal manual paste - never an error, never a retry loop.
                $browserProc = Get-Process -ErrorAction SilentlyContinue |
                    Where-Object { $_.MainWindowHandle -ne 0 -and $_.ProcessName -match '^(chrome|msedge|firefox|iexplore|brave|opera)$' } |
                    Sort-Object StartTime -Descending | Select-Object -First 1

                if ($browserProc) {
                    [CustodyTool.Win32Foreground]::SetForegroundWindow($browserProc.MainWindowHandle) | Out-Null
                    Start-Sleep -Milliseconds 500
                    [System.Windows.Forms.SendKeys]::SendWait('^v')
                    # Deliberately NOT auto-pressing Enter/Next here - the code box is the only thing
                    # touched. Clicking "Next" is left to you, since what follows (account picker,
                    # "Are you trying to sign in to..." consent, etc.) varies and shouldn't be driven
                    # blindly. This is also where the password step begins, which always stays manual.
                    Write-Status "Code auto-pasted into the browser - click Next/Verify there, then enter your password to finish signing in." -Level Info
                } else {
                    Write-Status "Couldn't auto-focus the browser window - just paste the code yourself (it's already on your clipboard)." -Level Info
                }
            } catch {
                Write-Status "Auto-paste didn't work - just paste the code yourself (it's already on your clipboard)." -Level Info
            }
        }

        $token = $null
        $deadline = (Get-Date).AddSeconds($deviceResp.expires_in)
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds $deviceResp.interval
            try {
                $tokenResp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/token" -Body @{
                    grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
                    client_id   = $clientId
                    device_code = $deviceResp.device_code
                } -ErrorAction Stop
                $token = $tokenResp.access_token
                break
            } catch {
                $err = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($err.error -eq 'authorization_pending') { continue }  # user hasn't finished signing in yet - keep polling
                throw
            }
        }

        if (-not $token) {
            Write-Status "Sign-in timed out - skipping auto-email. The form is still saved at: $newFormPath" -Level Warning
        } elseif (-not $formBytesForEmail) {
            Write-Status "Could not read the form's contents earlier - skipping auto-email. The form is still saved at: $newFormPath" -Level Warning
        } else {
            # Look up the email address of whoever just signed in, so they can be CC'd automatically
            # on the notification email - Power Automate's "Send an email" action always sends FROM
            # jcarlos's own connection (that's fixed per-flow, not something that changes per run),
            # but the CC line can list whoever actually uploaded the form. Uses the same token from
            # sign-in - no extra permission needed, reading your own profile is part of any Graph scope.
            $uploaderEmail = $null
            try {
                $meResp = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/me?`$select=mail,userPrincipalName" -Headers @{ Authorization = "Bearer $token" } -ErrorAction Stop
                $uploaderEmail = if ($meResp.mail) { $meResp.mail } else { $meResp.userPrincipalName }
            } catch { }

            Write-Status "Signed in - uploading form to OneDrive for emailing..." -Level Info
            # Routing info for the Power Automate flow ("[ITN]" when the group inbox was chosen,
            # "CC=someone_AT_domain.com" with the signed-in uploader's own email) travels in the
            # UPLOADED file's name itself, e.g. "[ITN] [CC someone_AT_domain.com] Name custody Date.xlsx".
            # ":" can't be used in a Windows filename, so "_AT_" replaces "@" instead, which is safe.
            # The flow's Subject/CC/Attachment-Name fields parse these tags back out of the filename
            # and strip them for display, so the recipient never sees the tag - just a clean name.
            $ccTag = if ($uploaderEmail) { "[CC $($uploaderEmail -replace '@', '_AT_')] " } else { '' }
            $groupTag = if ($sendToGroup) { '[ITN] ' } else { '' }
            $uploadFileName = "$groupTag$ccTag$newFileName"
            # Same per-location subfolder used for the local Desktop save (see $locationFolder above)
            # is mirrored here in OneDrive too, purely for organization - e.g.
            # ".../CustodyFormsToEmail/MALL OF QATAR/[CC ...] Name custody Date.xlsx". This is just
            # tidiness, not routing - the flow still reads CC/ITN from the FILENAME tags above, same
            # as always, so which subfolder the file sits in has no effect on where the email goes.
            # Encode the location and filename SEPARATELY (rather than the whole path at once) since
            # EscapeDataString would also encode the "/" separator between them - Windows filenames
            # can never contain a literal "/" themselves, so this can't be ambiguous.
            $uploadPath = "CustodyFormsToEmail/$([Uri]::EscapeDataString($location))/$([Uri]::EscapeDataString($uploadFileName))"
            # Uses the copy read into memory before the file was opened in Excel (see above) -
            # avoids "file in use" errors now that Excel has it open for you to view/print.
            $uploadUrl = "https://graph.microsoft.com/v1.0/drives/$targetDriveId/root:/$($uploadPath):/content"
            Invoke-RestMethod -Method Put -Uri $uploadUrl -Headers @{ Authorization = "Bearer $token" } -Body $formBytesForEmail -ContentType 'application/octet-stream' -ErrorAction Stop | Out-Null
            $recipientForMessage = if ($sendToGroup) { 'itn@sherborneqatar.org' } else { 'jcarlos@sherborneqatar.org' }
            $ccNote = if ($uploaderEmail) { " (CC: $uploaderEmail)" } else { '' }
            Write-Status "Uploaded - the email will arrive at $recipientForMessage$ccNote within a few minutes (the free OneDrive trigger checks periodically, not instantly)." -Level Success

            # Clean up the OneDrive copy's filename after the flow has almost certainly already
            # picked it up, so the file sitting in OneDrive long-term reads as a normal clean name
            # instead of keeping the "[CC ...]"/"[ITN]" tag forever.
            #
            # v3.9: bumped from 90s to 8 minutes AND moved to a separate detached process, after
            # confirming live (via Power Automate run history) that the free/shared-tier OneDrive
            # trigger was regularly taking 1-5+ minutes to fire - well past the old 90-second
            # delay - which meant the rename was consistently winning the race and silently wiping
            # the "[CC ...]" tag before the trigger ever read it (CC came back blank on every
            # recent run). 8 minutes is comfortably past every delay observed so far, but this is
            # still a fixed guess, not a guarantee - an unusually slow poll cycle could theoretically
            # still lose the race. Running this in a separate process (Start-Process, not Start-Job)
            # means it survives even if this console window is closed right after upload - the
            # rename happens invisibly a few minutes later instead of blocking you from closing the
            # tool. NOTE: the access token is passed as a command-line argument to that process, so
            # it's briefly visible to anything with admin/Task Manager access on this computer for
            # the few minutes the background process runs - acceptable since the token is short-lived
            # (~1 hour) and scoped to the signed-in user's own account.
            $renameDelaySeconds = 480
            try {
                Write-Status "The OneDrive filename will be tidied up in the background in a few minutes (no need to wait - you can close this window now)." -Level Info
                $renameUrl = "https://graph.microsoft.com/v1.0/drives/$targetDriveId/root:/$uploadPath"
                $renameHelper = @"
param([string]`$DelaySeconds, [string]`$RenameUrl, [string]`$CleanName, [string]`$Token, [string]`$LogPath)
Start-Sleep -Seconds ([int]`$DelaySeconds)
try {
    `$body = @{ name = `$CleanName } | ConvertTo-Json
    Invoke-RestMethod -Method Patch -Uri `$RenameUrl -Headers @{ Authorization = "Bearer `$Token"; 'Content-Type' = 'application/json' } -Body `$body -ErrorAction Stop | Out-Null
    Add-Content -LiteralPath `$LogPath -Value "`$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  [Info] Renamed OneDrive copy back to clean filename (background): `$CleanName"
} catch {
    Add-Content -LiteralPath `$LogPath -Value "`$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  [Warning] Could not tidy up the OneDrive filename afterward (background): `$(`$_.Exception.Message)"
}
"@
                $helperPath = Join-Path $env:TEMP "CustodyRename_$(Get-Random).ps1"
                Set-Content -LiteralPath $helperPath -Value $renameHelper -Encoding UTF8
                Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList @(
                    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$helperPath`"",
                    "`"$renameDelaySeconds`"", "`"$renameUrl`"", "`"$newFileName`"", "`"$token`"", "`"$script:LogPath`""
                ) | Out-Null
                Write-Log "[Info] Launched background rename (will fire in $renameDelaySeconds seconds)."
            } catch {
                # Non-fatal - the email has already gone out either way by this point. Worst case the
                # OneDrive copy keeps its tagged name, which is harmless, just less tidy.
                Write-Log "[Warning] Could not launch background rename: $($_.Exception.Message)"
            }
        }
    } catch {
        # Capture the real underlying error detail (Graph/Microsoft errors often put the useful
        # message in ErrorDetails rather than the generic Exception.Message), and log it in full
        # so it's there to read/copy even after the window closes.
        $detail = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $detail = "$detail | $($_.ErrorDetails.Message)" }
        Write-Status "Could not email the form automatically - it's still saved at: $newFormPath" -Level Warning
        Write-Log "[Error] Email upload failed: $detail"
        Open-LogOnError
    }
    }
}

Start-Sleep -Seconds 2
