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
    Version: 2.6
    Requires: Windows 10/11, PowerShell 5.1+
    LastModified: 2026-07-20
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
        $elevInner   = "& ([scriptblock]::Create((Get-Content -Raw -LiteralPath '$scriptSelf'))) -TemplatePath '$TemplatePath' -OutputFolder '$OutputFolder'"
        $elevEncoded = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($elevInner))
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $elevEncoded"
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
$brandModel = "$($cs.Manufacturer) $modelName".Trim()
if ($sysSku -and $sysSku.Trim() -ne '' -and $sysSku.Trim() -ne $modelName.Trim()) { $brandModel += " (SKU: $($sysSku.Trim()))" }

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
$defaultUser = $env:USERNAME
$dateIssued  = Get-Date -Format "dd-MMM-yyyy"
$itemDescription = "$brandModel`nDevice Serial: $deviceSerial`nCPU: $cpuFull`nOS: $osInfo`nRAM: ${ramGB}GB`nGPU: $gpuInfo`nScreen: $screenSize`nStorage: $storageInfo"

Write-Status "`nDetected specifications:" -Level Success
Write-Host "  Category : $itemCategory`n  Brand    : $brandModel`n  Serial   : $deviceSerial`n  CPU      : $cpuFull`n  OS       : $osInfo`n  RAM      : ${ramGB}GB`n  GPU      : $gpuInfo`n  Screen   : $screenSize`n  Storage  : $storageInfo`n  Username : $defaultUser`n"

$location   = Read-Host "Location (e.g., Qatar, BH, UK)"
$company    = "Sherborne $location".Trim()
$department = Read-Host "Department"
$staffName  = Read-Host "Staff/Custodian name [$defaultUser]"
if ([string]::IsNullOrWhiteSpace($staffName)) { $staffName = $defaultUser }
$serialTag  = Read-Host "Asset Tag Number"

Write-Status "Generating form..." -Level Info

$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
$safeName = ($staffName -replace "[$([regex]::Escape($invalidChars))]", "").Trim()
$fileDate = Get-Date -Format "dd-MM-yyyy"
$newFileName = "$safeName custody $fileDate.xlsx"
$newFormPath = Join-Path $OutputFolder $newFileName

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
if ($EmailForm) {
    try {
        Write-Status "`nSigning in to email the form (uses your Microsoft 365 account)..." -Level Info

        # Uses the well-known Microsoft Graph PowerShell/CLI public client ID (no app registration
        # needed on your end) with the device code flow: safe to use in a public script because it
        # never handles or stores a password - only a short-lived, single-purpose sign-in code.
        $clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
        $tenant   = 'common'
        $scope    = 'Files.ReadWrite'

        $deviceResp = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenant/oauth2/v2.0/devicecode" -Body @{ client_id = $clientId; scope = $scope } -ErrorAction Stop

        # Microsoft's device-code sign-in page does NOT auto-fill the code (this is a Microsoft
        # limitation, not something this script can bypass) - you always have to type or paste it
        # in yourself. To make that as painless as possible: the code is copied to your clipboard
        # automatically, so you can just paste it into the box on the page that opens.
        Write-Status "`n=========================================================" -Level Warning
        Write-Status "  Sign-in code: $($deviceResp.user_code)" -Level Warning
        Write-Status "  (already copied to your clipboard - just paste it in)" -Level Warning
        Write-Status "=========================================================`n" -Level Warning
        try { Set-Clipboard -Value $deviceResp.user_code -ErrorAction Stop } catch { }

        try {
            Start-Process $deviceResp.verification_uri -ErrorAction Stop
        } catch {
            Write-Status "Could not auto-open the browser - go to $($deviceResp.verification_uri) manually." -Level Warning
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
            Write-Status "Signed in - uploading form to OneDrive for emailing..." -Level Info
            # Uses the copy read into memory before the file was opened in Excel (see above) -
            # avoids "file in use" errors now that Excel has it open for you to view/print.
            $uploadUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/CustodyFormsToEmail/$([Uri]::EscapeDataString($newFileName)):/content"
            Invoke-RestMethod -Method Put -Uri $uploadUrl -Headers @{ Authorization = "Bearer $token" } -Body $formBytesForEmail -ContentType 'application/octet-stream' -ErrorAction Stop | Out-Null
            Write-Status "Uploaded - the email will arrive at jcarlos@sherborneqatar.org within a few minutes (the free OneDrive trigger checks periodically, not instantly)." -Level Success
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

Start-Sleep -Seconds 2
