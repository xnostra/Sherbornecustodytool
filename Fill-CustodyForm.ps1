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
    Version: 2.0
    Requires: Windows 10/11, PowerShell 5.1+
    LastModified: 2026-07-20
    Dependencies: None (zero external dependencies)

.LINK
    https://github.com/xnostra/Sherbornecustodytool

#>

[CmdletBinding()]
param(
    [string]$TemplatePath = "",
    [string]$OutputFolder = ""
)

$ErrorActionPreference = "Stop"

# Handle PSScriptRoot being empty when run via Invoke-Expression
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

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

function Write-Status {
    param([string]$Message, [ValidateSet('Info', 'Success', 'Warning', 'Error')] [string]$Level = 'Info')
    $colors = @{ 'Info' = 'Cyan'; 'Success' = 'Green'; 'Warning' = 'Yellow'; 'Error' = 'Red' }
    Write-Host $Message -ForegroundColor $colors[$Level]
}

trap {
    $msg = $_.Exception.Message
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($msg -match 'Access.*denied|UnauthorizedAccessException|requires elevation' -and -not $isAdmin) {
        Write-Status "Admin rights needed - requesting elevation..." -Level Warning
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -TemplatePath `"$TemplatePath`" -OutputFolder `"$OutputFolder`""
        exit 0
    }
    Write-Status "`nERROR:" -Level Error
    Write-Status $_.Exception.Message -Level Error
    Write-Status $_.InvocationInfo.PositionMessage -Level Warning
    Read-Host "Press Enter to close"
    exit 1
}

Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
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
    $pageSetup.SetAttribute('fitToWidth', '1'); $pageSetup.SetAttribute('fitToHeight', '1'); $pageSetup.SetAttribute('orientation', 'portrait')
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

$descLineCount = ($itemDescription -split "`n").Count
$autoRowHeight = ($descLineCount + 1) * 14

Fill-Workbook -Path $newFormPath -Values $values -SheetName "ITAssetTrackForm" -WrapCells @("C9") -ShrinkCells @("E9") -RowHeights @{ "9" = $autoRowHeight }

Write-Status "`nForm completed!" -Level Success
Write-Host "Saved: $newFormPath`n"
Write-Status "Ready to print." -Level Info
Start-Sleep -Seconds 2
