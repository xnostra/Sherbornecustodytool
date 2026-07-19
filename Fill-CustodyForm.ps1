<#
Fill-CustodyForm.ps1
Auto-fills "custody form.xlsx" (IT Asset Custody/Tracking Form) using detailed
spec info pulled automatically from the local machine (brand, model, CPU,
gen, OS, RAM, GPU, screen size, storage size/type), plus a few prompts for
details Windows can't know (Company, Department, staff name, and your own
asset tag number since you use a different tagging system).

The description cell wraps onto multiple lines and the row is made taller
so it prints in full. The asset tag cell auto-shrinks its font if you type
a long tag so it still fits on one line.

ZERO-DEPENDENCY / FULLY PORTABLE VERSION
-----------------------------------------
No external module needed (no ImportExcel, no Excel install, no
Install-Module, no internet, no admin rights). Edits the .xlsx directly via
.NET's built-in System.IO.Compression classes, which ship with every
Windows PowerShell 5.1+ install. Copy this whole folder to a USB and run.

USB folder layout:
    CustodyTool\Fill-CustodyForm.ps1
    CustodyTool\custody form.xlsx   <- template (must stay next to the script)
    CustodyTool\Filled\             <- output goes here (auto-created)
#>


param(
    [string]$TemplatePath = (Join-Path $PSScriptRoot "custody form.xlsx"),
    [string]$OutputFolder = (Join-Path $PSScriptRoot "Filled")
)

# --- Always allow local scripts to run whenever this is launched ---
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

trap {
    $msg = $_.Exception.Message
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if ($msg -match 'Access.*is denied|UnauthorizedAccessException|requires elevation' -and -not $isAdmin) {
        Write-Host "This step needs admin rights - requesting elevation..." -ForegroundColor Yellow
        Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    }
    Write-Host ""
    Write-Host "SOMETHING WENT WRONG:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "The form was NOT filled in because of the error above." -ForegroundColor Yellow
    Read-Host "Press Enter to close this window"
    exit 1
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

if (-not (Test-Path $TemplatePath)) {
    Write-Error "Template not found at: $TemplatePath (expected next to this script)"
    exit 1
}
if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

function ColLetterToNum([string]$col) {
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
    if ($OriginalIndex -lt $xfNodes.Count) {
        $newXf = $xfNodes.Item($OriginalIndex).CloneNode($true)
    } else {
        $newXf = $StylesXml.CreateElement('xf', $ns)
        $newXf.SetAttribute('numFmtId', '0')
        $newXf.SetAttribute('fontId', '0')
        $newXf.SetAttribute('fillId', '0')
        $newXf.SetAttribute('borderId', '0')
        $newXf.SetAttribute('xfId', '0')
    }
    $newXf.SetAttribute('applyAlignment', '1')
    $align = $newXf.SelectSingleNode('s:alignment', $NsMgr)
    if (-not $align) {
        $align = $StylesXml.CreateElement('alignment', $ns)
        $newXf.AppendChild($align) | Out-Null
    }
    $align.SetAttribute('horizontal', 'left')
    $align.SetAttribute('vertical', 'center')
    if ($Mode -eq 'wrap') {
        $align.SetAttribute('wrapText', '1')
    } elseif ($Mode -eq 'shrink') {
        $align.SetAttribute('shrinkToFit', '1')
    }
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
    $xmlDoc.LoadXml($reader.ReadToEnd())
    $reader.Close()
    return $xmlDoc
}

function Write-ZipXml {
    param($Zip, [string]$EntryName, $XmlDoc)
    $Zip.GetEntry($EntryName).Delete()
    $newEntry = $Zip.CreateEntry($EntryName)
    $stream = $newEntry.Open()
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $writerSettings = New-Object System.Xml.XmlWriterSettings
    $writerSettings.Encoding = $utf8NoBom
    $writer = [System.Xml.XmlWriter]::Create($stream, $writerSettings)
    $XmlDoc.Save($writer)
    $writer.Close()
    $stream.Close()
}

function Set-PrintLayout {
    # Makes printing "just work": fits the sheet to one page wide, centers it,
    # and sets even, print-safe margins so nothing gets clipped on the left/right.
    param($SheetXml, $NsMgr, $RootNode)

    $order = @('sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr',
               'sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate',
               'customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations',
               'hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks',
               'customProperties','cellWatches','ignoredErrors','smartTags','drawing','drawingHF','picture',
               'oleObjects','controls','webPublishItems','tableParts','extLst')

    function Get-OrAddElement([string]$TagName) {
        $node = $RootNode.SelectSingleNode("s:$TagName", $NsMgr)
        if (-not $node) {
            $node = $SheetXml.CreateElement($TagName, $ns)
            $tagIndex = [array]::IndexOf($order, $TagName)
            $target = $null
            foreach ($child in $RootNode.ChildNodes) {
                if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                $childIndex = [array]::IndexOf($order, $child.LocalName)
                if ($childIndex -ge 0 -and $childIndex -gt $tagIndex) { $target = $child; break }
            }
            if ($target) { $RootNode.InsertBefore($node, $target) | Out-Null }
            else { $RootNode.AppendChild($node) | Out-Null }
        }
        return $node
    }

    $sheetPr = Get-OrAddElement 'sheetPr'
    $pageSetUpPr = $sheetPr.SelectSingleNode('s:pageSetUpPr', $NsMgr)
    if (-not $pageSetUpPr) {
        $pageSetUpPr = $SheetXml.CreateElement('pageSetUpPr', $ns)
        $sheetPr.AppendChild($pageSetUpPr) | Out-Null
    }
    $pageSetUpPr.SetAttribute('fitToPage', '1')

    $printOptions = Get-OrAddElement 'printOptions'
    $printOptions.SetAttribute('horizontalCentered', '1')

    $pageMargins = Get-OrAddElement 'pageMargins'
    $pageMargins.SetAttribute('left', '0.5')
    $pageMargins.SetAttribute('right', '0.5')
    $pageMargins.SetAttribute('top', '0.75')
    $pageMargins.SetAttribute('bottom', '0.75')
    $pageMargins.SetAttribute('header', '0.3')
    $pageMargins.SetAttribute('footer', '0.3')

    $pageSetup = Get-OrAddElement 'pageSetup'
    $pageSetup.SetAttribute('fitToWidth', '1')
    $pageSetup.SetAttribute('fitToHeight', '1')
    $pageSetup.SetAttribute('orientation', 'portrait')
}

function Fill-Workbook {
    param(
        [string]$Path,
        [hashtable]$Values,
        [string]$SheetName,
        [string[]]$WrapCells = @(),
        [string[]]$ShrinkCells = @(),
        [hashtable]$RowHeights = @{}
    )

    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $wbXml = Read-ZipXml -Zip $zip -EntryName "xl/workbook.xml"
        $nsMgr = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
        $nsMgr.AddNamespace('s', $ns)
        $nsMgr.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $sheetNode = $wbXml.SelectSingleNode("//s:sheets/s:sheet[@name='$SheetName']", $nsMgr)
        if (-not $sheetNode) { throw "Sheet '$SheetName' not found in workbook.xml" }
        $rId = $sheetNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

        $relsXml = Read-ZipXml -Zip $zip -EntryName "xl/_rels/workbook.xml.rels"
        $relNode = $relsXml.SelectSingleNode("//*[local-name()='Relationship' and @Id='$rId']")
        $target = $relNode.GetAttribute('Target') -replace '^/', ''
        if ($target -notmatch '^xl/') { $target = "xl/$target" }

        $sheetXml = Read-ZipXml -Zip $zip -EntryName $target
        $sheetNsMgr = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
        $sheetNsMgr.AddNamespace('s', $ns)
        Set-PrintLayout -SheetXml $sheetXml -NsMgr $sheetNsMgr -RootNode $sheetXml.DocumentElement

        $stylesXml = Read-ZipXml -Zip $zip -EntryName "xl/styles.xml"
        $stylesNsMgr = New-Object System.Xml.XmlNamespaceManager($stylesXml.NameTable)
        $stylesNsMgr.AddNamespace('s', $ns)
        $cellXfsNode = $stylesXml.SelectSingleNode("//s:styleSheet/s:cellXfs", $stylesNsMgr)
        $styleCache = @{}

        foreach ($cellRef in $Values.Keys) {
            $rowNum = ($cellRef -replace '\D', '')
            $row = $sheetXml.SelectSingleNode("//s:sheetData/s:row[@r='$rowNum']", $sheetNsMgr)
            if (-not $row) { Write-Warning "Row $rowNum not found for cell $cellRef, skipping"; continue }
            Set-CellInlineValue -RowNode $row -XmlDoc $sheetXml -NsMgr $sheetNsMgr -CellRef $cellRef -Value $Values[$cellRef]
        }

        foreach ($cellRef in $Values.Keys) {
            $rowNum = ($cellRef -replace '\D', '')
            $row = $sheetXml.SelectSingleNode("//s:sheetData/s:row[@r='$rowNum']", $sheetNsMgr)
            if (-not $row) { continue }
            $cell = $row.SelectSingleNode("s:c[@r='$cellRef']", $sheetNsMgr)
            if (-not $cell) { continue }
            $origIdx = 0
            if ($cell.GetAttribute('s')) { $origIdx = [int]$cell.GetAttribute('s') }
            $mode = if ($WrapCells -contains $cellRef) { 'wrap' } elseif ($ShrinkCells -contains $cellRef) { 'shrink' } else { 'align' }
            $newIdx = Get-StyleVariantIndex -CellXfsNode $cellXfsNode -StylesXml $stylesXml -NsMgr $stylesNsMgr -OriginalIndex $origIdx -Mode $mode -Cache $styleCache
            $cell.SetAttribute('s', $newIdx)
        }

        foreach ($rowNum in $RowHeights.Keys) {
            $row = $sheetXml.SelectSingleNode("//s:sheetData/s:row[@r='$rowNum']", $sheetNsMgr)
            if ($row) {
                $row.SetAttribute('ht', [string]$RowHeights[$rowNum])
                $row.SetAttribute('customHeight', '1')
            }
        }

        Write-ZipXml -Zip $zip -EntryName $target -XmlDoc $sheetXml
        Write-ZipXml -Zip $zip -EntryName "xl/styles.xml" -XmlDoc $stylesXml
    } finally {
        $zip.Dispose()
    }
}

# ================= Auto-collected machine info =================
$cs   = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$enc  = Get-CimInstance Win32_SystemEnclosure
$cpu  = Get-CimInstance Win32_Processor | Select-Object -First 1
$os   = Get-CimInstance Win32_OperatingSystem

$laptopTypes = 8,9,10,11,12,14,18,21
$isLaptop = ($enc.ChassisTypes | Where-Object { $laptopTypes -contains $_ }).Count -gt 0
$itemCategory = if ($isLaptop) { "Laptop/Notebook" } else { "Desktop Computer" }

# Rounds a raw reported size to the nearest real-world marketing size.
# Mode 'ceiling': for RAM - Windows always reports a bit LESS than the true
#   installed size (memory reserved for firmware/graphics), so round UP to
#   the next size that's actually sold, never down.
# Mode 'nearest': for storage - drives are marketed in decimal GB but
#   Windows reports binary GiB, so after converting units the estimate is
#   close on both sides of the true value; pick whichever standard size is closest.
function Get-NearestStandardSize {
    param([double]$ActualGB, [double[]]$Standards, [string]$Mode = 'ceiling')
    if ($Mode -eq 'nearest') {
        $best = $Standards[0]
        $bestDiff = [Math]::Abs($Standards[0] - $ActualGB)
        foreach ($s in $Standards) {
            $diff = [Math]::Abs($s - $ActualGB)
            if ($diff -lt $bestDiff) { $bestDiff = $diff; $best = $s }
        }
        return $best
    } else {
        foreach ($s in $Standards) {
            if ($ActualGB -le ($s * 1.04)) { return $s }
        }
        return $Standards[-1]
    }
}
$ramStandards     = 1,2,3,4,6,8,12,16,24,32,48,64,96,128,192,256,384,512
$storageStandards = 16,32,64,120,128,240,250,256,320,480,500,512,640,750,960,1000,1024,1500,1920,2000,2048,3000,3840,4000,4096,6000,8000,10000,16000
$GiBtoGB = 1.073741824  # decimal-GB manufacturers use vs binary-GiB Windows reports

# Model name + SKU for auditors (SKU gives an exact, unambiguous model
# reference even when "Model" itself is a generic marketing name)
$modelName = $cs.Model
$sysSku = $cs.SystemSKUNumber
$brandModel = "$($cs.Manufacturer) $modelName".Trim()
if ($sysSku -and $sysSku.Trim() -ne '' -and $sysSku.Trim() -ne $modelName.Trim()) {
    $brandModel += " (SKU: $($sysSku.Trim()))"
}

$cpuName = $cpu.Name -replace '\s+', ' '
$cpuGen = $null
if ($cpuName -match 'i[3579]-(\d{4,5})') {
    $num = $Matches[1]
    $cpuGen = if ($num.Length -eq 5) { $num.Substring(0,2) } else { $num.Substring(0,1) }
}
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

function Format-StorageLabel([double]$SizeGB) {
    if ($SizeGB -ge 1000 -and ($SizeGB % 1000) -eq 0) { return "$([int]($SizeGB/1000))TB" }
    elseif ($SizeGB -ge 1000) { return "$([Math]::Round($SizeGB/1000,1))TB" }
    else { return "${SizeGB}GB" }
}

# Only the computer's own main/system drive - external or USB-attached
# drives are deliberately excluded so they don't get logged as the asset's storage.
$storageInfo = "Unknown"
try {
    $osDisk = Get-Disk -ErrorAction Stop | Where-Object { $_.IsBoot -or $_.IsSystem } | Select-Object -First 1
    if (-not $osDisk) { $osDisk = Get-Disk -ErrorAction Stop | Sort-Object Number | Select-Object -First 1 }
    $physDisk = Get-PhysicalDisk -ErrorAction Stop | Where-Object { $_.DeviceId -eq $osDisk.Number.ToString() } | Select-Object -First 1
    if (-not $physDisk) { $physDisk = Get-PhysicalDisk -ErrorAction Stop | Where-Object { $_.BusType -ne 'USB' } | Select-Object -First 1 }

    $rawStorageGiB = $physDisk.Size / 1GB
    $storageEstimateGB = $rawStorageGiB * $GiBtoGB
    $storageGB = Get-NearestStandardSize -ActualGB $storageEstimateGB -Standards $storageStandards -Mode 'nearest'
    $storageType = if ($physDisk.BusType -eq 'NVMe') { "NVMe SSD" }
                   elseif ($physDisk.MediaType -eq 'SSD') { "SSD" }
                   elseif ($physDisk.MediaType -eq 'HDD') { "HDD" }
                   else { "$($physDisk.MediaType)" }
    $storageInfo = "$(Format-StorageLabel $storageGB) $storageType"
} catch {
    try {
        $mainDrive = Get-CimInstance Win32_DiskDrive -ErrorAction Stop | Where-Object { $_.MediaType -match 'Fixed' -or $_.InterfaceType -ne 'USB' } | Select-Object -First 1
        if ($mainDrive) {
            $rawStorageGiB = $mainDrive.Size / 1GB
            $storageEstimateGB = $rawStorageGiB * $GiBtoGB
            $storageGB = Get-NearestStandardSize -ActualGB $storageEstimateGB -Standards $storageStandards -Mode 'nearest'
            $storageInfo = "$(Format-StorageLabel $storageGB) (type unknown)"
        }
    } catch { }
}

$defaultUser = $env:USERNAME
$dateIssued  = Get-Date -Format "dd-MMM-yyyy"

# One spec per line so it wraps cleanly and prints in full
$deviceSerial = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { "Unknown" }
$itemDescription = "$brandModel`nDevice Serial: $deviceSerial`nCPU: $cpuFull`nOS: $osInfo`nRAM: ${ramGB}GB`nGPU: $gpuInfo`nScreen: $screenSize`nStorage: $storageInfo"

Write-Host "`nDetected from this machine:" -ForegroundColor Cyan
Write-Host "  Category : $itemCategory"
Write-Host "  Brand    : $brandModel"
Write-Host "  Serial   : $deviceSerial"
Write-Host "  CPU      : $cpuFull"
Write-Host "  OS       : $osInfo"
Write-Host "  RAM      : ${ramGB}GB"
Write-Host "  GPU      : $gpuInfo"
Write-Host "  Screen   : $screenSize"
Write-Host "  Storage  : $storageInfo"
Write-Host "  Username : $defaultUser`n"

# ================= Info that can't be auto-detected =================
$location   = Read-Host "Location (e.g. Qatar, BH)"
$company    = "Sherborne $location".Trim()
$department = Read-Host "Department"
$staffName  = Read-Host "Staff/Custodian full name [$defaultUser]"
if ([string]::IsNullOrWhiteSpace($staffName)) { $staffName = $defaultUser }
$serialTag  = Read-Host "Serial/Tag No (your own asset tag - type it in)"

# ================= Copy the template and fill it in =================
$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
$safeName = ($staffName -replace "[$([regex]::Escape($invalidChars))]", "").Trim()
$fileDate = Get-Date -Format "dd-MM-yyyy"
$newFileName = "$safeName custody $fileDate.xlsx"
$newFormPath = Join-Path $OutputFolder $newFileName

Copy-Item -Path $TemplatePath -Destination $newFormPath -Force

# QID left blank as-is (it's optional on the physical form - fill by hand if needed)
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

# Auto-fit row 9's height to however many lines the description actually has,
# instead of a fixed guess - so it's never cramped and never wastes space.
$descLineCount = ($itemDescription -split "`n").Count
$autoRowHeight = ($descLineCount + 1) * 14

Fill-Workbook -Path $newFormPath -Values $values -SheetName "ITAssetTrackForm" `
    -WrapCells @("C9") -ShrinkCells @("E9") -RowHeights @{ "9" = $autoRowHeight }

Write-Host "`nDone. New form saved: $newFormPath" -ForegroundColor Green
Write-Host "Row height and print layout (margins, fit-to-page, centering) are set automatically - just print." -ForegroundColor DarkGray
Write-Host "`nClosing in 3 seconds..." -ForegroundColor DarkGray
Start-Sleep -Seconds 3
