<#
.SYNOPSIS
    One-liner launcher for IT Asset Custody Form Tool
    Downloads template and script, then executes the custody form tool.

.DESCRIPTION
    This script can be run with:
    irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex

.NOTES
    Version: 2.0
    Author: Sherborne Custody Tool Team
    LastModified: 2026-07-20

.LINK
    https://github.com/xnostra/Sherbornecustodytool
#>

$repoUrl = "https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main"
$custodyScriptUrl = "$repoUrl/Fill-CustodyForm.ps1"
$templateUrl = "$repoUrl/custody%20form.xlsx"

# Create working directory
$workDir = Join-Path $env:TEMP "CustodyTool_$(Get-Random)"
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

$scriptPath = Join-Path $workDir "Fill-CustodyForm.ps1"
$templatePath = Join-Path $workDir "custody form.xlsx"

Write-Host "IT Asset Custody Form Tool" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

try {
    # Download the main script
    Write-Host "Downloading custody form script..." -ForegroundColor Yellow
    $custodyScript = Invoke-RestMethod -Uri $custodyScriptUrl -ErrorAction Stop
    $custodyScript | Out-File -FilePath $scriptPath -Encoding UTF8 -Force
    Write-Host "✓ Script downloaded" -ForegroundColor Green

    # Download the template file using WebClient (handles binary files better)
    Write-Host "Downloading Excel template..." -ForegroundColor Yellow
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($templateUrl, $templatePath)
    Write-Host "✓ Template downloaded" -ForegroundColor Green

    Write-Host ""
    Write-Host "Executing custody form tool..." -ForegroundColor Cyan
    Write-Host ""

    # Execute the script from the working directory with explicit paths
    Push-Location $workDir
    & $scriptPath -TemplatePath $templatePath -OutputFolder (Join-Path $workDir "Filled")
    Pop-Location

    # Auto-open the generated file
    Start-Sleep -Seconds 1
    $filledFolder = Join-Path $workDir "Filled"
    if (Test-Path $filledFolder) {
        $latestFile = Get-ChildItem -Path $filledFolder -Filter "*.xlsx" -ErrorAction SilentlyContinue | Sort-Object CreationTime -Descending | Select-Object -First 1
        if ($latestFile) {
            Write-Host ""
            Write-Host "Opening generated form..." -ForegroundColor Green
            Start-Process $latestFile.FullName
        }
    }

    Write-Host ""
    Write-Host "Cleanup in 60 seconds..." -ForegroundColor DarkGray
    Start-Sleep -Seconds 60

    # Cleanup
    Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Host ""
    Write-Host "ERROR: Failed to execute custody tool" -ForegroundColor Red
    Write-Host "URL (Script): $custodyScriptUrl" -ForegroundColor Red
    Write-Host "URL (Template): $templateUrl" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to close"

    # Cleanup on error
    Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}
