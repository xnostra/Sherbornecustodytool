<#
.SYNOPSIS
    One-liner launcher for IT Asset Custody Form Tool
    Downloads and executes the custody form script from GitHub with a single command.

.DESCRIPTION
    This script can be run with:
    irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex

.NOTES
    Version: 1.0
    Author: Sherborne Custody Tool Team
    LastModified: 2026-07-20

.LINK
    https://github.com/xnostra/Sherbornecustodytool
#>

$gitHubRawUrl = "https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/Fill-CustodyForm.ps1"
$outputFolder = Join-Path $PSScriptRoot "Filled"

Write-Host "IT Asset Custody Form Tool" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Downloading script..." -ForegroundColor Yellow

try {
    $custodyScript = Invoke-RestMethod -Uri $gitHubRawUrl -ErrorAction Stop
    Write-Host "Script downloaded successfully." -ForegroundColor Green
    Write-Host "Executing custody form tool..." -ForegroundColor Cyan
    Write-Host ""

    # Execute the downloaded script
    Invoke-Expression $custodyScript

    # Auto-open the most recently created file in the Filled folder
    if (Test-Path $outputFolder) {
        $latestFile = Get-ChildItem -Path $outputFolder -Filter "*.xlsx" -ErrorAction SilentlyContinue | Sort-Object CreationTime -Descending | Select-Object -First 1
        if ($latestFile) {
            Write-Host ""
            Write-Host "Opening generated form..." -ForegroundColor Cyan
            Start-Sleep -Milliseconds 500
            & $latestFile.FullName
        }
    }
}
catch {
    Write-Host "ERROR: Failed to download or execute custody script" -ForegroundColor Red
    Write-Host "URL: $gitHubRawUrl" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
