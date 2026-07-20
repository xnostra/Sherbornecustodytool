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
}
catch {
    Write-Host "ERROR: Failed to download or execute custody script" -ForegroundColor Red
    Write-Host "URL: $gitHubRawUrl" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
