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

    # Capture output from the script
    $output = Invoke-Expression $custodyScript

    # Look for the file path in the output
    $formPath = $null
    if ($output) {
        foreach ($line in $output) {
            if ($line -match "FORMPATH:(.+)") {
                $formPath = $matches[1]
                break
            }
        }
    }

    # Give it time to finish writing the file
    Start-Sleep -Seconds 1

    # If we found the path, open it
    if ($formPath -and (Test-Path $formPath)) {
        Write-Host ""
        Write-Host "Opening generated form..." -ForegroundColor Green
        Start-Process $formPath
        Write-Host "File opened successfully!" -ForegroundColor Green
    } elseif ($formPath) {
        Write-Host "File was created at: $formPath" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "ERROR: Failed to download or execute custody script" -ForegroundColor Red
    Write-Host "URL: $gitHubRawUrl" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Start-Sleep -Seconds 3
    exit 1
}
