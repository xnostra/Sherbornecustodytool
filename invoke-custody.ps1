<#
.SYNOPSIS
    One-liner launcher for IT Asset Custody Form Tool
    Downloads template and script, then executes the custody form tool.

.DESCRIPTION
    This script can be run with:
    irm https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main/invoke-custody.ps1 | iex

.NOTES
    Version: 2.8
    Author: Sherborne Custody Tool Team
    LastModified: 2026-07-22

.LINK
    https://github.com/xnostra/Sherbornecustodytool
#>

$repoUrl = "https://raw.githubusercontent.com/xnostra/Sherbornecustodytool/main"
$custodyScriptUrl = "$repoUrl/Fill-CustodyForm.ps1"
$templateUrl = "$repoUrl/custody%20form.xlsx"

# Set to $true to auto-email the finished form (you'll get one Microsoft sign-in prompt per run).
# Set to $false to just save/open it on the Desktop, no emailing.
$autoEmail = $true

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

    # Set output to user's Desktop
    $desktopPath = Join-Path $env:USERPROFILE "Desktop"

    # Execute the script from the working directory with explicit paths. We build a scriptblock
    # from its TEXT and call that instead of calling the downloaded FILE directly with "&" -
    # calling a .ps1 FILE is still subject to Execution Policy even inside a one-liner that itself
    # bypassed it to get this far. A scriptblock isn't a file, so Execution Policy never applies,
    # and normal param() binding (-TemplatePath/-OutputFolder) still works exactly as if the file
    # had been called directly. This keeps working even where running scripts is disabled by policy.
    Push-Location $workDir
    $scriptText = Get-Content -Raw -LiteralPath $scriptPath
    $scriptBlock = [scriptblock]::Create($scriptText)
    if ($autoEmail) {
        & $scriptBlock -TemplatePath $templatePath -OutputFolder $desktopPath -EmailForm
    } else {
        & $scriptBlock -TemplatePath $templatePath -OutputFolder $desktopPath
    }
    Pop-Location

    # Fill-CustodyForm.ps1 now opens the finished file itself (works the same way no matter which
    # launcher is used), so nothing further to do here.
    Write-Host ""
    Write-Host "Done! Your form is on the Desktop." -ForegroundColor Green

    # Cleanup temp folder only
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
