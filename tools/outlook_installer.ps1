# PromptEmail - Simple PowerShell 5.1 Installer
param(
    [string]$Environment = "Prod",
    [switch]$Help,
    [switch]$Uninstall
)

function Write-Color {
    param([string]$Text, [string]$Color = "White")
    Write-Host $Text -ForegroundColor $Color
}

if ($Help) {
    Write-Color "PromptEmail Installer for PowerShell 5.1" "Blue"
    Write-Color "Usage: .\outlook_installer.ps1 [-Environment Prod|Dev|Test] [-Uninstall] [-Help]" "White"
    exit 0
}

$InstallPath = Join-Path $env:APPDATA "PromptEmail"
$ManifestUrls = @{
    "Prod" = "https://293354421824-outlook-email-assistant-prod.s3.us-east-1.amazonaws.com/manifest.xml"
    "Dev"  = "https://293354421824-outlook-email-assistant-dev.s3.us-east-1.amazonaws.com/manifest.xml" 
    "Test" = "https://293354421824-outlook-email-assistant-test.s3.us-east-1.amazonaws.com/manifest.xml"
}

if ($Uninstall) {
    Write-Color "Uninstalling PromptEmail..." "Yellow"
    $RegPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\PromptEmail",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer\PromptEmail"
    )
    foreach ($path in $RegPaths) {
        if (Test-Path $path) {
            Remove-Item $path -Force -ErrorAction SilentlyContinue
            Write-Color "Removed: $path" "Green"
        }
    }
    Write-Color "Uninstall complete" "Green"
    exit 0
}

Write-Color "Installing PromptEmail ($Environment)..." "Blue"

# Stop Outlook
$outlook = Get-Process OUTLOOK -ErrorAction SilentlyContinue
if ($outlook) {
    $outlook | Stop-Process -Force
    Write-Color "Stopped Outlook" "Green"
}

# Create install directory
if (!(Test-Path $InstallPath)) {
    New-Item $InstallPath -ItemType Directory -Force | Out-Null
}

# Download manifest
$ManifestFile = Join-Path $InstallPath "manifest.xml"
$Url = $ManifestUrls[$Environment]
try {
    $web = New-Object System.Net.WebClient
    $web.DownloadFile($Url, $ManifestFile)
    Write-Color "Downloaded manifest" "Green"
} catch {
    Write-Color "Download failed: $_" "Red"
    exit 1
}

# Install registry entries
$RegPaths = @(
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer",
    "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer"
)

foreach ($basePath in $RegPaths) {
    $fullPath = Join-Path $basePath "PromptEmail"
    try {
        if (!(Test-Path $basePath)) {
            New-Item $basePath -Force | Out-Null
        }
        if (!(Test-Path $fullPath)) {
            New-Item $fullPath -Force | Out-Null
        }
        Set-ItemProperty $fullPath -Name "(Default)" -Value $ManifestFile
        Write-Color "Registry: $fullPath" "Green"
    } catch {
        Write-Color "Registry failed: $fullPath" "Yellow"
    }
}

Write-Color "Installation complete! Start Outlook and look for PromptEmail in the ribbon." "Green"
