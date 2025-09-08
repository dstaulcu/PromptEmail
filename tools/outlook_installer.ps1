# Outlook Email Assistant - Windows Installer Script
# This script downloads the manifest from S3, stops Outlook, clears cache, and configures registry for sideloading

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Dev", "Test", "Prod")]
    [string]$Environment = "",  # Empty default - will be determined from registry or fallback to Prod
    
    [Parameter(Mandatory=$false)]
    [string]$InstallPath = "$env:APPDATA\OutlookEmailAssistant",
    
    [Parameter(Mandatory=$false)]
    [switch]$Silent = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$UninstallOnly = $false,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Dev", "Test", "Prod")]
    [string]$SetEnvironmentRegistry = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowEnvironmentRegistry = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$ShowDiagnostics = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$Help = $false
)

# Function to write colored output
function Write-Status {
    param([string]$Message, [string]$Color = "White")
    if (-not $Silent) {
        Write-Host $Message -ForegroundColor $Color
    }
}

# Function to show help
function Show-Help {
    Write-Status "Outlook Email Assistant - Windows Installer" "Blue"
    Write-Status "=============================================" "Blue"
    Write-Status ""
    Write-Status "This script installs the Outlook Email Assistant add-in by:" "White"
    Write-Status "1. Downloading the manifest from S3" "White"
    Write-Status "2. Stopping Outlook processes" "White"
    Write-Status "3. Clearing Outlook add-in cache" "White"
    Write-Status "4. Adding registry keys for sideloading" "White"
    Write-Status ""
    Write-Status "Usage:" "Yellow"
    Write-Status "  .\outlook_installer.ps1 -Environment Prod" "White"
    Write-Status "  .\outlook_installer.ps1 -UninstallOnly" "White"
    Write-Status "  .\outlook_installer.ps1 -SetEnvironmentRegistry Test" "White"
    Write-Status "  .\outlook_installer.ps1 -ShowEnvironmentRegistry" "White"
    Write-Status "  .\outlook_installer.ps1 -ShowDiagnostics" "White"
    Write-Status ""
    Write-Status "Parameters:" "Yellow"
    Write-Status "  -Environment    Environment to install from (Dev, Test, Prod)" "White"
    Write-Status "                  If not specified, checks registry for environment setting" "Gray"
    Write-Status "                  Registry: HKCU\\SOFTWARE\\YourCompany\\OutlookEmailAssistant\\Environment" "Gray"
    Write-Status "                  Default: Prod" "Gray"
    Write-Status "  -InstallPath    Installation directory [Default: %APPDATA%\\OutlookEmailAssistant]" "White"
    Write-Status "  -Silent         Run silently without user prompts" "White"
    Write-Status "  -UninstallOnly  Only remove the add-in, don't install" "White"
    Write-Status "  -SetEnvironmentRegistry  Set environment registry key (Dev, Test, Prod) and exit" "White"
    Write-Status "  -ShowEnvironmentRegistry Show current environment registry settings and exit" "White"
    Write-Status "  -ShowDiagnostics Check current add-in installation status and troubleshooting info" "White"
    Write-Status "  -Help           Show this help message" "White"
    Write-Status ""
    Write-Status "Enterprise Deployment:" "Yellow"
    Write-Status "  For enterprise deployment, set the registry key:" "White"
    Write-Status "  HKEY_CURRENT_USER\\SOFTWARE\\YourCompany\\OutlookEmailAssistant" "Gray"
    Write-Status "  Value: Environment = 'Dev', 'Test', or 'Prod'" "Gray"
    Write-Status ""
    Write-Status "  To set the registry key programmatically:" "White"
    Write-Status "  .\\outlook_installer.ps1 -SetEnvironmentRegistry <Dev|Test|Prod>" "Gray"
    Write-Status ""
}

# Function to check if running as administrator
function Test-Administrator {
    # Not needed for HKCU-only operations, but keeping for compatibility
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to get environment from registry
function Get-EnvironmentFromRegistry {
    param([string]$DefaultEnvironment = "Prod")
    
    # Only check HKCU registry path for low-rights environments
    $registryPath = "HKCU:\SOFTWARE\YourCompany\OutlookEmailAssistant"
    
    try {
        if (Test-Path $registryPath) {
            $environment = Get-ItemProperty -Path $registryPath -Name "Environment" -ErrorAction SilentlyContinue
            if ($environment -and $environment.Environment) {
                $envValue = $environment.Environment
                # Validate environment value
                if ($envValue -in @("Dev", "Test", "Prod")) {
                    Write-Status "Environment '$envValue' detected from user registry: $registryPath" "Green"
                    return $envValue
                } else {
                    Write-Status "Warning: Invalid environment value '$envValue' in registry $registryPath. Must be Dev, Test, or Prod." "Yellow"
                }
            }
        }
    } catch {
        Write-Status "Warning: Could not read registry path $registryPath`: $($_.Exception.Message)" "Yellow"
    }
    
    Write-Status "No valid environment registry key found. Using default: $DefaultEnvironment" "Gray"
    return $DefaultEnvironment
}

# Function to get manifest URL from environment configuration
function Get-ManifestUrl {
    param([string]$Env)
    
    $configPath = Join-Path (Split-Path $PSScriptRoot) "tools\deployment-environments.json"
    if (-not (Test-Path $configPath)) {
        Write-Status "Warning: deployment-environments.json not found, using default URLs" "Yellow"
        switch ($Env) {
            "Dev" { return "https://293354421824-outlook-email-assistant-dev.s3.us-east-1.amazonaws.com/manifest.xml" }
            "Test" { return "https://293354421824-outlook-email-assistant-test.s3.us-east-1.amazonaws.com/manifest.xml" }
            "Prod" { return "https://293354421824-outlook-email-assistant-prod.s3.us-east-1.amazonaws.com/manifest.xml" }
        }
    }
    
    try {
        $config = Get-Content $configPath | ConvertFrom-Json
        $envConfig = $config.environments.$Env
        if ($envConfig -and $envConfig.publicUri) {
            return "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)/manifest.xml"
        }
    } catch {
        Write-Status "Warning: Could not parse deployment configuration: $($_.Exception.Message)" "Yellow"
    }
    
    # Fallback URLs
    switch ($Env) {
        "Dev" { return "https://293354421824-outlook-email-assistant-dev.s3.us-east-1.amazonaws.com/manifest.xml" }
        "Test" { return "https://293354421824-outlook-email-assistant-test.s3.us-east-1.amazonaws.com/manifest.xml" }
        "Prod" { return "https://293354421824-outlook-email-assistant-prod.s3.us-east-1.amazonaws.com/manifest.xml" }
        default { return "https://293354421824-outlook-email-assistant-prod.s3.us-east-1.amazonaws.com/manifest.xml" }
    }
}

# Function to download manifest file
function Get-ManifestFile {
    param([string]$Url, [string]$Destination)
    
    Write-Status "Downloading manifest from: $Url" "Blue"
    
    try {
        # Create destination directory if it doesn't exist
        $destinationDir = Split-Path $Destination -Parent
        if (-not (Test-Path $destinationDir)) {
            New-Item -Path $destinationDir -ItemType Directory -Force | Out-Null
        }
        
        # Download with error handling
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($Url, $Destination)
        
        Write-Status "Manifest downloaded successfully to: $Destination" "Green"
        return $true
    } catch {
        Write-Status "Failed to download manifest: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to validate manifest file
function Test-ManifestFile {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        Write-Status "Manifest file not found: $Path" "Red"
        return $false
    }
    
    try {
        # Read the file content first to check what we got
        $content = Get-Content $Path -Raw
        if ([string]::IsNullOrWhiteSpace($content)) {
            Write-Status "Manifest file is empty" "Red"
            return $false
        }
        
        Write-Status "Manifest file size: $($content.Length) characters" "Gray"
        
        # Try to parse as XML
        [xml]$manifest = $content
        
        # Look for OfficeApp node (handle both with and without namespaces)
        $officeApp = $manifest.SelectSingleNode("//OfficeApp")
        if (-not $officeApp) {
            # Try without namespace prefix
            $officeApp = $manifest.OfficeApp
        }
        if (-not $officeApp) {
            # Try with namespace manager
            $nsManager = New-Object System.Xml.XmlNamespaceManager($manifest.NameTable)
            $nsManager.AddNamespace("o", "http://schemas.microsoft.com/office/appforoffice/1.1")
            $officeApp = $manifest.SelectSingleNode("//o:OfficeApp", $nsManager)
        }
        
        if ($officeApp) {
            Write-Status "Manifest validation passed - OfficeApp node found" "Green"
            return $true
        } else {
            Write-Status "Invalid manifest format - OfficeApp node not found" "Red"
            Write-Status "Root element: $($manifest.DocumentElement.LocalName)" "Gray"
            Write-Status "Root namespace: $($manifest.DocumentElement.NamespaceURI)" "Gray"
            return $false
        }
    } catch {
        Write-Status "Manifest validation failed: $($_.Exception.Message)" "Red"
        Write-Status "First 200 characters of file content:" "Gray"
        try {
            $content = Get-Content $Path -Raw
            $preview = $content.Substring(0, [Math]::Min(200, $content.Length))
            Write-Status $preview "Gray"
        } catch {
            Write-Status "Could not read file for preview" "Gray"
        }
        return $false
    }
}

# Function to stop Outlook processes
function Stop-OutlookProcesses {
    Write-Status "Stopping Outlook processes..." "Blue"
    
    $outlookProcesses = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    if ($outlookProcesses) {
        $outlookProcesses | Stop-Process -Force
        Write-Status "Outlook processes stopped" "Green"
        Start-Sleep -Seconds 3
    } else {
        Write-Status "No Outlook processes found" "Gray"
    }
}

# Function to clear Outlook cache
function Clear-OutlookCache {
    Write-Status "Clearing Outlook add-in cache..." "Blue"
    
    $cachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\15.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\14.0\Wef"
    )
    
    foreach ($cachePath in $cachePaths) {
        if (Test-Path $cachePath) {
            $maxRetries = 5
            $retryDelay = 2
            $attempt = 0
            $success = $false
            
            while (-not $success -and $attempt -lt $maxRetries) {
                try {
                    Remove-Item "$cachePath\*" -Recurse -Force -ErrorAction Stop
                    $success = $true
                    Write-Status "Cleared cache: $cachePath" "Green"
                } catch {
                    Write-Status "Attempt $($attempt + 1) to clear cache failed: $($_.Exception.Message)" "Yellow"
                    Start-Sleep -Seconds $retryDelay
                    $attempt++
                }
            }
            
            if (-not $success) {
                Write-Status "Warning: Some cache files could not be deleted: $cachePath" "Yellow"
            }
        }
    }
}

# Function to add registry keys for sideloading
function Add-SideloadRegistryKeys {
    param(
        [string]$ManifestPath, 
        [string]$AddInName = "OutlookEmailAssistant",
        [string]$ManifestUrl = ""
    )
    
    Write-Status "Adding registry keys for sideloading (using Process Monitor insights)..." "Blue"
    
    # Registry paths that Outlook actually checks (from Process Monitor log)
    $sideloadPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\OutlookSideloadManifestPath",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer\OutlookSideloadManifestPath"
    )
    
    # Fallback to traditional developer paths
    $fallbackPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer"
    )
    
    $success = $false
    
    # Register in the specific path Outlook checks for sideloading
    Write-Status "Setting OutlookSideloadManifestPath (the path Outlook actually checks)..." "Gray"
    foreach ($regPath in $sideloadPaths) {
        try {
            # Create the specific OutlookSideloadManifestPath key that Outlook checks
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
                Write-Status "Created sideload manifest path: $regPath" "Green"
            }
            
            # Set the default value to our manifest path (Outlook checks the default value)
            Set-ItemProperty -Path $regPath -Name "(Default)" -Value $ManifestPath -Type String
            Write-Status "Set sideload manifest: $regPath = $ManifestPath" "Green"
            $success = $true
            
        } catch {
            Write-Status "Warning: Could not set sideload path at $regPath`: $($_.Exception.Message)" "Yellow"
        }
    }
    
    # Also register in traditional developer paths as fallback
    Write-Status "Setting fallback developer registry entries..." "Gray"
    foreach ($regPath in $fallbackPaths) {
        try {
            # Create the Developer key if it doesn't exist
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force | Out-Null
                Write-Status "Created fallback registry path: $regPath" "Green"
            }
            
            # Add the manifest path with our add-in name
            Set-ItemProperty -Path $regPath -Name $AddInName -Value $ManifestPath -Type String
            Write-Status "Added fallback registry entry: $regPath\$AddInName = $ManifestPath" "Green"
            $success = $true
            
            # Reduce developer warnings (optional)
            Set-ItemProperty -Path $regPath -Name "SkipSecurityWarnings" -Value 1 -Type DWord -ErrorAction SilentlyContinue
            
        } catch {
            Write-Status "Warning: Could not add fallback registry key at $regPath`: $($_.Exception.Message)" "Yellow"
        }
    }
    
    # Try alternative sideload locations that Outlook might prefer
    try {
        # Some versions of Outlook look for manifests in the Wef folder directly
        $wefPath = "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef"
        if (Test-Path $wefPath) {
            $alternativeManifest = Join-Path $wefPath "$AddInName.xml"
            Copy-Item $ManifestPath $alternativeManifest -Force -ErrorAction SilentlyContinue
            Write-Status "Copied manifest to WEF cache location: $alternativeManifest" "Gray"
        }
    } catch {
        Write-Status "Warning: Could not copy to WEF cache: $($_.Exception.Message)" "Yellow"
    }
    
    # Add WEF trust settings
    try {
        $wefTrustPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\0"
        if (-not (Test-Path $wefTrustPath)) {
            New-Item -Path $wefTrustPath -Force | Out-Null
            Write-Status "Created WEF trust catalog path" "Gray"
        }
        
        # Trust the S3 domain
        if (-not [string]::IsNullOrEmpty($ManifestUrl)) {
            $manifestDomain = ([System.Uri]$ManifestUrl).Host
            Set-ItemProperty -Path $wefTrustPath -Name "Id" -Value "{$([System.Guid]::NewGuid().ToString())}" -Type String -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $wefTrustPath -Name "Url" -Value "https://$manifestDomain/" -Type String -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $wefTrustPath -Name "Flags" -Value 1 -Type DWord -ErrorAction SilentlyContinue
            Write-Status "Added trusted catalog for domain: $manifestDomain" "Gray"
        }
        
    } catch {
        Write-Status "Warning: Could not set WEF trust settings: $($_.Exception.Message)" "Yellow"
    }
    
    if ($success) {
        Write-Status "Registry keys added successfully using Process Monitor insights!" "Green"
        Write-Status ""
        Write-Status "KEY INSIGHT: Used OutlookSideloadManifestPath registry key" "Yellow"
        Write-Status "This is the exact path Outlook checks during startup" "White"
        Write-Status ""
        return $true
    } else {
        Write-Status "Failed to add registry keys" "Red"
        return $false
    }
}

# Function to remove registry keys (for uninstall)
function Remove-SideloadRegistryKeys {
    param([string]$AddInName = "OutlookEmailAssistant")
    
    Write-Status "Removing registry keys..." "Blue"
    
    # Remove the primary OutlookSideloadManifestPath keys that Outlook actually checks
    $sideloadPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\OutlookSideloadManifestPath",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer\OutlookSideloadManifestPath",
        "HKCU:\SOFTWARE\Microsoft\Office\14.0\WEF\Developer\OutlookSideloadManifestPath"
    )
    
    foreach ($regPath in $sideloadPaths) {
        try {
            if (Test-Path $regPath) {
                # Remove the entire OutlookSideloadManifestPath key
                Remove-Item -Path $regPath -Recurse -Force -ErrorAction Stop
                Write-Status "Removed sideload registry path: $regPath" "Green"
            }
        } catch {
            Write-Status "Warning: Could not remove sideload registry path $regPath`: $($_.Exception.Message)" "Yellow"
        }
    }
    
    # Remove the fallback developer registry entries
    $fallbackPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\14.0\WEF\Developer"
    )
    
    foreach ($regPath in $fallbackPaths) {
        try {
            if (Test-Path $regPath) {
                $property = Get-ItemProperty -Path $regPath -Name $AddInName -ErrorAction SilentlyContinue
                if ($property) {
                    Remove-ItemProperty -Path $regPath -Name $AddInName -ErrorAction Stop
                    Write-Status "Removed fallback registry entry: $regPath\$AddInName" "Green"
                }
            }
        } catch {
            Write-Status "Warning: Could not remove fallback registry key: $($_.Exception.Message)" "Yellow"
        }
    }
    
    # Remove WEF trust settings
    try {
        $wefTrustPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\0"
        if (Test-Path $wefTrustPath) {
            Remove-Item -Path $wefTrustPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Status "Removed WEF trust catalog" "Green"
        }
    } catch {
        Write-Status "Warning: Could not remove WEF trust settings: $($_.Exception.Message)" "Yellow"
    }
    
    # Remove Outlook's internal add-in cache and registration
    Write-Status "Removing Outlook's internal add-in registrations..." "Gray"
    
    # Remove from Office add-ins registry (where Outlook stores installed add-ins)
    $officeAddInPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\AddIns",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\Wef\AddIns"
    )
    
    foreach ($addInPath in $officeAddInPaths) {
        try {
            if (Test-Path $addInPath) {
                # Look for any subkeys that might contain our add-in
                $subKeys = Get-ChildItem -Path $addInPath -ErrorAction SilentlyContinue
                foreach ($subKey in $subKeys) {
                    try {
                        $props = Get-ItemProperty -Path $subKey.PSPath -ErrorAction SilentlyContinue
                        # Check if this registry entry references our add-in by name or manifest
                        if ($props -and (
                            ($props.PSObject.Properties.Name -contains "DisplayName" -and $props.DisplayName -like "*Email*Assistant*") -or
                            ($props.PSObject.Properties.Name -contains "Id" -and $props.Id -like "*OutlookEmail*") -or
                            ($props.PSObject.Properties.Name -contains "SourceLocation" -and $props.SourceLocation -like "*manifest*")
                        )) {
                            Remove-Item -Path $subKey.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                            Write-Status "Removed Office add-in registration: $($subKey.PSPath)" "Green"
                        }
                    } catch {
                        # Continue with other subkeys
                    }
                }
            }
        } catch {
            Write-Status "Warning: Could not access Office add-ins registry at $addInPath" "Yellow"
        }
    }
    
    # Clear any cached manifest files that Outlook copies internally
    $outlookCachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\15.0\Wef",
        "$env:LOCALAPPDATA\Microsoft\Office\14.0\Wef",
        "$env:APPDATA\Microsoft\Office\16.0\Wef",
        "$env:APPDATA\Microsoft\AddIns"
    )
    
    foreach ($cachePath in $outlookCachePaths) {
        try {
            if (Test-Path $cachePath) {
                # Look for any files related to our add-in
                $addInFiles = Get-ChildItem -Path $cachePath -Recurse -ErrorAction SilentlyContinue | Where-Object { 
                    $_.Name -like "*OutlookEmail*" -or $_.Name -like "*manifest*" 
                }
                foreach ($file in $addInFiles) {
                    try {
                        Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
                        Write-Status "Removed cached add-in file: $($file.FullName)" "Green"
                    } catch {
                        # Continue with other files
                    }
                }
            }
        } catch {
            Write-Status "Warning: Could not clear cache path $cachePath" "Yellow"
        }
    }
}

# Function to cleanup installation files
function Remove-InstallationFiles {
    param([string]$InstallPath)
    
    Write-Status "Cleaning up installation files..." "Blue"
    
    try {
        if (Test-Path $InstallPath) {
            Remove-Item $InstallPath -Recurse -Force
            Write-Status "Installation directory removed: $InstallPath" "Green"
        }
    } catch {
        Write-Status "Warning: Could not remove installation directory: $($_.Exception.Message)" "Yellow"
    }
}

# Function to verify installation
function Test-Installation {
    param([string]$ManifestPath)
    
    Write-Status "Verifying installation..." "Blue"
    
    # Check if manifest file exists
    if (-not (Test-Path $ManifestPath)) {
        Write-Status "Installation verification failed: Manifest file not found" "Red"
        return $false
    }
    
    # Check if registry keys exist
    $registryPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer"
    )
    
    $registryFound = $false
    foreach ($regPath in $registryPaths) {
        try {
            if (Test-Path $regPath) {
                $property = Get-ItemProperty -Path $regPath -Name "OutlookEmailAssistant" -ErrorAction SilentlyContinue
                if ($property) {
                    $registryFound = $true
                    break
                }
            }
        } catch {
            # Continue checking other paths
        }
    }
    
    if ($registryFound) {
        Write-Status "Installation verification passed" "Green"
        return $true
    } else {
        Write-Status "Installation verification failed: Registry keys not found" "Red"
        return $false
    }
}

# Main installation function
function Install-OutlookEmailAssistant {
    param([string]$ManifestUrl, [string]$InstallPath)
    
    Write-Status "Starting Outlook Email Assistant installation..." "Blue"
    Write-Status "Environment: $Environment" "Gray"
    Write-Status "Manifest URL: $ManifestUrl" "Gray"
    Write-Status "Install Path: $InstallPath" "Gray"
    Write-Status ""
    
    # Step 1: Download manifest
    $manifestPath = Join-Path $InstallPath "manifest.xml"
    if (-not (Get-ManifestFile -Url $ManifestUrl -Destination $manifestPath)) {
        throw "Failed to download manifest file"
    }
    
    # Step 2: Validate manifest
    if (-not (Test-ManifestFile -Path $manifestPath)) {
        throw "Invalid manifest file"
    }
    
    # Step 3: Stop Outlook processes
    Stop-OutlookProcesses
    
    # Step 4: Clear cache
    Clear-OutlookCache
    
    # Step 5: Add registry keys
    if (-not (Add-SideloadRegistryKeys -ManifestPath $manifestPath -ManifestUrl $ManifestUrl)) {
        Write-Status "Warning: Registry keys may not have been added properly" "Yellow"
    }
    
    # Step 6: Verify installation
    if (Test-Installation -ManifestPath $manifestPath) {
        Write-Status "Installation completed successfully!" "Green"
        Write-Status ""
        Write-Status "The Outlook Email Assistant add-in has been installed." "White"
        Write-Status "Start Outlook to see the add-in in the ribbon." "White"
        Write-Status ""
        Write-Status "TROUBLESHOOTING: If the add-in doesn't appear in Outlook:" "Yellow"
        Write-Status "1. Completely close and restart Outlook (don't just minimize)" "White"
        Write-Status "2. Check Windows Event Viewer > Applications and Services > Microsoft Office Alerts" "White"
        Write-Status "3. Try manual installation via Outlook:" "White"
        Write-Status "   - File > Manage Add-ins > My add-ins > Custom add-ins > Add from file" "Gray"
        Write-Status "   - Browse to: $manifestPath" "Gray"
        Write-Status "4. Alternative: Upload to Office 365 admin center for organization-wide deployment" "White"
        Write-Status ""
        
        if (-not $Silent) {
            $response = Read-Host "Would you like to start Outlook now? (y/n)"
            if ($response -eq 'y' -or $response -eq 'Y' -or $response -eq '') {
                Write-Status "Starting Outlook..." "Blue"
                Start-Process OUTLOOK.EXE
            } else {
                # Show diagnostics if not starting Outlook immediately
                Show-AddInDiagnostics -ManifestPath $manifestPath -ManifestUrl $ManifestUrl
            }
        } else {
            # Always show diagnostics in silent mode
            Show-AddInDiagnostics -ManifestPath $manifestPath -ManifestUrl $ManifestUrl
        }
        return $true
    } else {
        throw "Installation verification failed"
    }
}

# Main uninstall function
function Uninstall-OutlookEmailAssistant {
    param([string]$InstallPath)
    
    Write-Status "Starting Outlook Email Assistant uninstallation..." "Blue"
    Write-Status "This will perform a DEEP cleanup to remove all traces of the add-in" "Yellow"
    
    # Step 1: Stop Outlook processes (multiple times if needed)
    Write-Status "Ensuring Outlook is completely stopped..." "Blue"
    for ($i = 1; $i -le 3; $i++) {
        Stop-OutlookProcesses
        if (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) {
            break
        }
        Start-Sleep -Seconds 2
    }
    
    # Step 2: Clear cache aggressively 
    Clear-OutlookCache
    
    # Step 3: Remove registry keys (now includes internal Outlook caches)
    Remove-SideloadRegistryKeys
    
    # Step 4: Additional deep cleanup
    Write-Status "Performing additional cleanup for stubborn add-in caches..." "Blue"
    
    # Clear Office add-in store cache
    try {
        $officeStorePath = "$env:LOCALAPPDATA\Microsoft\Office\AddInTelemetryCache"
        if (Test-Path $officeStorePath) {
            Remove-Item "$officeStorePath\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Status "Cleared Office add-in telemetry cache" "Green"
        }
    } catch {
        Write-Status "Warning: Could not clear Office add-in store cache" "Yellow"
    }
    
    # Clear Outlook form cache (sometimes stores add-in references)
    try {
        $outlookFormCache = "$env:LOCALAPPDATA\Microsoft\FORMS"
        if (Test-Path $outlookFormCache) {
            $addInForms = Get-ChildItem -Path $outlookFormCache -Recurse -ErrorAction SilentlyContinue | 
                         Where-Object { $_.Name -like "*Email*Assistant*" }
            foreach ($form in $addInForms) {
                Remove-Item -Path $form.FullName -Recurse -Force -ErrorAction SilentlyContinue
                Write-Status "Removed Outlook form cache: $($form.Name)" "Green"
            }
        }
    } catch {
        Write-Status "Warning: Could not clear Outlook form cache" "Yellow"
    }
    
    # Step 5: Remove installation files
    Remove-InstallationFiles -InstallPath $InstallPath
    
    Write-Status ""
    Write-Status "DEEP UNINSTALL COMPLETED!" "Green"
    Write-Status "All traces of the add-in should now be removed." "White"
    Write-Status ""
    Write-Status "IMPORTANT: You must now:" "Yellow"
    Write-Status "1. Completely restart your computer (or at least sign out/in)" "White"
    Write-Status "2. Or alternatively, restart Outlook and wait 30 seconds before checking" "White"
    Write-Status "3. If add-in still appears, check OWA (Outlook Web Access) and remove it there" "White"
    Write-Status ""
    
    return $true
}

# Function to set environment registry key (for enterprise deployment)
function Set-EnvironmentRegistry {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Dev", "Test", "Prod")]
        [string]$Environment
    )
    
    # Always use HKCU for low-rights environments
    $registryPath = "HKCU:\SOFTWARE\YourCompany\OutlookEmailAssistant"
    
    try {
        # Create registry key if it doesn't exist
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force | Out-Null
            Write-Status "Created registry key: $registryPath" "Green"
        }
        
        # Set environment value
        Set-ItemProperty -Path $registryPath -Name "Environment" -Value $Environment -Type String
        Write-Status "Set Environment = '$Environment' in $registryPath" "Green"
        
        # Verify the setting
        $verification = Get-ItemProperty -Path $registryPath -Name "Environment" -ErrorAction SilentlyContinue
        if ($verification -and $verification.Environment -eq $Environment) {
            Write-Status "Registry setting verified successfully" "Green"
            return $true
        } else {
            Write-Status "Warning: Could not verify registry setting" "Yellow"
            return $false
        }
        
    } catch {
        Write-Status "Error setting registry key: $($_.Exception.Message)" "Red"
        return $false
    }
}

# Function to show add-in diagnostics
function Show-AddInDiagnostics {
    param([string]$ManifestPath, [string]$ManifestUrl)
    
    Write-Status ""
    Write-Status "Add-in Installation Diagnostics" "Blue"
    Write-Status "===============================" "Blue"
    Write-Status ""
    
    # Check manifest file
    Write-Status "Manifest File Check:" "Yellow"
    if (Test-Path $ManifestPath) {
        Write-Status "✓ Manifest file exists: $ManifestPath" "Green"
        $fileSize = (Get-Item $ManifestPath).Length
        Write-Status "  File size: $fileSize bytes" "Gray"
    } else {
        Write-Status "✗ Manifest file not found: $ManifestPath" "Red"
    }
    
    # Check registry entries
    Write-Status ""
    Write-Status "Registry Entries Check:" "Yellow"
    
    # Check the specific paths Outlook looks for (from Process Monitor log)
    $sideloadPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\OutlookSideloadManifestPath",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer\OutlookSideloadManifestPath"
    )
    
    foreach ($path in $sideloadPaths) {
        if (Test-Path $path) {
            $entry = Get-ItemProperty -Path $path -Name "(Default)" -ErrorAction SilentlyContinue
            if ($entry -and $entry."(Default)") {
                Write-Status "✓ Sideload manifest path found: $path" "Green"
                Write-Status "  Value: $($entry."(Default)")" "Gray"
            } else {
                Write-Status "⚠ Sideload path exists but no default value: $path" "Yellow"
            }
        } else {
            Write-Status "✗ Sideload path not found: $path" "Red"
        }
    }
    
    # Check fallback developer paths
    $devPaths = @(
        "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer",
        "HKCU:\SOFTWARE\Microsoft\Office\15.0\WEF\Developer"
    )
    
    foreach ($path in $devPaths) {
        if (Test-Path $path) {
            $entry = Get-ItemProperty -Path $path -Name "OutlookEmailAssistant" -ErrorAction SilentlyContinue
            if ($entry) {
                Write-Status "✓ Fallback registry entry found: $path\OutlookEmailAssistant" "Green"
                Write-Status "  Value: $($entry.OutlookEmailAssistant)" "Gray"
            } else {
                Write-Status "⚠ Fallback path exists but no OutlookEmailAssistant entry: $path" "Yellow"
            }
        } else {
            Write-Status "✗ Fallback registry path not found: $path" "Yellow"
        }
    }
    
    # Check trust settings
    Write-Status ""
    Write-Status "Trust Settings Check:" "Yellow"
    $trustPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\0"
    if (Test-Path $trustPath) {
        $trust = Get-ItemProperty -Path $trustPath -ErrorAction SilentlyContinue
        if ($trust -and $trust.Url) {
            Write-Status "✓ Trusted catalog configured: $($trust.Url)" "Green"
        } else {
            Write-Status "✗ Trusted catalog entry incomplete" "Yellow"
        }
    } else {
        Write-Status "✗ Trusted catalog not configured" "Yellow"
    }
    
    # Check Outlook processes
    Write-Status ""
    Write-Status "Outlook Process Check:" "Yellow"
    $outlookProc = Get-Process OUTLOOK -ErrorAction SilentlyContinue
    if ($outlookProc) {
        Write-Status "⚠ Outlook is currently running - restart required for add-in to load" "Yellow"
        Write-Status "  PID: $($outlookProc.Id)" "Gray"
    } else {
        Write-Status "✓ Outlook is not running - ready for restart" "Green"
    }
    
    Write-Status ""
    Write-Status "Next Steps:" "Yellow"
    Write-Status "1. Completely close Outlook if running" "White"
    Write-Status "2. Start Outlook and check the ribbon for 'PromptEmail'" "White"
    Write-Status "3. Manual fallback: File > Manage Add-ins > My add-ins > Custom" "White"
    Write-Status ""
}

# Main script execution
try {
    # Show help if requested
    if ($Help) {
        Show-Help
        exit 0
    }
    
    # Handle registry setting mode
    if (-not [string]::IsNullOrEmpty($SetEnvironmentRegistry)) {
        Write-Status "Setting environment registry key..." "Blue"
        Write-Status "Registry scope: Current User (HKCU) - suitable for low-rights environments" "Gray"
        
        if (Set-EnvironmentRegistry -Environment $SetEnvironmentRegistry) {
            Write-Status "Environment registry key set successfully!" "Green"
            Write-Status "Future installations will use environment: $SetEnvironmentRegistry" "Green"
        } else {
            Write-Status "Failed to set environment registry key" "Red"
            exit 1
        }
        exit 0
    }
    
    # Handle showing current registry settings
    if ($ShowEnvironmentRegistry) {
        Write-Status "Current Environment Registry Settings (HKCU)" "Blue"
        Write-Status "=============================================" "Blue"
        Write-Status ""
        
        $registryPath = "HKCU:\SOFTWARE\YourCompany\OutlookEmailAssistant"
        
        try {
            if (Test-Path $registryPath) {
                $environment = Get-ItemProperty -Path $registryPath -Name "Environment" -ErrorAction SilentlyContinue
                if ($environment -and $environment.Environment) {
                    Write-Status "Current User Registry: $($environment.Environment)" "Green"
                    Write-Status ""
                    Write-Status "Active environment: $($environment.Environment)" "Green"
                } else {
                    Write-Status "Registry key exists but no Environment value found" "Yellow"
                    Write-Status ""
                    Write-Status "Default 'Prod' environment would be used" "Yellow"
                }
            } else {
                Write-Status "No environment registry key found" "Gray"
                Write-Status ""
                Write-Status "Default 'Prod' environment would be used" "Yellow"
            }
        } catch {
            Write-Status "Could not access registry: $($_.Exception.Message)" "Red"
        }
        exit 0
    }
    
    # Handle showing diagnostics
    if ($ShowDiagnostics) {
        Write-Status "Add-in Diagnostics Mode" "Blue"
        $Environment = if ([string]::IsNullOrEmpty($Environment)) { Get-EnvironmentFromRegistry -DefaultEnvironment "Prod" } else { $Environment }
        $ManifestUrl = Get-ManifestUrl -Env $Environment
        $manifestPath = Join-Path $InstallPath "manifest.xml"
        Show-AddInDiagnostics -ManifestPath $manifestPath -ManifestUrl $ManifestUrl
        exit 0
    }
    if (-not (Test-Administrator)) {
        Write-Status "Note: Running with standard user privileges (recommended for low-rights environments)" "Green"
    }
    
    # Handle uninstall-only mode
    if ($UninstallOnly) {
        Uninstall-OutlookEmailAssistant -InstallPath $InstallPath
        exit 0
    }
    
    # Determine environment (registry check if not specified via parameter)
    if ([string]::IsNullOrEmpty($Environment)) {
        $Environment = Get-EnvironmentFromRegistry -DefaultEnvironment "Prod"
    } else {
        Write-Status "Environment '$Environment' specified via parameter" "Green"
    }
    
    # Determine manifest URL from environment
    $ManifestUrl = Get-ManifestUrl -Env $Environment
    
    # Confirm installation if not silent
    if (-not $Silent) {
        Write-Status "Outlook Email Assistant Installer" "Blue"
        Write-Status "=================================" "Blue"
        Write-Status ""
        Write-Status "This will install the Outlook Email Assistant add-in:" "White"
        Write-Status "- Environment: $Environment" "Gray"
        Write-Status "- Manifest URL: $ManifestUrl" "Gray"
        Write-Status "- Install Path: $InstallPath" "Gray"
        Write-Status ""
        Write-Status "The installer will:" "White"
        Write-Status "1. Download the manifest file from S3" "White"
        Write-Status "2. Stop any running Outlook processes" "White"
        Write-Status "3. Clear the Outlook add-in cache" "White"
        Write-Status "4. Add registry keys for sideloading" "White"
        Write-Status ""
        
        $response = Read-Host "Continue with installation? (y/n)"
        if ($response -ne 'y' -and $response -ne 'Y' -and $response -ne '') {
            Write-Status "Installation cancelled by user" "Yellow"
            exit 0
        }
        Write-Status ""
    }
    
    # Perform installation
    Install-OutlookEmailAssistant -ManifestUrl $ManifestUrl -InstallPath $InstallPath
    
} catch {
    Write-Status "Installation failed: $($_.Exception.Message)" "Red"
    Write-Status ""
    Write-Status "Troubleshooting tips:" "Yellow"
    Write-Status "1. Run as Administrator" "White"
    Write-Status "2. Ensure internet connectivity" "White"
    Write-Status "3. Close all Office applications" "White"
    Write-Status "4. Check if Windows Defender or antivirus is blocking the script" "White"
    
    exit 1
}
