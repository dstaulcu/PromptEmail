<#
.SYNOPSIS
    Builds and deploys the PromptEmail Outlook Add-in to AWS S3

.DESCRIPTION
    This script builds the add-in using webpack, optionally increments the version,
    and deploys the assets to AWS S3 with environment-specific URL rewriting.

.PARAMETER Environment
    Deployment environment: 'Dev', 'Test', or 'Prod'

.PARAMETER DryRun
    Show what would be deployed without actually deploying

.PARAMETER IncrementVersion
    Automatically increment the package.json version before building.
    Valid values: 'major', 'minor', 'patch'

.EXAMPLE
    # Deploy to production with patch version increment
    .\deploy_web_assets.ps1 -Environment Prod -IncrementVersion patch

.EXAMPLE
    # Deploy to development without version change
    .\deploy_web_assets.ps1 -Environment Dev

.EXAMPLE
    # Dry run for production with minor version bump
    .\deploy_web_assets.ps1 -Environment Prod -IncrementVersion minor -DryRun
#>

param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Dev', 'Test', 'Prod')]
    [string]$Environment,
    [switch]$DryRun,
    [Parameter(Mandatory = $false)]
    [ValidateSet('major', 'minor', 'patch')]
    [string]$IncrementVersion
)

function Update-EmbeddedUrls {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param (
        [Parameter(Mandatory)]
        [string]$RootPath,
        [string]$NewHost,
        [string]$NewScheme = "https"
    )
    
    # Comprehensive URL pattern to match HTTP/HTTPS/S3 URLs
    $urlPattern = '(?i)(?:https?://|s3://)[a-zA-Z0-9\-\.]+(?:\.[a-zA-Z0-9\-\.]+)*(?:\:[0-9]+)?(?:/[^\s"''<>]*)?'
    
    # Whitelist of URLs to skip warnings and string substitution
    $urlWhitelist = @(
        # Office.js related URLs that should be ignored (managed separately)
        'https://appsforoffice.microsoft.com*',
        '*appsforoffice.microsoft.com*',
        # GitHub URLs that should never be replaced for source, wiki, and issues
        'https://github.com/*',
        '*github.com*',
        # Telemetry URLs that should never be replaced for telemetry
        'https://splunk.company.com*',
        '*splunk.company.com*',
        # Microsoft schema URLs
        'http://schemas.microsoft.com/*',
        '*schemas.microsoft.com*',
        # Generic placeholder/template URLs that should never be replaced
        '*your-username*',
        '*localhost*',
        '*your-organization.com*'
    )

    $files = Get-ChildItem -Path $RootPath -Recurse -File
    $urlResults = @()
    
    foreach ($file in $files) {
        try {
            $content = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
            
            # Find regular URLs
            $urlMatches = [regex]::Matches($content, $urlPattern)
            foreach ($match in $urlMatches) {
                $urlResults += [PSCustomObject]@{
                    File = $file.FullName
                    URL  = $match.Value
                    Type = "DirectURL"
                }
            }
        }
        catch {
            Write-Warning "Could not read file: $($file.FullName)"
        }
    }

    Write-Host "Found $($urlResults.Count) direct URLs"

    # Process direct URLs (existing logic)
    foreach ($url in $urlResults) {
        # Check whitelist (skip if matches any whitelist entry)
        $isWhitelisted = $false
        foreach ($white in $urlWhitelist) {
            if ($white.Contains('*')) {
                # Handle wildcard patterns
                $pattern = $white -replace '\*', '.*'
                if ($url.URL -match $pattern) {
                    $isWhitelisted = $true
                    break
                }
            } elseif ($url.URL -like "$white*") {
                $isWhitelisted = $true
                break
            }
        }
        if ($isWhitelisted) {
            Write-Host "[Whitelist] Skipping URL: $($url.URL) in file: $($url.File)"
            continue
        }

        # Skip URLs in protected configuration files
        if ($url.File -like "*ai-providers.json") {
            Write-Host "[Protected File] Skipping URL: $($url.URL) in ai-providers.json"
            continue
        }

        # Check if any name is contained in the URL
        $fileNames = $files | Select-Object -ExpandProperty name
        # Only match S3 URLs that contain outlook-email-assistant buckets - be more restrictive
        $otherStrings = "293354421824-outlook-email-assistant.*\.s3"
        $matchFound = $fileNames | Where-Object { $url.URL -like "*$_*" }

        # prepare url for POTENTIAL replacement
        $do_replacement = $false
        if ($url.URL -like "s3://*") {
            $originalUri = [System.Uri]("https://" + $url.URL.Substring(6))
            $scheme = "s3"
        }
        else {
            $originalUri = [System.Uri]$url.URL
            $scheme = $originalUri.Scheme
        }

        if ($matchFound) {
            $do_replacement = $true
            Write-Host "Public file match found: $($matchFound -join ', ') for url: $($url.URL) in file: $($url.File)"
            # we need to get the fullpath to file from matching filename
            $matchFoundFullName = ($files | Where-Object { $_.name -eq $matchFound }[0]).FullName
            # Normalize path for reference in URL
            $AbsolutePath_new = $matchFoundFullName -replace ".*\\public", ""
            $AbsolutePath_new = $AbsolutePath_new -replace "\\", "/"
            $AbsolutePath_new = $AbsolutePath_new -replace "^([^/])", "/$1"
        }
        elseif ($url.URL -match $otherStrings) {
            $do_replacement = $true
            Write-Host "String match found for url: $($url.URL) in file: $($url.File)"
            $AbsolutePath_new = $originalUri.AbsolutePath
        }
        # Normalize double slashes (except after protocol) for all replacements
        if ($do_replacement -and $AbsolutePath_new) {
            $AbsolutePath_new = $AbsolutePath_new -replace '^/+', '/'
            $AbsolutePath_new = $AbsolutePath_new -replace '://', '___PROTOCOL_SLASH___'
            $AbsolutePath_new = $AbsolutePath_new -replace '/{2,}', '/'
            $AbsolutePath_new = $AbsolutePath_new -replace '___PROTOCOL_SLASH___', '://'
        }
        else {
            $do_replacement = $false
            if ($url.url -match '\.[^\.]{0,3}$') {
                Write-status "⚠️  No public file name or match found in SEEMINGLY file-oriented url: $($url.URL) in file: $($url.File)" 'yellow'
            } else {
                # Write-status "⚠️  No public file name or match found in SEEMLINGLY folder-oriented url: $($url.URL) in file: $($url.File)"
            }
        }

        if ($do_replacement -eq $true) {
            # Build new URI
            $builder = New-Object System.UriBuilder $originalUri
            $builder.Host = $NewHost
            $builder.Path = $AbsolutePath_new
            $builder.Query = $originalUri.Query
            $newUri = $builder.Uri            

            # Restore s3:// scheme if needed
            $newUriString = $newUri.AbsoluteUri
            if ($scheme -eq "s3") {
                $newUriString = $newUriString -replace "^https://", "s3://"
            }

            # Normalize double slashes in the final URL (except after protocol)
            $newUriString = $newUriString -replace '://', '___PROTOCOL_SLASH___'
            $newUriString = $newUriString -replace '/{2,}', '/'
            $newUriString = $newUriString -replace '___PROTOCOL_SLASH___', '://'

            write-status "`tReplacing direct URL:"
            write-status "`t  Original: `"$($url.URL)`""
            write-status "`t       New: `"$($newUriString)`""
            write-status "`t  Mapping: $($originalUri.Host)$($originalUri.AbsolutePath) → $($NewHost)$($AbsolutePath_new)"

            if ($PSCmdlet.ShouldProcess($url.File, "Replace '$($url.URL)' with '$newUriString'")) {
                $content = Get-Content -Path $url.File -Raw
                $contentUpdated = $content -replace [regex]::Escape($url.URL), $newUriString
                Set-Content -Path $url.File -Value $contentUpdated
            }
        }
    }
}

# Update references in online.orig and copy to online
function Write-Status {
    param([string]$Message, [string]$Color = "White")
    $validColors = @('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')
    if (-not $Color -or ($validColors -notcontains $Color)) {
        $Color = 'White'
    }
    Write-Host $Message -ForegroundColor $Color
}

function Update-PackageVersion {
    param([string]$IncrementType)
    
    Write-Status "Incrementing package version ($IncrementType)..." 'Blue'
    
    $packageJsonPath = Join-Path $PSScriptRoot "..\package.json"
    if (-not (Test-Path $packageJsonPath)) {
        Write-Status "ERROR: package.json not found at $packageJsonPath" 'Red'
        return $false
    }
    
    try {
        # Read and parse package.json
        $packageContent = Get-Content $packageJsonPath -Raw | ConvertFrom-Json
        $currentVersion = $packageContent.version
        
        Write-Status "Current version: $currentVersion" 'Yellow'
        
        # Parse version (major.minor.patch)
        if ($currentVersion -match '^(\d+)\.(\d+)\.(\d+)$') {
            $major = [int]$matches[1]
            $minor = [int]$matches[2] 
            $patch = [int]$matches[3]
            
            # Increment based on type
            switch ($IncrementType) {
                'major' { 
                    $major++; $minor = 0; $patch = 0 
                }
                'minor' { 
                    $minor++; $patch = 0 
                }
                'patch' { 
                    $patch++ 
                }
            }
            
            $newVersion = "$major.$minor.$patch"
            $packageContent.version = $newVersion
            
            # Write back to file with proper formatting
            $packageContent | ConvertTo-Json -Depth 10 | Set-Content $packageJsonPath -Encoding UTF8
            
            Write-Status "Version updated: $currentVersion → $newVersion" 'Green'
            return $true
        }
        else {
            Write-Status "ERROR: Invalid version format in package.json: $currentVersion" 'Red'
            return $false
        }
    }
    catch {
        Write-Status "ERROR: Failed to update package version: $($_.Exception.Message)" 'Red'
        return $false
    }
}

function Test-Prerequisites {
    Write-Status "Checking prerequisites..." 'Blue'
    
    # Check if AWS CLI is installed
    try {
        aws --version | Out-Null
        Write-Status "✓ AWS CLI found" 'Green'
    }
    catch {
        Write-Status "✗ AWS CLI not found. Please install AWS CLI." 'Red'
        exit 1
    }
    
    # Check AWS credentials and permissions
    Write-Status "Validating AWS credentials and permissions..." 'Yellow'
    try {
        # Test basic AWS credential validity
        $whoamiOutput = aws sts get-caller-identity 2>&1
        if ($LASTEXITCODE -eq 0) {
            $identity = $whoamiOutput | ConvertFrom-Json
            Write-Status "✓ AWS credentials are valid" 'Green'
            Write-Status "  Account: $($identity.Account)" 'Cyan'
            Write-Status "  User/Role: $($identity.Arn.Split('/')[-1])" 'Cyan'
        }
        else {
            Write-Status "✗ AWS credentials are invalid or expired." 'Red'
            Write-Status "Error details: $whoamiOutput" 'Red'
            Write-Status "Please run 'aws configure' or refresh your temporary credentials." 'Yellow'
            exit 1
        }
        
        # Test S3 access to the target bucket
        $bucketName = $envConfig.s3Uri.host
        Write-Status "Testing S3 bucket access: $bucketName..." 'Yellow'
        
        $s3TestOutput = aws s3 ls "s3://$bucketName" --region $envConfig.region 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Status "✓ S3 bucket access confirmed" 'Green'
        }
        else {
            Write-Status "✗ Cannot access S3 bucket: $bucketName" 'Red'
            Write-Status "Error details: $s3TestOutput" 'Red'
            Write-Status "Please verify:" 'Yellow'
            Write-Status "  - Bucket exists and you have access" 'Yellow'
            Write-Status "  - Your AWS credentials have S3 permissions" 'Yellow'
            Write-Status "  - The specified region ($($envConfig.region)) is correct" 'Yellow'
            exit 1
        }
    }
    catch {
        Write-Status "✗ Exception during AWS credential validation: $_" 'Red'
        Write-Status "Please verify your AWS configuration and try again." 'Yellow'
        exit 1
    }
        
    Write-Status "✓ All prerequisites met" 'Green'
}

function Deploy-Assets {
    param(
        [Parameter(Mandatory)]
        [string]$BuildDir
    )

    Write-Status "Starting optimized S3 deployment..." $Blue
    Write-Status "Source: $BuildDir" 'Cyan'
    Write-Status "Target: $S3BaseUrl" 'Cyan'

    # Use standard S3 sync (modification time + size) for accurate change detection
    Write-Status "Performing intelligent sync (modification time + size comparison)..." $Blue
    
    if ($DryRun) {
        Write-Status "DRY RUN - showing what would be synced..." $Yellow
        $syncOutput = aws s3 sync $BuildDir $S3BaseUrl --region $Region --delete --dryrun 2>&1
    }
    else {
        Write-Status "Syncing files (uploads changed files, deletes obsolete files)..." $Blue
        $syncOutput = aws s3 sync $BuildDir $S3BaseUrl --region $Region --delete 2>&1
    }

    if ($LASTEXITCODE -eq 0) {
        $uploadCount = 0
        $deleteCount = 0
        $skipCount = 0

        # Parse sync output to count operations
        foreach ($line in $syncOutput) {
            if ($line -match "^upload:") { $uploadCount++ }
            elseif ($line -match "^delete:") { $deleteCount++ }
            elseif ($line -match "^(download|copy):") { $uploadCount++ }
        }

        # Calculate skipped files (files that didn't need updating)
        $totalFiles = (Get-ChildItem -Path $BuildDir -Recurse -File).Count
        $skipCount = $totalFiles - $uploadCount

        if ($DryRun) {
            Write-Status "DRY RUN RESULTS:" $Yellow
            Write-Status "  Files that would be uploaded: $uploadCount" $Yellow
            Write-Status "  Files that would be deleted: $deleteCount" $Yellow
            Write-Status "  Files that would be skipped (unchanged): $skipCount" $Yellow
        }
        else {
            Write-Status "SYNC COMPLETED SUCCESSFULLY!" 'Green'
            Write-Status "  Files uploaded: $uploadCount" 'Green'
            Write-Status "  Files deleted (obsolete): $deleteCount" 'Green'
            Write-Status "  Files skipped (unchanged): $skipCount" 'Cyan'
            
            if ($uploadCount -eq 0 -and $deleteCount -eq 0) {
                Write-Status "  🚀 No changes detected - deployment was super fast!" 'Green'
            }
            else {
                Write-Status "  ⚡ Only changed files were processed - much faster than full upload!" 'Green'
            }
        }

        # Show the actual operations if there were any
        if ($uploadCount -gt 0 -or $deleteCount -gt 0) {
            Write-Status "" # Empty line
            Write-Status "Operations performed:" $Blue
            foreach ($line in $syncOutput) {
                if ($line -match "^(upload|delete|download|copy):") {
                    $operation = $line.Split(':')[0]
                    $file = $line.Substring($line.IndexOf(':') + 1).Trim()
                    
                    switch ($operation) {
                        'upload' { Write-Status "  ✓ Uploaded: $file" 'Green' }
                        'delete' { Write-Status "  🗑️ Deleted: $file" $Yellow }
                        'download' { Write-Status "  ⬇️ Downloaded: $file" 'Cyan' }
                        'copy' { Write-Status "  📋 Copied: $file" 'Cyan' }
                    }
                }
            }
        }

        # Set content types for web assets (only for files that were actually uploaded)
        if (-not $DryRun -and $uploadCount -gt 0) {
            Write-Status "" # Empty line
            Write-Status "Setting correct content types for uploaded web assets..." $Blue

            # Extract uploaded file paths from sync output
            $uploadedFiles = @()
            foreach ($line in $syncOutput) {
                if ($line -match "^upload:.*to (.*)$") {
                    $s3Path = $matches[1]
                    $uploadedFiles += $s3Path
                }
            }

            if ($uploadedFiles.Count -gt 0) {
                $contentTypeMap = @{
                    '.html' = 'text/html'
                    '.js'   = 'application/javascript'  
                    '.json' = 'application/json'
                    '.xml'  = 'text/xml'
                    '.css'  = 'text/css'
                    '.png'  = 'image/png'
                }

                $totalUpdated = 0
                foreach ($extension in $contentTypeMap.Keys) {
                    $contentType = $contentTypeMap[$extension]
                    $matchingFiles = $uploadedFiles | Where-Object { $_ -like "*$extension" }
                    
                    if ($matchingFiles.Count -gt 0) {
                        foreach ($file in $matchingFiles) {
                            $output = aws s3 cp $file $file --region $Region --metadata-directive REPLACE --content-type $contentType 2>&1
                            if ($LASTEXITCODE -eq 0) {
                                $totalUpdated++
                            }
                            else {
                                Write-Status "  ⚠️ Warning: Failed to set content-type for '$file': $output" $Yellow
                            }
                        }
                        Write-Status "  ✓ Set content-type '$contentType' for $($matchingFiles.Count) uploaded $extension files" 'Green'
                    }
                }
                
                if ($totalUpdated -gt 0) {
                    Write-Status "  🎯 Optimized: Updated content-type for $totalUpdated files (only changed files)" 'Green'
                }
            }
            else {
                Write-Status "  ℹ️ No uploaded files detected for content-type updates" 'Cyan'
            }
        }
    }
    else {
        Write-Status "✗ S3 sync failed!" 'Red'
        Write-Status "AWS CLI output: $syncOutput" 'Red'
        
        # Check if this looks like a credential issue
        if ($syncOutput -match "expired|invalid|credentials|token|authentication|authorization") {
            Write-Status "This appears to be an AWS credential issue." $Yellow
            Write-Status "Your credentials may have expired. Please refresh your AWS credentials and try again." $Yellow
        }
        exit 1
    }
}

function Test-Deployment {

    Write-Status "Verifying deployment..." 'Blue'
    $baseUrl = $HttpBaseUrl
    # Test index.html accessibility
    try {
        $response = Invoke-WebRequest -Uri "$baseUrl/index.html" -Method Head -TimeoutSec 10
        if ($response.StatusCode -eq 200) {
            Write-Status "✓ index.html is accessible" 'Green'
        }
        else {
            Write-Status "✗ index.html returned status: $($response.StatusCode)" 'Red'
        }
    }
    catch {
        Write-Status "✗ Failed to verify index.html accessibility" 'Red'
        Write-Status $_.Exception.Message 'Red'
    }
}

function Show-NextSteps {
    Write-Status "`nDeployment Summary:" 'Blue'
    Write-Status "Environment: $Environment" 'Blue'
    Write-Status "Bucket: $BucketName" 'Blue'
    Write-Status "Region: $Region" 'Blue'
    Write-Status "Base URL: $HttpBaseUrl" 'Blue'
    Write-Status "`nNext Steps:" 'Blue'
    Write-Status "1. Validate the manifest: npm run validate-manifest" 
    Write-Status "2. Sideload the manifest in Outlook" 
    Write-Status "3. Test the add-in functionality" 
}

function Update-InstallerUrls {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    
    Write-Status "Updating installer URLs for ALL environments..." 'Blue'
    
    # Path to the installer script
    $installerPath = Join-Path $ProjectRoot 'tools\outlook_installer.ps1'
    
    if (-not (Test-Path $installerPath)) {
        Write-Status "⚠ Installer script not found at: $installerPath" 'Yellow'
        return
    }
    
    # Load deployment environments configuration
    $configPath = Join-Path $ProjectRoot 'tools\deployment-environments.json'
    if (-not (Test-Path $configPath)) {
        Write-Status "⚠ Deployment environments config not found: $configPath" 'Yellow'
        return
    }
    
    try {
        # Read the configuration
        $config = Get-Content $configPath | ConvertFrom-Json
        
        # Read the current installer content
        $installerContent = Get-Content $installerPath -Raw
        $updatedContent = $installerContent
        $totalChanges = 0
        
        # Update URLs for all environments
        foreach ($envName in @('Dev', 'Test', 'Prod')) {
            $envConfig = $config.environments.$envName
            if (-not $envConfig) {
                Write-Status "⚠ Environment '$envName' not found in deployment config" 'Yellow'
                continue
            }
            
            # Build the correct URL based on the environment configuration
            $correctUrl = "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)/manifest.xml"
            
            # Define potential old URLs that might exist (in case URLs got out of sync)
            $possibleOldUrls = @(
                "https://293354421824-outlook-email-assistant-$($envName.ToLower()).s3.us-east-1.amazonaws.com/manifest.xml",
                "https://293354421824-outlook-email-assistant-$($envName.ToLower()).s3.amazonaws.com/manifest.xml",
                # Add any other variations that might exist
                "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)/manifest.xml"
            )
            
            # Try to replace any old URL patterns with the correct one
            $envChanges = 0
            foreach ($oldUrl in $possibleOldUrls) {
                if ($updatedContent -ne $updatedContent.Replace($oldUrl, $correctUrl)) {
                    $updatedContent = $updatedContent.Replace($oldUrl, $correctUrl)
                    $envChanges++
                }
            }
            
            if ($envChanges -gt 0) {
                Write-Status "  ✓ Updated $envName environment URLs ($envChanges changes)" 'Green'
                Write-Status "    New URL: $correctUrl" 'Gray'
                $totalChanges += $envChanges
            } else {
                Write-Status "  ℹ $envName environment URLs already correct" 'Gray'
            }
        }
        
        # Write the updated content if any changes were made
        if ($totalChanges -gt 0) {
            if ($PSCmdlet.ShouldProcess($installerPath, "Update all environment URLs")) {
                Set-Content -Path $installerPath -Value $updatedContent -Encoding UTF8
                Write-Status "✓ Installer script updated with current URLs for all environments ($totalChanges total changes)" 'Green'
            } else {
                Write-Status "[DryRun] Would update installer script with current URLs for all environments" 'Yellow'
            }
        } else {
            Write-Status "ℹ All environment URLs in installer script are already up to date" 'Gray'
        }
    }
    catch {
        Write-Status "✗ Failed to update installer URLs: $_" 'Red'
        throw
    }
}

# Main execution
Write-Status "PromptEmail Outlook Add-in Deployment" 'Blue'
Write-Status "=====================================" 'Blue'

# Ensure $ProjectRoot is set before any usage
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$srcDir = Join-Path $ProjectRoot 'src'
$publicDir = Join-Path $ProjectRoot 'public'

# Initialize environment and config 
$ConfigPath = Join-Path $PSScriptRoot 'deployment-environments.json'
if (Test-Path $ConfigPath) {
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    $envConfig = $config.environments.$Environment
    if ($envConfig.publicUri -and $envConfig.s3Uri -and $envConfig.region) {
        $Region = $envConfig.region
        $PublicBaseUrl = "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)"
        $S3BaseUrl = "$($envConfig.s3Uri.protocol)://$($envConfig.s3Uri.host)"
        # For compatibility with rest of script
        $HttpBaseUrl = $PublicBaseUrl
        $BucketName = $envConfig.s3Uri.host.Split('.')[0]
        
        # Configure Office.js URL from environment config
        if ($envConfig.officeJsUrl) {
            $OfficeJsPath = $envConfig.officeJsUrl
            Write-Status "✓ Using environment-specific Office.js URL: $OfficeJsPath" 'Green'
        } else {
            $OfficeJsPath = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js'
            Write-Status "⚠ No Office.js URL configured for $Environment, using default: $OfficeJsPath" 'Yellow'
        }
    }
    else {
        throw "Missing publicUri, s3Uri, or region for environment '$Environment' in deployment-environments.json."
    }
}
else {
    throw "deployment-environments.json not found in tools/."
}


# Run prerequisites check as early as possible
Test-Prerequisites

# Increment version if requested
if ($IncrementVersion) {
    $versionUpdateSuccess = Update-PackageVersion -IncrementType $IncrementVersion
    if (-not $versionUpdateSuccess) {
        Write-Status "ERROR: Failed to update package version. Deployment aborted." 'Red'
        exit 1
    }
}

# clear the public folder if it already exists
if (test-path -path $publicDir) {
    remove-item -Path $publicDir -recurse
    mkdir -Path $publicDir | Out-Null
}

# Ensure required npm packages are installed
Write-Status "Checking webpack installation..." "Blue"
try {
    # First check if webpack is already available and working
    $webpackVersion = $null
    $webpackCliVersion = $null
    $webpackWorking = $false
    
    try {
        $webpackVersion = & webpack --version 2>$null
        $webpackCliVersion = & webpack-cli --version 2>$null
        if ($webpackVersion -and $webpackCliVersion) {
            $webpackWorking = $true
            Write-Status "✓ webpack is already available (webpack: $webpackVersion, webpack-cli: $webpackCliVersion)" "Green"
        }
    }
    catch {
        Write-Status "webpack not found or not working, will install..." "Yellow"
    }
    
    # Install webpack if not working
    if (-not $webpackWorking) {
        Write-Status "Installing webpack and webpack-cli..." "Yellow"
        $installOutput = & npm install -g webpack webpack-cli --save-dev 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Status "✓ Global webpack packages installed successfully" "Green"
            
            # Verify the installation worked
            try {
                $newWebpackVersion = & webpack --version 2>$null
                $newWebpackCliVersion = & webpack-cli --version 2>$null
                if ($newWebpackVersion -and $newWebpackCliVersion) {
                    Write-Status "✓ Verified webpack installation (webpack: $newWebpackVersion, webpack-cli: $newWebpackCliVersion)" "Green"
                }
                else {
                    Write-Status "⚠️  Webpack installed but verification failed, continuing anyway..." "Yellow"
                }
            }
            catch {
                Write-Status "⚠️  Could not verify webpack installation, continuing anyway..." "Yellow"
            }
        }
        else {
            Write-Status "⚠️  Global webpack install had issues, trying local install..." "Yellow"
            Write-Host $installOutput
            
            # Try local install as fallback
            $localInstallOutput = & npm install webpack webpack-cli --save-dev 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Status "✓ Local webpack packages installed successfully" "Green"
            }
            else {
                Write-Status "✗ Both global and local webpack installs failed. See details below:" "Red"
                Write-Host $localInstallOutput
                exit 1
            }
        }
    }
    
    # Install/update local project dependencies
    Write-Status "Installing/updating local project dependencies..." "Yellow"
    $depsOutput = & npm install 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Status "✓ Local dependencies installed successfully" "Green"
    }
    else {
        Write-Status "✗ Failed to install local dependencies. See details below:" "Red"
        Write-Host $depsOutput
        exit 1
    }
}
catch {
    Write-Status "✗ Exception during package installation: $_" "Red"
    exit 1
}

# Run npm build and capture output
Write-Status "Starting build process..." "Blue"
try {
    $buildOutput = & npm run build 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Status "✓ Build completed successfully" "Green"
    }
    else {
        Write-Status "✗ Build failed. See details below:" "Red"
        Write-Host $buildOutput
        exit 1
    }
}
catch {
    Write-Status "✗ Exception during build: $_" "Red"
    if ($buildOutput) { Write-Host $buildOutput }
    exit 1
}

# Generate environment-specific manifest from template
$srcManifestTemplate = Join-Path $srcDir 'manifest.template.xml'
$srcManifest = Join-Path $srcDir 'manifest.xml'  # Keep for backwards compatibility
$publicManifest = Join-Path $publicDir 'manifest.xml'

if ($DryRun) {
    Write-Status "[DryRun] Would generate environment-specific manifest for $Environment" 'Yellow'
}
elseif (Test-Path $srcManifestTemplate) {
    # Load environment configuration
    $envConfig = $config.environments.$Environment
    $baseUrl = "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)"
    $hostDomain = $envConfig.publicUri.host
    
    # Read template and replace placeholders
    $manifestContent = Get-Content $srcManifestTemplate -Raw
    $manifestContent = $manifestContent.Replace('{{MANIFEST_ID}}', $envConfig.manifestId)
    $manifestContent = $manifestContent.Replace('{{DISPLAY_NAME}}', $envConfig.displayName)
    $manifestContent = $manifestContent.Replace('{{DESCRIPTION}}', $envConfig.description)
    $manifestContent = $manifestContent.Replace('{{GROUP_LABEL}}', $envConfig.groupLabel)
    $manifestContent = $manifestContent.Replace('{{BUTTON_LABEL}}', $envConfig.buttonLabel)
    $manifestContent = $manifestContent.Replace('{{BUTTON_TOOLTIP}}', $envConfig.buttonTooltip)
    $manifestContent = $manifestContent.Replace('{{BASE_URL}}', $baseUrl)
    $manifestContent = $manifestContent.Replace('{{HOST_DOMAIN}}', $hostDomain)
    
    # Write environment-specific manifest to public directory
    $manifestContent | Set-Content $publicManifest -Encoding UTF8
    Write-Status "✓ Generated $Environment environment manifest with ID: $($envConfig.manifestId)" 'Green'
}
elseif (Test-Path $srcManifest) {
    # Fallback to copying existing manifest (backwards compatibility)
    Copy-Item $srcManifest $publicManifest -Force
    Write-Status "✓ Copied manifest.xml to public/ (using legacy manifest)" 'Yellow'
}
else {
    Write-Status "✗ Neither src/manifest.template.xml nor src/manifest.xml found!" 'Red'
}

# Update Office.js path in public/taskpane.html
$publicTaskpaneHtml = Join-Path $publicDir 'taskpane.html'
$srcTaskpaneHtml = Join-Path $srcDir 'taskpane/taskpane.html'
if ($DryRun) {
    Write-Status "[DryRun] Would update Office.js path in public/taskpane.html" 'Yellow'
} elseif (Test-Path $srcTaskpaneHtml) {
    Copy-Item $srcTaskpaneHtml $publicTaskpaneHtml -Force
    $taskpaneContent = Get-Content $publicTaskpaneHtml -Raw
    $officeJsPattern = '<script src="https://[^>]*appsforoffice\.microsoft\.com/lib/1/hosted/office\.js"></script>'
    $newOfficeJsTag = "<script src=`"$OfficeJsPath`"></script>"
    $updatedTaskpaneContent = $taskpaneContent -replace $officeJsPattern, $newOfficeJsTag
    Set-Content $publicTaskpaneHtml $updatedTaskpaneContent
    Write-Status "✓ Updated Office.js path in public/taskpane.html to $OfficeJsPath" 'Green'
} else {
    Write-Status "✗ src/taskpane/taskpane.html not found!" 'Red'
}

# copy src\config\telemetry.json to .\public\config
$srcTelemetryConfig = Join-Path $srcDir 'config\telemetry.json'
$publicTelemetryConfig = Join-Path $publicDir 'config\telemetry.json'
if ($DryRun) {
    Write-Status "[DryRun] Would copy telemetry.json to public/config/" 'Yellow'
}
elseif (Test-Path $srcTelemetryConfig) {
    Copy-Item $srcTelemetryConfig $publicTelemetryConfig -Force
    Write-Status "✓ Copied telemetry.json to public/config/" 'Green'
}
else {
        Write-Status "⚠️ src/config/telemetry.json not found - telemetry configuration will use defaults" 'Yellow'
}


# Update online references before deployment
if ($DryRun) {
    Write-Status "[DryRun] Would update embedded URLs and normalize manifest.xml" 'Yellow'
}
else {
    Write-Status "Updating Urls in public folder files..." 'Blue'
    try {
        # Update embedded URLs in public files using new URI-spec config
        Update-EmbeddedUrls -RootPath (Join-Path $ProjectRoot 'public') -NewHost $envConfig.publicUri.host -NewScheme $envConfig.publicUri.protocol
        Write-Status "✓ Urls in public folder files updated" 'Green'

        # Post-process manifest.xml to normalize all URLs (remove double slashes except after protocol)
        $publicManifestPath = Join-Path $publicDir 'manifest.xml'
        if (Test-Path $publicManifestPath) {
            $manifestContent = Get-Content $publicManifestPath -Raw
            $normalizedManifest = $manifestContent -replace '://', '___PROTOCOL_SLASH___'
            $normalizedManifest = $normalizedManifest -replace '/{2,}', '/'
            $normalizedManifest = $normalizedManifest -replace '___PROTOCOL_SLASH___', '://'
            if ($manifestContent -ne $normalizedManifest) {
                Set-Content $publicManifestPath $normalizedManifest
                Write-Status "✓ Normalized double slashes in manifest.xml URLs" 'Green'
            }
        }
    }
    catch {
        Write-Status "✗ Failed to update Urls in publif folder files: $_" 'Red'
        exit 1
    }
}

# deploy content of .\public to target web server (e.g. s3)
Deploy-Assets -BuildDir $publicDir
# verify index.html is web-accessible in web server
Test-Deployment

# Update installer script with current environment URLs
# This ensures the standalone installer always has the correct URLs for all environments
if (-not $DryRun) {
    Update-InstallerUrls
} else {
    Write-Status "[DryRun] Would update installer URLs for all environments" 'Yellow'
}

Show-NextSteps
