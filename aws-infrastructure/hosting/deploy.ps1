#!/usr/bin/env powershell
# Deploy script for S3 hosting configuration
# Configures S3 buckets for static website hosting

param(
    [Parameter(Mandatory=$false)]
    [string]$Environment = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$AllEnvironments,
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory=$false)]
    [switch]$Help
)

$ErrorActionPreference = "Stop"

# Helper function for status messages
function Write-Status {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

# Show help information
if ($Help) {
    Write-Status "S3 Hosting Configuration Script" "Green"
    Write-Status "================================" "Green"
    Write-Status ""
    Write-Status "This script configures S3 buckets for static website hosting."
    Write-Status ""
    Write-Status "USAGE:" "Yellow"
    Write-Status "  .\deploy.ps1 -Environment <name>     Configure specific environment"
    Write-Status "  .\deploy.ps1 -AllEnvironments        Configure all environments"
    Write-Status "  .\deploy.ps1 -DryRun                 Show what would be done"
    Write-Status "  .\deploy.ps1 -Help                   Show this help"
    Write-Status ""
    Write-Status "EXAMPLES:" "Yellow"
    Write-Status "  .\deploy.ps1 -Environment Dev"
    Write-Status "  .\deploy.ps1 -AllEnvironments -DryRun"
    Write-Status ""
    Write-Status "REQUIREMENTS:" "Yellow"
    Write-Status "  - AWS CLI configured with appropriate permissions"
    Write-Status "  - S3 bucket creation and policy permissions"
    Write-Status "  - deployment-environments.json configuration file"
    exit 0
}

if (-not $Environment -and -not $AllEnvironments) {
    Write-Status "Error: Please specify -Environment <name> or -AllEnvironments" "Red"
    Write-Status ""
    
    # Show help automatically
    Write-Status "S3 Hosting Configuration Script" "Green"
    Write-Status "================================" "Green"
    Write-Status ""
    Write-Status "This script configures S3 buckets for static website hosting."
    Write-Status ""
    Write-Status "USAGE:" "Yellow"
    Write-Status "  .\deploy.ps1 -Environment <name>     Configure specific environment"
    Write-Status "  .\deploy.ps1 -AllEnvironments        Configure all environments"
    Write-Status "  .\deploy.ps1 -DryRun                 Show what would be done"
    Write-Status "  .\deploy.ps1 -Help                   Show this help"
    Write-Status ""
    Write-Status "EXAMPLES:" "Yellow"
    Write-Status "  .\deploy.ps1 -Environment Dev"
    Write-Status "  .\deploy.ps1 -AllEnvironments -DryRun"
    Write-Status ""
    Write-Status "REQUIREMENTS:" "Yellow"
    Write-Status "  - AWS CLI configured with appropriate permissions"
    Write-Status "  - S3 bucket creation and policy permissions"
    Write-Status "  - deployment-environments.json configuration file"
    exit 1
}

Write-Status "S3 Bucket Configuration for PromptEmail Add-in" "Blue"
Write-Status "===============================================" "Blue"

# Load deployment environments
$configPath = "..\..\tools\deployment-environments.json"
if (-not (Test-Path $configPath)) {
    Write-Status "Configuration file not found: $configPath" "Red"
    exit 1
}

try {
    $config = Get-Content $configPath | ConvertFrom-Json
    Write-Status "Loaded configuration from: $configPath" "Green"
} catch {
    Write-Status "Failed to load configuration: $($_.Exception.Message)" "Red"
    exit 1
}

# Determine which environments to process
$environmentsToProcess = @()

if ($AllEnvironments) {
    $environmentsToProcess = $config.environments.PSObject.Properties.Name
    Write-Status "Processing all environments: $($environmentsToProcess -join ', ')" "Cyan"
} else {
    if (-not $config.environments.$Environment) {
        Write-Status "Error: Environment '$Environment' not found in configuration" "Red"
        Write-Status "Available environments: $($config.environments.PSObject.Properties.Name -join ', ')" "Yellow"
        Write-Status ""
        
        # Show help automatically
        Write-Status "S3 Hosting Configuration Script" "Green"
        Write-Status "================================" "Green"
        Write-Status ""
        Write-Status "This script configures S3 buckets for static website hosting."
        Write-Status ""
        Write-Status "USAGE:" "Yellow"
        Write-Status "  .\deploy.ps1 -Environment <name>     Configure specific environment"
        Write-Status "  .\deploy.ps1 -AllEnvironments        Configure all environments"
        Write-Status "  .\deploy.ps1 -DryRun                 Show what would be done"
        Write-Status "  .\deploy.ps1 -Help                   Show this help"
        Write-Status ""
        Write-Status "EXAMPLES:" "Yellow"
        Write-Status "  .\deploy.ps1 -Environment Dev"
        Write-Status "  .\deploy.ps1 -AllEnvironments -DryRun"
        Write-Status ""
        Write-Status "REQUIREMENTS:" "Yellow"
        Write-Status "  - AWS CLI configured with appropriate permissions"
        Write-Status "  - S3 bucket creation and policy permissions"
        Write-Status "  - deployment-environments.json configuration file"
        exit 1
    }
    $environmentsToProcess = @($Environment)
    Write-Status "Processing environment: $Environment" "Cyan"
}

# Function to set up S3 bucket
function Set-S3BucketConfiguration {
    param(
        [string]$EnvironmentName,
        [object]$EnvConfig,
        [bool]$IsDryRun
    )
    
    $bucketName = $EnvConfig.s3BucketName
    $region = $EnvConfig.awsRegion
    
    Write-Status "`nConfiguring S3 bucket for $EnvironmentName environment..." "Yellow"
    Write-Status "Bucket: $bucketName" "White"
    Write-Status "Region: $region" "White"
    
    if ($IsDryRun) {
        Write-Status "[DRY RUN] Would configure bucket: $bucketName" "Cyan"
        return
    }
    
    # Check if bucket exists
    try {
        aws s3api head-bucket --bucket $bucketName --region $region 2>$null
        Write-Status "Bucket exists: $bucketName" "Green"
    } catch {
        # Create bucket
        try {
            if ($region -eq "us-east-1") {
                aws s3api create-bucket --bucket $bucketName --region $region | Out-Null
            } else {
                aws s3api create-bucket --bucket $bucketName --region $region --create-bucket-configuration LocationConstraint=$region | Out-Null
            }
            Write-Status "Created bucket: $bucketName" "Green"
        } catch {
            Write-Status "Failed to create bucket: $($_.Exception.Message)" "Red"
            return
        }
    }
    
    # Configure website hosting
    try {
        $websiteConfig = @{
            IndexDocument = @{ Suffix = "index.html" }
            ErrorDocument = @{ Key = "error.html" }
        } | ConvertTo-Json -Depth 3
        
        $websiteConfigFile = [System.IO.Path]::GetTempFileName()
        $websiteConfig | Out-File -FilePath $websiteConfigFile -Encoding UTF8
        
        aws s3api put-bucket-website --bucket $bucketName --website-configuration "file://$websiteConfigFile" --region $region | Out-Null
        Remove-Item $websiteConfigFile -Force
        
        Write-Status "Configured website hosting" "Green"
    } catch {
        Write-Status "Failed to configure website hosting: $($_.Exception.Message)" "Red"
        return
    }
    
    # Configure public read policy
    try {
        $policyDocument = @{
            Version = "2012-10-17"
            Statement = @(
                @{
                    Sid = "PublicReadGetObject"
                    Effect = "Allow"
                    Principal = "*"
                    Action = "s3:GetObject"
                    Resource = "arn:aws:s3:::$bucketName/*"
                }
            )
        } | ConvertTo-Json -Depth 4
        
        $policyFile = [System.IO.Path]::GetTempFileName()
        $policyDocument | Out-File -FilePath $policyFile -Encoding UTF8
        
        aws s3api put-bucket-policy --bucket $bucketName --policy "file://$policyFile" --region $region | Out-Null
        Remove-Item $policyFile -Force
        
        Write-Status "Applied public read policy" "Green"
    } catch {
        Write-Status "Failed to apply bucket policy: $($_.Exception.Message)" "Red"
        return
    }
    
    # Disable block public access (required for public website)
    try {
        aws s3api put-public-access-block --bucket $bucketName --public-access-block-configuration BlockPublicAcls=false,IgnorePublicAcls=false,BlockPublicPolicy=false,RestrictPublicBuckets=false --region $region | Out-Null
        Write-Status "Configured public access settings" "Green"
    } catch {
        Write-Status "Failed to configure public access: $($_.Exception.Message)" "Red"
        return
    }
    
    # Get website URL
    $websiteUrl = "https://$bucketName.s3-website-$region.amazonaws.com"
    Write-Status "Website URL: $websiteUrl" "Cyan"
    
    Write-Status "Bucket configuration completed for $EnvironmentName" "Green"
}

# Function to show configuration summary
function Show-ConfigurationSummary {
    param([array]$Environments)
    
    Write-Status "`nConfiguration Summary:" "Yellow"
    Write-Status "======================" "Yellow"
    
    foreach ($envName in $Environments) {
        $envConfig = $config.environments.$envName
        $bucketName = $envConfig.s3BucketName
        $region = $envConfig.awsRegion
        $websiteUrl = "https://$bucketName.s3-website-$region.amazonaws.com"
        
        Write-Status "`nEnvironment: $envName" "White"
        Write-Status "  Bucket: $bucketName" "Gray"
        Write-Status "  Region: $region" "Gray"
        Write-Status "  Website URL: $websiteUrl" "Gray"
    }
    
    Write-Status "`nNext Steps:" "Yellow"
    Write-Status "1. Deploy your web assets using: ..\..\tools\deploy_web_assets.ps1" "White"
    Write-Status "2. Test your deployment using: .\test.ps1" "White"
    Write-Status "3. Update your manifest files with the correct URLs" "White"
}

# Main execution
try {
    # Process each environment
    foreach ($envName in $environmentsToProcess) {
        $envConfig = $config.environments.$envName
        Set-S3BucketConfiguration -EnvironmentName $envName -EnvConfig $envConfig -IsDryRun $DryRun
    }
    
    # Show summary
    Show-ConfigurationSummary -Environments $environmentsToProcess
    
    if ($DryRun) {
        Write-Status "Dry run completed - no changes were made" "Green"
    } else {
        Write-Status "S3 bucket configuration completed!" "Green"
    }

} catch {
    Write-Status "Deployment failed: $($_.Exception.Message)" "Red"
    exit 1
}

Write-Status "`nFor more information, see: docs/DEPLOYMENT.md" "Cyan"