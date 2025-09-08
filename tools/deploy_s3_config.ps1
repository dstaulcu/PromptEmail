# S3 Bucket Configuration Script for PromptEmail Add-in
# Creates and configures S3 buckets based on deployment-environments.json

param(
    [string]$Environment = "",
    [switch]$AllEnvironments = $false,
    [switch]$DryRun = $false,
    [switch]$Help = $false
)

function Write-Status {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Show-Help {
    Write-Status "S3 Bucket Configuration Script" "Blue"
    Write-Status "=============================" "Blue"
    Write-Status ""
    Write-Status "Creates and configures S3 buckets for PromptEmail add-in deployment" "White"
    Write-Status ""
    Write-Status "Usage:" "Yellow"
    Write-Status "  .\deploy_s3_config.ps1 -Environment Dev" "White"
    Write-Status "  .\deploy_s3_config.ps1 -Environment Test" "White"
    Write-Status "  .\deploy_s3_config.ps1 -Environment Prod" "White"
    Write-Status "  .\deploy_s3_config.ps1 -AllEnvironments" "White"
    Write-Status ""
    Write-Status "Parameters:" "Yellow"
    Write-Status "  -Environment    Specific environment to configure (Dev, Test, Prod)" "White"
    Write-Status "  -AllEnvironments Configure all environments" "White"
    Write-Status "  -DryRun         Show what would be done without executing" "White"
    Write-Status "  -Help           Show this help message" "White"
    Write-Status ""
    Write-Status "Examples:" "Yellow"
    Write-Status "  .\deploy_s3_config.ps1 -Environment Dev -DryRun" "Cyan"
    Write-Status "  .\deploy_s3_config.ps1 -AllEnvironments" "Cyan"
}

function Get-DeploymentEnvironments {
    $configPath = ".\deployment-environments.json"
    
    if (-not (Test-Path $configPath)) {
        Write-Status "❌ Configuration file not found: $configPath" "Red"
        Write-Status "Make sure you're running this script from the tools directory" "Yellow"
        exit 1
    }
    
    try {
        $config = Get-Content $configPath | ConvertFrom-Json
        return $config.environments
    } catch {
        Write-Status "❌ Failed to parse configuration file: $($_.Exception.Message)" "Red"
        exit 1
    }
}

function Test-AWSPrerequisites {
    Write-Status "🔍 Checking AWS prerequisites..." "Blue"
    
    # Check AWS CLI
    try {
        $awsVersion = aws --version 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "AWS CLI not found"
        }
        Write-Status "✅ AWS CLI found: $($awsVersion.Split()[0])" "Green"
    } catch {
        Write-Status "❌ AWS CLI not found. Please install AWS CLI first." "Red"
        exit 1
    }
    
    # Check AWS credentials
    try {
        $identity = aws sts get-caller-identity 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "AWS credentials not configured"
        }
        $identityObj = $identity | ConvertFrom-Json
        Write-Status "✅ AWS credentials configured for: $($identityObj.Arn)" "Green"
    } catch {
        Write-Status "❌ AWS credentials not configured. Run 'aws configure' first." "Red"
        exit 1
    }
}

function Test-BucketExists {
    param([string]$BucketName, [string]$Region)
    
    try {
        aws s3api head-bucket --bucket $BucketName --region $Region 2>$null
        return $LASTEXITCODE -eq 0
    } catch {
        return $false
    }
}

function New-S3Bucket {
    param(
        [string]$BucketName,
        [string]$Region,
        [string]$Environment,
        [bool]$DryRunMode
    )
    
    Write-Status "🪣 Configuring S3 bucket for $Environment environment..." "Blue"
    Write-Status "   Bucket: $BucketName" "White"
    Write-Status "   Region: $Region" "White"
    
    if ($DryRunMode) {
        Write-Status "   [DRY RUN] Would create/configure bucket" "Yellow"
        return
    }
    
    # Check if bucket exists
    if (Test-BucketExists -BucketName $BucketName -Region $Region) {
        Write-Status "   ✅ Bucket already exists" "Green"
    } else {
        Write-Status "   📦 Creating bucket..." "Cyan"
        
        try {
            if ($Region -eq "us-east-1") {
                # us-east-1 doesn't need LocationConstraint
                aws s3api create-bucket --bucket $BucketName --region $Region
            } else {
                # Other regions need LocationConstraint
                aws s3api create-bucket --bucket $BucketName --region $Region --create-bucket-configuration LocationConstraint=$Region
            }
            
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to create bucket"
            }
            Write-Status "   ✅ Bucket created successfully" "Green"
        } catch {
            Write-Status "   ❌ Failed to create bucket: $($_.Exception.Message)" "Red"
            return
        }
    }
    
    # Configure bucket for static website hosting
    Write-Status "   🌐 Configuring static website hosting..." "Cyan"
    try {
        $websiteConfig = @{
            IndexDocument = @{ Suffix = "index.html" }
            ErrorDocument = @{ Key = "index.html" }
        } | ConvertTo-Json -Compress
        
        aws s3api put-bucket-website --bucket $BucketName --website-configuration $websiteConfig
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to configure website hosting"
        }
        Write-Status "   ✅ Website hosting configured" "Green"
    } catch {
        Write-Status "   ❌ Failed to configure website hosting: $($_.Exception.Message)" "Red"
    }
    
    # Configure public read access policy
    Write-Status "   🔓 Configuring public read access..." "Cyan"
    try {
        $policyDocument = @{
            Version = "2012-10-17"
            Statement = @(
                @{
                    Sid = "PublicReadGetObject"
                    Effect = "Allow"
                    Principal = "*"
                    Action = "s3:GetObject"
                    Resource = "arn:aws:s3:::$BucketName/*"
                }
            )
        } | ConvertTo-Json -Depth 4 -Compress
        
        # First, ensure public access block is configured to allow policy
        aws s3api put-public-access-block --bucket $BucketName --public-access-block-configuration "BlockPublicAcls=false,IgnorePublicAcls=false,BlockPublicPolicy=false,RestrictPublicBuckets=false"
        
        # Then apply the bucket policy
        aws s3api put-bucket-policy --bucket $BucketName --policy $policyDocument
        
        if ($LASTEXITCODE -ne 0) {
            throw "Failed to configure bucket policy"
        }
        Write-Status "   ✅ Public read access configured" "Green"
    } catch {
        Write-Status "   ❌ Failed to configure public access: $($_.Exception.Message)" "Red"
        Write-Status "   ℹ️  You may need to manually configure bucket policy in AWS Console" "Yellow"
    }
    
    # Test bucket accessibility
    Write-Status "   🔍 Testing bucket accessibility..." "Cyan"
    try {
        $testUrl = "https://$BucketName.s3.$Region.amazonaws.com"
        $response = Invoke-WebRequest -Uri $testUrl -Method HEAD -TimeoutSec 10 -ErrorAction SilentlyContinue
        
        if ($response.StatusCode -eq 404) {
            # 404 is expected for empty bucket - means it's accessible
            Write-Status "   ✅ Bucket is publicly accessible" "Green"
        } elseif ($response.StatusCode -eq 200) {
            Write-Status "   ✅ Bucket is accessible with content" "Green"
        } else {
            Write-Status "   ⚠️  Bucket response: $($response.StatusCode)" "Yellow"
        }
    } catch {
        Write-Status "   ⚠️  Could not test bucket accessibility (may be normal for new bucket)" "Yellow"
    }
    
    Write-Status "   🎉 Bucket configuration completed!" "Green"
    Write-Status ""
}

function Show-ConfigurationSummary {
    param($Environments)
    
    Write-Status "📋 Configuration Summary:" "Blue"
    Write-Status "========================" "Blue"
    
    foreach ($env in $Environments.PSObject.Properties) {
        $envName = $env.Name
        $envConfig = $env.Value
        $bucketName = $envConfig.s3Uri.host
        $region = $envConfig.region
        $publicUrl = "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)"
        
        Write-Status ""
        Write-Status "Environment: $envName" "Yellow"
        Write-Status "  Bucket: $bucketName" "White"
        Write-Status "  Region: $region" "White"  
        Write-Status "  Public URL: $publicUrl" "Cyan"
    }
    
    Write-Status ""
    Write-Status "Next Steps:" "Yellow"
    Write-Status "1. Run build_and_deploy.ps1 to upload your add-in files" "White"
    Write-Status "2. Update manifest.xml with the correct environment URLs" "White"
    Write-Status "3. Test add-in functionality" "White"
}

# Main execution
if ($Help) {
    Show-Help
    exit 0
}

if (-not $Environment -and -not $AllEnvironments) {
    Write-Status "❌ Please specify -Environment <name> or -AllEnvironments" "Red"
    Write-Status "Use -Help for more information" "Yellow"
    exit 1
}

Write-Status "🚀 S3 Bucket Configuration for PromptEmail Add-in" "Blue"
Write-Status "=================================================" "Blue"
Write-Status ""

# Load configuration and check prerequisites
$environments = Get-DeploymentEnvironments
Test-AWSPrerequisites

# Determine which environments to process
$envsToProcess = @()
if ($AllEnvironments) {
    $envsToProcess = $environments.PSObject.Properties
} else {
    $envConfig = $environments.$Environment
    if (-not $envConfig) {
        Write-Status "❌ Environment '$Environment' not found in configuration" "Red"
        Write-Status "Available environments: $($environments.PSObject.Properties.Name -join ', ')" "Yellow"
        exit 1
    }
    $envsToProcess = @([PSCustomObject]@{ Name = $Environment; Value = $envConfig })
}

# Process each environment
foreach ($env in $envsToProcess) {
    $envName = $env.Name
    $envConfig = $env.Value
    $bucketName = $envConfig.s3Uri.host
    $region = $envConfig.region
    
    New-S3Bucket -BucketName $bucketName -Region $region -Environment $envName -DryRunMode $DryRun
}

# Show summary
Show-ConfigurationSummary -Environments $environments

if ($DryRun) {
    Write-Status "✨ Dry run completed - no changes were made" "Green"
} else {
    Write-Status "✅ S3 bucket configuration completed!" "Green"
}
