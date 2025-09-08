# AWS Infrastructure Provisioning Script for Splunk Gateway
# This script creates an API Gateway with Lambda functions to proxy Splunk HEC requests

param(
    [Parameter(Mandatory=$true)]
    [string]$StackName = "outlook-assistant-splunk-gateway",
    
    [Parameter(Mandatory=$true)]
    [string]$SplunkHecToken,
    
    [Parameter(Mandatory=$true)]
    [string]$SplunkHecUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$Region = "us-east-1",
    
    [Parameter(Mandatory=$false)]
    [string]$Environment = "prod",
    
    [Parameter(Mandatory=$false)]
    [string]$AllowedOrigin = "https://your-account-outlook-assistant-prod.s3.us-east-1.amazonaws.com"
)

Write-Host "Deploying Splunk Gateway Infrastructure..." -ForegroundColor Green
Write-Host "Stack Name: $StackName" -ForegroundColor Yellow
Write-Host "Region: $Region" -ForegroundColor Yellow
Write-Host "Environment: $Environment" -ForegroundColor Yellow
Write-Host "Allowed Origin: $AllowedOrigin" -ForegroundColor Yellow

# Clean up SplunkHecUrl - remove /services/collector path if present since Lambda adds it
$CleanSplunkHecUrl = $SplunkHecUrl -replace '/services/collector.*$', ''
if ($CleanSplunkHecUrl -ne $SplunkHecUrl) {
    Write-Host "Cleaned Splunk URL: $SplunkHecUrl -> $CleanSplunkHecUrl" -ForegroundColor Yellow
}
Write-Host "Splunk HEC URL: $CleanSplunkHecUrl" -ForegroundColor Yellow

# Check if AWS CLI is configured
try {
    aws sts get-caller-identity --output text --query 'Account' | Out-Null
    Write-Host "✓ AWS CLI configured" -ForegroundColor Green
} catch {
    Write-Error "AWS CLI not configured or not installed"
    exit 1
}

# Package Lambda function
Write-Host "Packaging Lambda function..." -ForegroundColor Yellow
$lambdaZipPath = Join-Path $PSScriptRoot "splunk-gateway-lambda.zip"

# Create temporary directory for Lambda package
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

try {
    # Copy Lambda files to temp directory
    Copy-Item (Join-Path $PSScriptRoot "lambda" "*") -Destination $tempDir -Recurse -Force
    
    # Create zip file
    if (Test-Path $lambdaZipPath) {
        Remove-Item $lambdaZipPath -Force
    }
    
    # Use PowerShell to create zip (works on both Windows and PowerShell Core)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $lambdaZipPath)
    
    Write-Host "✓ Lambda package created: $lambdaZipPath" -ForegroundColor Green
    
} finally {
    # Clean up temp directory
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

# Deploy CloudFormation stack
Write-Host "Deploying CloudFormation stack..." -ForegroundColor Yellow

$templatePath = Join-Path $PSScriptRoot "cloudformation-template.yaml"

try {
    # Deploy stack using parameter overrides in Key=Value format
    aws cloudformation deploy `
        --template-file $templatePath `
        --stack-name $StackName `
        --parameter-overrides `
            SplunkHecToken=$SplunkHecToken `
            SplunkHecUrl=$CleanSplunkHecUrl `
            Environment=$Environment `
            AllowedOrigin=$AllowedOrigin `
        --capabilities CAPABILITY_IAM CAPABILITY_NAMED_IAM `
        --region $Region
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Stack deployed successfully" -ForegroundColor Green
        
        # Get stack outputs
        $outputs = aws cloudformation describe-stacks `
            --stack-name $StackName `
            --region $Region `
            --query 'Stacks[0].Outputs' `
            --output json | ConvertFrom-Json
        
        Write-Host "`nStack Outputs:" -ForegroundColor Green
        foreach ($output in $outputs) {
            Write-Host "$($output.OutputKey): $($output.OutputValue)" -ForegroundColor Yellow
        }
        
        # Save outputs to file for easy reference
        $outputsFile = Join-Path $PSScriptRoot "deployment-outputs.json"
        $outputs | ConvertTo-Json -Depth 3 | Out-File -FilePath $outputsFile -Encoding UTF8
        Write-Host "`nOutputs saved to: $outputsFile" -ForegroundColor Green
        
    } else {
        Write-Error "Stack deployment failed"
        exit 1
    }
    
} catch {
    Write-Error "Stack deployment failed: $($_.Exception.Message)"
    exit 1
}

Write-Host "`n✓ Splunk Gateway deployment completed successfully!" -ForegroundColor Green
Write-Host "You can now update your application to use the API Gateway endpoints instead of direct Splunk HEC calls." -ForegroundColor Yellow
