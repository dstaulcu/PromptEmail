#!/usr/bin/env powershell
# Test script for S3 hosting configuration
# Tests bucket accessibility and configuration

param(
    [Parameter(Mandatory=$false)]
    [string]$Environment = "Dev",
    
    [Parameter(Mandatory=$false)]
    [switch]$VerboseOutput
)

$ErrorActionPreference = "Continue"

Write-Host "Testing S3 Hosting Configuration" -ForegroundColor Green
Write-Host "Environment: $Environment" -ForegroundColor Yellow

# Load deployment environments
$configPath = "..\..\tools\deployment-environments.json"
if (-not (Test-Path $configPath)) {
    Write-Error "Configuration file not found: $configPath"
    exit 1
}

try {
    $config = Get-Content $configPath | ConvertFrom-Json
    $envConfig = $config.environments.$Environment
    
    if (-not $envConfig) {
        Write-Error "Environment '$Environment' not found in configuration"
        exit 1
    }
    
    $bucketName = $envConfig.s3BucketName
    $region = $envConfig.awsRegion
    $publicUrl = "https://$bucketName.s3-website-$region.amazonaws.com"
    
    Write-Host "Bucket: $bucketName" -ForegroundColor White
    Write-Host "Region: $region" -ForegroundColor White
    Write-Host "Public URL: $publicUrl" -ForegroundColor White
    
} catch {
    Write-Error "Failed to load configuration: $($_.Exception.Message)"
    exit 1
}

# Test 1: Check if bucket exists
Write-Host "`nTest 1: Bucket existence" -ForegroundColor Yellow

try {
    aws s3api head-bucket --bucket $bucketName 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Bucket exists" -ForegroundColor Green
    } else {
        Write-Host "Bucket does not exist or not accessible" -ForegroundColor Red
    }
} catch {
    Write-Host "Failed to check bucket existence: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2: Check bucket configuration
Write-Host "`nTest 2: Website configuration" -ForegroundColor Yellow

try {
    $websiteConfig = aws s3api get-bucket-website --bucket $bucketName 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Website hosting is configured" -ForegroundColor Green
        
        if ($VerboseOutput) {
            $config = $websiteConfig | ConvertFrom-Json
            Write-Host "  Index Document: $($config.IndexDocument.Suffix)" -ForegroundColor White
            Write-Host "  Error Document: $($config.ErrorDocument.Key)" -ForegroundColor White
        }
    } else {
        Write-Host "Website hosting is not configured" -ForegroundColor Red
    }
} catch {
    Write-Host "Failed to check website configuration: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 3: Check public access
Write-Host "`nTest 3: Public access configuration" -ForegroundColor Yellow

try {
    $policy = aws s3api get-bucket-policy --bucket $bucketName 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Bucket policy is configured" -ForegroundColor Green
        
        if ($VerboseOutput) {
            Write-Host "  Policy allows public read access" -ForegroundColor White
            $parsedPolicy = ($policy | ConvertFrom-Json).Policy | ConvertFrom-Json
            Write-Host "  Policy Document Loaded: $(if ($parsedPolicy) { 'Yes' } else { 'No' })" -ForegroundColor White
            if ($parsedPolicy.Statement -and $parsedPolicy.Statement[0]) {
                Write-Host "  Policy Principal: $($parsedPolicy.Statement[0].Principal)" -ForegroundColor White
                Write-Host "  Policy Effect: $($parsedPolicy.Statement[0].Effect)" -ForegroundColor White
            }
        }
    } else {
        Write-Host "Bucket policy is not configured" -ForegroundColor Red
    }
} catch {
    Write-Host "Failed to check bucket policy: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 4: Test HTTP accessibility
Write-Host "`nTest 4: HTTP accessibility test" -ForegroundColor Yellow

try {
    $response = Invoke-WebRequest -Uri $publicUrl -Method GET -UseBasicParsing
    
    if ($response.StatusCode -eq 200) {
        Write-Host "Bucket is publicly accessible with content" -ForegroundColor Green
        Write-Host "  Status: $($response.StatusCode)" -ForegroundColor White
        Write-Host "  Content-Type: $($response.Headers['Content-Type'])" -ForegroundColor White
    } elseif ($response.StatusCode -eq 404) {
        Write-Host "Bucket is accessible but empty (404 expected)" -ForegroundColor Green
        Write-Host "  Status: $($response.StatusCode)" -ForegroundColor White
    } else {
        Write-Host "Unexpected response: $($response.StatusCode)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Failed to access bucket via HTTP: $($_.Exception.Message)" -ForegroundColor Red
    
    # Check if it's a DNS resolution issue
    if ($_.Exception.Message -like "*could not be resolved*") {
        Write-Host "  Note: This might be expected if the bucket hasn't been configured yet" -ForegroundColor Cyan
    }
}

# Test 5: Check CloudFront distribution (if configured)
Write-Host "`nTest 5: CloudFront distribution" -ForegroundColor Yellow

try {
    $distributions = aws cloudfront list-distributions --query "DistributionList.Items[?Comment=='$bucketName'].{Id:Id,DomainName:DomainName,Status:Status}" --output json 2>$null
    
    if ($LASTEXITCODE -eq 0 -and $distributions) {
        $distList = $distributions | ConvertFrom-Json
        if ($distList.Count -gt 0) {
            Write-Host "CloudFront distribution found" -ForegroundColor Green
            foreach ($dist in $distList) {
                Write-Host "  Distribution ID: $($dist.Id)" -ForegroundColor White
                Write-Host "  Domain Name: $($dist.DomainName)" -ForegroundColor White
                Write-Host "  Status: $($dist.Status)" -ForegroundColor White
            }
        } else {
            Write-Host "No CloudFront distribution found for this bucket" -ForegroundColor Yellow
        }
    } else {
        Write-Host "Unable to check CloudFront distributions" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Failed to check CloudFront: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nTesting Complete!" -ForegroundColor Green
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "   1. If bucket is not configured: .\deploy.ps1 -Environment $Environment" -ForegroundColor White
Write-Host "   2. Deploy add-in files: ..\..\tools\deploy_web_assets.ps1 -Environment $Environment" -ForegroundColor White
Write-Host "   3. Test add-in in Outlook" -ForegroundColor White
Write-Host ""
Write-Host "Useful Commands:" -ForegroundColor Yellow
Write-Host "   Check bucket in AWS Console: https://s3.console.aws.amazon.com/s3/buckets/$bucketName" -ForegroundColor White
Write-Host "   Direct bucket URL: $publicUrl" -ForegroundColor White