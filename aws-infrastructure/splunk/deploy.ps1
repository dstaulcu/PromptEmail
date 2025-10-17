#!/usr/bin/env powershell
# Deploy script for Splunk HEC API Proxy
# Creates Lambda function and API Gateway for Splunk HTTP Event Collector proxy

param(
    [Parameter(Mandatory=$false)]
    [string]$Environment = "Dev",
    
    [Parameter(Mandatory=$false)]
    [string]$SplunkHecUrl = "",
    
    [Parameter(Mandatory=$false)]
    [string]$AllowedOrigin = "*",
    
    [Parameter(Mandatory=$false)]
    [string]$SplunkIndex = "",
    
    [Parameter(Mandatory=$false)]
    [string]$SplunkHost = "",
    
    [Parameter(Mandatory=$false)]
    [string]$SplunkSource = "",
    
    [Parameter(Mandatory=$false)]
    [string]$SplunkSourcetype = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Help
)

$ErrorActionPreference = "Stop"

# Show help information
if ($Help) {
    Write-Host "Splunk HEC API Proxy Deployment Script" -ForegroundColor Green
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "This script deploys a Lambda function and API Gateway for Splunk HEC proxy."
    Write-Host ""
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host "  .\deploy.ps1 -Environment <name> -SplunkHecUrl <url>     Deploy to specific environment"
    Write-Host "  .\deploy.ps1 -Help                                       Show this help"
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -Environment      Target environment (Dev, Test, Prod)"
    Write-Host "  -SplunkHecUrl     Splunk HTTP Event Collector URL (required)"
    Write-Host "  -AllowedOrigin    CORS allowed origin (default: *)"
    Write-Host "  -SplunkIndex      Default Splunk index for events (optional)"
    Write-Host "  -SplunkHost       Default Splunk host for events (optional)"
    Write-Host "  -SplunkSource     Default Splunk source for events (optional)"
    Write-Host "  -SplunkSourcetype Default Splunk sourcetype for events (optional)"
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  .\deploy.ps1 -Environment Dev -SplunkHecUrl 'https://splunk.company.com:8088/services/collector/event'"
    Write-Host "  .\deploy.ps1 -Environment Prod -SplunkHecUrl 'https://splunk.company.com:8088/services/collector/event' -AllowedOrigin 'https://app.company.com'"
    Write-Host "  .\deploy.ps1 -Environment Prod -SplunkHecUrl 'https://splunk.company.com:8088/services/collector/event' -SplunkIndex 'myapp' -SplunkHost 'webserver' -SplunkSource 'outlook_addon' -SplunkSourcetype 'json'"
    Write-Host ""
    Write-Host "REQUIREMENTS:" -ForegroundColor Yellow
    Write-Host "  - AWS CLI configured with appropriate permissions"
    Write-Host "  - Lambda execution role: lambda-execution-role"
    Write-Host "  - deployment-environments.json configuration file"
    Write-Host "  - Splunk HEC endpoint accessible from AWS"
    Write-Host ""
    Write-Host "OUTPUTS:" -ForegroundColor Yellow
    Write-Host "  - Lambda function: splunk-hec-proxy-{Environment}"
    Write-Host "  - API Gateway: splunk-hec-api-{Environment}"
    Write-Host "  - API URL: https://{api-id}.execute-api.{region}.amazonaws.com/{stage}/splunk"
    Write-Host "  - deployment-info.json with deployment details"
    exit 0
}

Write-Host "Deploying Splunk HEC API Proxy" -ForegroundColor Green
Write-Host "Environment: $Environment" -ForegroundColor Yellow

# Load configuration
# Get the script's directory and construct path to tools directory
$scriptDir = $PSScriptRoot
$rootDir = Split-Path (Split-Path $scriptDir -Parent) -Parent
$configPath = Join-Path $rootDir "tools\deployment-environments.json"

Write-Host "Looking for configuration file at: $configPath" -ForegroundColor Cyan

if (-not (Test-Path $configPath)) {
    Write-Error "Configuration file not found: $configPath"
    Write-Host "Script directory: $scriptDir" -ForegroundColor Yellow
    Write-Host "Root directory: $rootDir" -ForegroundColor Yellow
    Write-Host "Expected config path: $configPath" -ForegroundColor Yellow
    exit 1
}

try {
    $config = Get-Content $configPath | ConvertFrom-Json
    $envConfig = $config.environments.$Environment
    
    if (-not $envConfig) {
        Write-Error "Environment '$Environment' not found in configuration"
        exit 1
    }
    
    $Region = $envConfig.region
    $FunctionName = "splunk-hec-proxy-$Environment"
    $Stage = $Environment.ToLower()
    
    # Validate required configuration properties
    if (-not $Region) {
        Write-Host "Error: Environment '$Environment' is missing required 'region' property" -ForegroundColor Red
        Write-Host ""
        
        # Show help automatically
        Write-Host "Splunk HEC API Proxy Deployment Script" -ForegroundColor Green
        Write-Host "=======================================" -ForegroundColor Green
        Write-Host ""
        Write-Host "This script deploys a Lambda function and API Gateway for Splunk HEC proxy."
        Write-Host ""
        Write-Host "USAGE:" -ForegroundColor Yellow
        Write-Host "  .\deploy.ps1 -Environment <name> -SplunkHecUrl <url>     Deploy to specific environment"
        Write-Host "  .\deploy.ps1 -Help                                       Show this help"
        Write-Host ""
        Write-Host "PARAMETERS:" -ForegroundColor Yellow
        Write-Host "  -Environment      Target environment (Dev, Test, Prod)"
        Write-Host "  -SplunkHecUrl     Splunk HTTP Event Collector URL (required)"
        Write-Host "  -AllowedOrigin    CORS allowed origin (default: *)"
        Write-Host ""
        Write-Host "EXAMPLES:" -ForegroundColor Yellow
        Write-Host "  .\deploy.ps1 -Environment Dev -SplunkHecUrl 'https://splunk.company.com:8088/services/collector/event'"
        Write-Host ""
        Write-Host "CONFIGURATION:" -ForegroundColor Yellow
        Write-Host "  The deployment-environments.json file must contain:"
        Write-Host "  {" -ForegroundColor Gray
        Write-Host "    \"environments\": {" -ForegroundColor Gray
        Write-Host "      \"$Environment\": {" -ForegroundColor Gray
        Write-Host "        \"region\": \"us-east-1\"" -ForegroundColor Gray
        Write-Host "      }" -ForegroundColor Gray
        Write-Host "    }" -ForegroundColor Gray
        Write-Host "  }" -ForegroundColor Gray
        exit 1
    }
    
    Write-Host "Region: $Region" -ForegroundColor White
    Write-Host "Function: $FunctionName" -ForegroundColor White
    Write-Host "Stage: $Stage" -ForegroundColor White
    
} catch {
    Write-Error "Failed to load configuration: $($_.Exception.Message)"
    exit 1
}

# Validate required parameters
if (-not $SplunkHecUrl) {
    Write-Host "Error: SplunkHecUrl parameter is required" -ForegroundColor Red
    Write-Host ""
    
    # Show help automatically
    Write-Host "Splunk HEC API Proxy Deployment Script" -ForegroundColor Green
    Write-Host "=======================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "This script deploys a Lambda function and API Gateway for Splunk HEC proxy."
    Write-Host ""
    Write-Host "USAGE:" -ForegroundColor Yellow
    Write-Host "  .\deploy.ps1 -Environment <name> -SplunkHecUrl <url>     Deploy to specific environment"
    Write-Host "  .\deploy.ps1 -Help                                       Show this help"
    Write-Host ""
    Write-Host "PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -Environment      Target environment (Dev, Test, Prod)"
    Write-Host "  -SplunkHecUrl     Splunk HTTP Event Collector URL (required)"
    Write-Host "  -AllowedOrigin    CORS allowed origin (default: *)"
    Write-Host ""
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "  .\deploy.ps1 -Environment Dev -SplunkHecUrl 'https://splunk.company.com:8088/services/collector/event'"
    Write-Host "  .\deploy.ps1 -Environment Prod -SplunkHecUrl 'https://splunk.company.com:8088/services/collector/event' -AllowedOrigin 'https://app.company.com'"
    Write-Host ""
    Write-Host "REQUIREMENTS:" -ForegroundColor Yellow
    Write-Host "  - AWS CLI configured with appropriate permissions"
    Write-Host "  - Lambda execution role: lambda-execution-role"
    Write-Host "  - deployment-environments.json configuration file"
    Write-Host "  - Splunk HEC endpoint accessible from AWS"
    Write-Host ""
    Write-Host "OUTPUTS:" -ForegroundColor Yellow
    Write-Host "  - Lambda function: splunk-hec-proxy-{Environment}"
    Write-Host "  - API Gateway: splunk-hec-api-{Environment}"
    Write-Host "  - API URL: https://{api-id}.execute-api.{region}.amazonaws.com/{stage}/splunk"
    Write-Host "  - deployment-info.json with deployment details"
    exit 1
}

# Get AWS account ID
$awsAccount = aws sts get-caller-identity --query 'Account' --output text
if (-not $awsAccount) {
    Write-Error "Failed to get AWS account ID. Please ensure AWS CLI is configured."
    exit 1
}

Write-Host "AWS Account: $awsAccount" -ForegroundColor White

# Step 1: Package Lambda function
Write-Host "`nPackaging Lambda function..." -ForegroundColor Yellow

$scriptDir = $PSScriptRoot
$lambdaDir = Join-Path $scriptDir "lambda"
$zipPath = Join-Path $scriptDir "splunk-hec-proxy.zip"

if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

# Create zip file with Lambda code
try {
    Push-Location $lambdaDir
    Compress-Archive -Path "*.js", "*.json" -DestinationPath $zipPath -Force
    Pop-Location
    Write-Host "Lambda package created: $zipPath" -ForegroundColor Green
} catch {
    Pop-Location
    Write-Error "Failed to create Lambda package: $($_.Exception.Message)"
    exit 1
}

# Step 2: Create or update Lambda function
Write-Host "`nDeploying Lambda function..." -ForegroundColor Yellow

# Build environment variables string
$envVars = @{
    'SPLUNK_HEC_URL' = $SplunkHecUrl
    'ALLOWED_ORIGIN' = $AllowedOrigin
}

# Add optional Splunk metadata parameters if provided
if ($SplunkIndex) { $envVars['SPLUNK_INDEX'] = $SplunkIndex }
if ($SplunkHost) { $envVars['SPLUNK_HOST'] = $SplunkHost }
if ($SplunkSource) { $envVars['SPLUNK_SOURCE'] = $SplunkSource }
if ($SplunkSourcetype) { $envVars['SPLUNK_SOURCETYPE'] = $SplunkSourcetype }

# Convert to the format AWS CLI expects: Key1=Value1,Key2=Value2
$envVarString = ($envVars.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ","

$functionExists = $false
try {
    aws lambda get-function --function-name $FunctionName --region $Region --output text --query 'Configuration.FunctionName' | Out-Null
    $functionExists = $true
    Write-Host "Lambda function exists, will update" -ForegroundColor Green
} catch {
    Write-Host "Lambda function doesn't exist, will create" -ForegroundColor Cyan
}

if ($functionExists) {
    # Update existing function code
    aws lambda update-function-code `
        --function-name $FunctionName `
        --zip-file "fileb://$zipPath" `
        --region $Region | Out-Null
    
    Write-Host "Updated Lambda function code" -ForegroundColor Green
    
    # Update environment variables
    aws lambda update-function-configuration `
        --function-name $FunctionName `
        --environment "Variables={$envVarString}" `
        --region $Region | Out-Null
        
    Write-Host "Updated Lambda configuration" -ForegroundColor Green
    
} else {
    # Create new function
    aws lambda create-function `
        --function-name $FunctionName `
        --runtime nodejs18.x `
        --role "arn:aws:iam::${awsAccount}:role/lambda-execution-role" `
        --handler index.handler `
        --zip-file "fileb://$zipPath" `
        --timeout 30 `
        --memory-size 128 `
        --environment "Variables={$envVarString}" `
        --region $Region | Out-Null
    
    Write-Host "Created Lambda function" -ForegroundColor Green
}

# Step 3: Create or get API Gateway
Write-Host "`nSetting up API Gateway..." -ForegroundColor Yellow

$apiName = "splunk-hec-api-$Environment"

try {
    $existingApis = aws apigateway get-rest-apis --region $Region --output json | ConvertFrom-Json
    $existingApi = $existingApis.items | Where-Object { $_.name -eq $apiName }
    
    if ($existingApi) {
        $apiId = $existingApi.id
        Write-Host "Using existing API Gateway: $apiId" -ForegroundColor Green
    } else {
        throw "API not found"
    }
} catch {
    # Create new API Gateway
    $apiResult = aws apigateway create-rest-api `
        --name $apiName `
        --description "API Gateway for $FunctionName" `
        --endpoint-configuration types=REGIONAL `
        --region $Region `
        --output json | ConvertFrom-Json
    $apiId = $apiResult.id
    Write-Host "Created API Gateway: $apiId" -ForegroundColor Green
}

# Get root resource ID
$rootResourceId = aws apigateway get-resources --rest-api-id $apiId --region $Region --output text --query 'items[?path==`/`].id'

# Step 4: Create /splunk resource (if it doesn't exist)
Write-Host "`nConfiguring API resources..." -ForegroundColor Yellow

try {
    $splunkResourceId = aws apigateway get-resources --rest-api-id $apiId --region $Region --output text --query 'items[?pathPart==`splunk`].id'
    
    if (-not $splunkResourceId) {
        $resourceResult = aws apigateway create-resource `
            --rest-api-id $apiId `
            --parent-id $rootResourceId `
            --path-part "splunk" `
            --region $Region `
            --output json | ConvertFrom-Json
        $splunkResourceId = $resourceResult.id
        Write-Host "Created /splunk resource: $splunkResourceId" -ForegroundColor Green
    } else {
        Write-Host "Using existing /splunk resource: $splunkResourceId" -ForegroundColor Green
    }
} catch {
    Write-Error "Failed to create /splunk resource: $($_.Exception.Message)"
    exit 1
}

# Step 5: Create OPTIONS method for CORS
Write-Host "`nConfiguring CORS..." -ForegroundColor Yellow

try {
    aws apigateway put-method `
        --rest-api-id $apiId `
        --resource-id $splunkResourceId `
        --http-method OPTIONS `
        --authorization-type NONE `
        --region $Region | Out-Null
        
    # Set up OPTIONS integration
    aws apigateway put-integration `
        --rest-api-id $apiId `
        --resource-id $splunkResourceId `
        --http-method OPTIONS `
        --type MOCK `
        --integration-http-method OPTIONS `
        --request-templates '{"application/json":"{\"statusCode\":200}"}' `
        --region $Region | Out-Null
        
    # Set up OPTIONS method response
    aws apigateway put-method-response `
        --rest-api-id $apiId `
        --resource-id $splunkResourceId `
        --http-method OPTIONS `
        --status-code 200 `
        --response-parameters method.response.header.Access-Control-Allow-Headers=false,method.response.header.Access-Control-Allow-Methods=false,method.response.header.Access-Control-Allow-Origin=false `
        --region $Region | Out-Null
        
    # Set up OPTIONS integration response
    aws apigateway put-integration-response `
        --rest-api-id $apiId `
        --resource-id $splunkResourceId `
        --http-method OPTIONS `
        --status-code 200 `
        --response-parameters method.response.header.Access-Control-Allow-Headers="'Content-Type,X-Amz-Date,Authorization,X-Api-Key,X-Amz-Security-Token'",method.response.header.Access-Control-Allow-Methods="'POST,OPTIONS'",method.response.header.Access-Control-Allow-Origin="'$AllowedOrigin'" `
        --region $Region | Out-Null
        
    Write-Host "Created OPTIONS method for CORS" -ForegroundColor Green
} catch {
    Write-Host "OPTIONS method might already exist" -ForegroundColor Cyan
}

# Step 6: Create POST method
Write-Host "`nConfiguring POST method..." -ForegroundColor Yellow

try {
    aws apigateway put-method `
        --rest-api-id $apiId `
        --resource-id $splunkResourceId `
        --http-method POST `
        --authorization-type NONE `
        --region $Region | Out-Null
        
    # Set up POST integration
    aws apigateway put-integration `
        --rest-api-id $apiId `
        --resource-id $splunkResourceId `
        --http-method POST `
        --type AWS_PROXY `
        --integration-http-method POST `
        --uri "arn:aws:apigateway:${Region}:lambda:path/2015-03-31/functions/arn:aws:lambda:${Region}:${awsAccount}:function:${FunctionName}/invocations" `
        --region $Region | Out-Null
        
    Write-Host "Created POST method" -ForegroundColor Green
} catch {
    Write-Host "POST method might already exist" -ForegroundColor Cyan
}

# Step 7: Grant API Gateway permission to invoke Lambda
Write-Host "`nSetting up Lambda permissions..." -ForegroundColor Yellow

$statementId = "apigateway-invoke-permission"
try {
    aws lambda add-permission `
        --function-name $FunctionName `
        --statement-id $statementId `
        --action lambda:InvokeFunction `
        --principal apigateway.amazonaws.com `
        --source-arn "arn:aws:execute-api:${Region}:${awsAccount}:${apiId}/*/*" `
        --region $Region | Out-Null
    Write-Host "Added Lambda invoke permission" -ForegroundColor Green
} catch {
    Write-Host "Lambda permission might already exist" -ForegroundColor Cyan
}

# Step 8: Deploy API Gateway
Write-Host "`nDeploying API Gateway stage..." -ForegroundColor Yellow

aws apigateway create-deployment `
    --rest-api-id $apiId `
    --stage-name $Stage `
    --stage-description "Deployed via AWS CLI on $(Get-Date)" `
    --region $Region | Out-Null

Write-Host "Deployed API Gateway stage: $Stage" -ForegroundColor Green

# Step 9: Get API URL
$apiUrl = "https://$apiId.execute-api.$Region.amazonaws.com/$Stage/splunk"

Write-Host "`nDeployment Complete!" -ForegroundColor Green
Write-Host "API URL: $apiUrl" -ForegroundColor Cyan
Write-Host "Splunk HEC URL: $SplunkHecUrl" -ForegroundColor White
Write-Host "Allowed Origin: $AllowedOrigin" -ForegroundColor White

# Save deployment info
$deploymentInfo = @{
    functionName = $FunctionName
    apiId = $apiId
    apiUrl = $apiUrl
    splunkHecUrl = $SplunkHecUrl
    region = $Region
    stage = $Stage
    allowedOrigin = $AllowedOrigin
} | ConvertTo-Json -Depth 3

$deploymentInfoPath = Join-Path $scriptDir "deployment-info.json"
$deploymentInfo | Out-File -FilePath $deploymentInfoPath -Encoding UTF8
Write-Host "Deployment info saved to: $deploymentInfoPath" -ForegroundColor Cyan

Write-Host "`nNext Steps:" -ForegroundColor Yellow
Write-Host "1. Test the deployment: .\test.ps1 -ApiUrl '$apiUrl'" -ForegroundColor White
Write-Host "2. Configure your application to use the API URL" -ForegroundColor White
Write-Host "3. Monitor CloudWatch Logs: /aws/lambda/$FunctionName" -ForegroundColor White

# Clean up
Remove-Item $zipPath -Force
Write-Host "`nCleanup completed" -ForegroundColor Green