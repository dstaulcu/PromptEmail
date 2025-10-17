#!/usr/bin/env powershell
# Test script for Bedrock API Proxy

param(
    [Parameter(Mandatory=$false)]
    [string]$ApiUrl = "",
    
    [Parameter(Mandatory=$false)]
    [SecureString]$Credentials,
    
    [Parameter(Mandatory=$false)]
    [switch]$VerboseOutput,
    
    [Parameter(Mandatory=$false)]
    [string]$TestMessage = "Hello! Please respond with a brief greeting."
)

$ErrorActionPreference = "Continue"

# Load API URL from deployment info if not provided
if (-not $ApiUrl) {
    $deploymentInfoPath = Join-Path $PSScriptRoot "deployment-info.json"
    if (Test-Path $deploymentInfoPath) {
        try {
            $deploymentInfo = Get-Content $deploymentInfoPath | ConvertFrom-Json
            $ApiUrl = $deploymentInfo.apiUrl
            Write-Host "Loaded API URL from deployment info: $ApiUrl" -ForegroundColor Green
        } catch {
            Write-Warning "Could not load deployment info: $($_.Exception.Message)"
        }
    }
}

if (-not $ApiUrl) {
    Write-Error "API URL is required. Either provide -ApiUrl parameter or ensure deployment-info.json exists."
    Write-Host "Example: .\test.ps1 -ApiUrl 'https://abc123.execute-api.us-east-1.amazonaws.com/dev/bedrock'"
    exit 1
}

Write-Host "Testing Bedrock API Proxy" -ForegroundColor Green
Write-Host "API URL: $ApiUrl" -ForegroundColor Yellow

# Test payload
$testPayload = @{
    model = "anthropic.claude-3-haiku-20240307-v1:0"
    messages = @(@{
        role = "user"
        content = $TestMessage
    })
    max_tokens = 100
} | ConvertTo-Json -Depth 3

# Test 1: OPTIONS request for CORS
Write-Host "`nTest 1: CORS preflight (OPTIONS)" -ForegroundColor Yellow

try {
    $preflightResponse = Invoke-WebRequest -Uri $ApiUrl -Method OPTIONS `
        -Headers @{
            "Origin" = "https://test-origin.com"
            "Access-Control-Request-Method" = "POST"
            "Access-Control-Request-Headers" = "Content-Type,Authorization"
        } `
        -UseBasicParsing
        
    Write-Host "CORS preflight Status: $($preflightResponse.StatusCode)" -ForegroundColor Green
    
    # Check CORS headers
    $corsHeaders = @(
        "Access-Control-Allow-Origin",
        "Access-Control-Allow-Methods", 
        "Access-Control-Allow-Headers",
        "Access-Control-Max-Age"
    )
    
    foreach ($header in $corsHeaders) {
        $headerValue = $preflightResponse.Headers[$header]
        if ($headerValue) {
            Write-Host "  $header : $headerValue" -ForegroundColor Green
        } else {
            Write-Host "  Missing $header" -ForegroundColor Red
        }
    }
    
} catch {
    Write-Host "CORS preflight failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2: POST without credentials
Write-Host "`nTest 2: POST without credentials (should fail)" -ForegroundColor Yellow

try {
    $response = Invoke-WebRequest -Uri $ApiUrl -Method POST `
        -ContentType "application/json" `
        -Headers @{
            "Origin" = "https://test-origin.com"
        } `
        -Body $testPayload `
        -UseBasicParsing
        
    Write-Host "Unexpected success without credentials!" -ForegroundColor Red
    
} catch {
    $response = $_.Exception.Response
    if ($response -and $response.StatusCode -eq 401) {
        Write-Host "Correctly rejected request without credentials (401)" -ForegroundColor Green
    } else {
        Write-Host "Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Test 3: POST with credentials (if provided)
if ($Credentials) {
    Write-Host "`nTest 3: POST with credentials" -ForegroundColor Yellow
    
    # Convert SecureString to plain text for HTTP header
    $credentialsText = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credentials))
    
    try {
        $authResponse = Invoke-WebRequest -Uri $ApiUrl -Method POST `
            -ContentType "application/json" `
            -Headers @{
                "Authorization" = "Bearer $credentialsText"
                "Origin" = "https://test-origin.com"
            } `
            -Body $testPayload `
            -UseBasicParsing
            
        Write-Host "Authenticated request Status: $($authResponse.StatusCode)" -ForegroundColor Green
        
        # Parse response
        try {
            $responseData = $authResponse.Content | ConvertFrom-Json
            Write-Host "Response received:" -ForegroundColor Green
            $modelId = if ($responseData.modelId) { $responseData.modelId } else { 'Not specified' }
            Write-Host "  Model: $modelId" -ForegroundColor White
            
            if ($responseData.content) {
                $content = $responseData.content
                if ($content.Length -gt 200) {
                    $content = $content.Substring(0, 200) + "..."
                }
                Write-Host "  Content: $content" -ForegroundColor White
            }
            
        } catch {
            Write-Host "Raw response:" -ForegroundColor Green
            Write-Host "  $($authResponse.Content)" -ForegroundColor White
        }
        
    } catch {
        Write-Host "Authenticated request failed: $($_.Exception.Message)" -ForegroundColor Red
        
        if ($_.Exception.Response) {
            try {
                $statusCode = $_.Exception.Response.StatusCode
                Write-Host "Error Status Code: $statusCode" -ForegroundColor Red
            } catch {
                Write-Host "Could not read error details" -ForegroundColor Red
            }
        }
    }
} else {
    Write-Host "`nSkipping credentials test (no credentials provided)" -ForegroundColor Cyan
    Write-Host "To test with credentials, use:" -ForegroundColor White
    Write-Host ".\test.ps1 -ApiUrl '$ApiUrl' -Credentials (ConvertTo-SecureString 'AKIA...:secret...' -AsPlainText -Force)" -ForegroundColor White
}

# Test 4: Invalid credential format
Write-Host "`nTest 4: Invalid credential format (should fail)" -ForegroundColor Yellow

try {
    $response = Invoke-WebRequest -Uri $ApiUrl -Method POST `
        -ContentType "application/json" `
        -Headers @{
            "Authorization" = "Bearer invalid-credential-format"
            "Origin" = "https://test-origin.com"
        } `
        -Body $testPayload `
        -UseBasicParsing
        
    Write-Host "Unexpected success with invalid credentials!" -ForegroundColor Red
    
} catch {
    $response = $_.Exception.Response
    if ($response -and $response.StatusCode -eq 401) {
        Write-Host "Correctly rejected invalid credentials (401)" -ForegroundColor Green
    } else {
        Write-Host "Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nTesting Complete!" -ForegroundColor Green
Write-Host "Monitoring:" -ForegroundColor Yellow
Write-Host "  CloudWatch Logs: /aws/lambda/{function-name}" -ForegroundColor White
Write-Host "  API Gateway Logs: Check API Gateway console" -ForegroundColor White