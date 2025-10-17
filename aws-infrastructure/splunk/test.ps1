#!/usr/bin/env powershell
# Test script for Splunk HEC API Proxy
# Tests API connectivity, CORS, authentication, and event ingestion

param(
    [Parameter(Mandatory=$false)]
    [string]$ApiUrl = "",
    
    [Parameter(Mandatory=$false)]
    [SecureString]$HecToken,
    
    [Parameter(Mandatory=$false)]
    [switch]$VerboseOutput,
    
    [Parameter(Mandatory=$false)]
    [string]$TestIndex = "main",
    
    [Parameter(Mandatory=$false)]
    [string]$TestSource = "test-source",
    
    [Parameter(Mandatory=$false)]
    [string]$TestSourceType = "json"
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
    Write-Host "Example: .\test.ps1 -ApiUrl 'https://abc123.execute-api.us-east-1.amazonaws.com/dev/splunk' -HecToken (ConvertTo-SecureString 'your-hec-token' -AsPlainText -Force)"
    exit 1
}

Write-Host "Testing Splunk HEC API Proxy" -ForegroundColor Green
Write-Host "API URL: $ApiUrl" -ForegroundColor Yellow

# Test payload for Splunk HEC
$testEvent = @{
    time = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    index = $TestIndex
    source = $TestSource
    sourcetype = $TestSourceType
    event = @{
        message = "Test event from PowerShell script"
        level = "INFO"
        timestamp = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
        test_run = $true
    }
} | ConvertTo-Json -Depth 3

# Test 1: OPTIONS request for CORS
Write-Host "`nTest 1: CORS preflight (OPTIONS)" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Gray

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

# Test 2: POST without HEC token
Write-Host "`nTest 2: POST without HEC token (should fail)" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Gray

try {
    $response = Invoke-WebRequest -Uri $ApiUrl -Method POST `
        -ContentType "application/json" `
        -Headers @{
            "Origin" = "https://test-origin.com"
        } `
        -Body $testEvent `
        -UseBasicParsing
        
    Write-Host "Unexpected success without HEC token!" -ForegroundColor Red
    
} catch {
    $response = $_.Exception.Response
    if ($response -and $response.StatusCode -eq 401) {
        Write-Host "Correctly rejected request without HEC token (401)" -ForegroundColor Green
    } elseif ($response -and $response.StatusCode -eq 403) {
        Write-Host "Correctly rejected request without HEC token (403)" -ForegroundColor Green
    } else {
        Write-Host "Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Test 3: POST with HEC token (if provided)
if ($HecToken) {
    Write-Host "`nTest 3: POST with HEC token" -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Gray
    
    # Convert SecureString to plain text for HTTP header
    $hecTokenText = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($HecToken))
    
    try {
        $authResponse = Invoke-WebRequest -Uri $ApiUrl -Method POST `
            -ContentType "application/json" `
            -Headers @{
                "Authorization" = "Splunk $hecTokenText"
                "Origin" = "https://test-origin.com"
            } `
            -Body $testEvent `
            -UseBasicParsing
            
        Write-Host "Authenticated request Status: $($authResponse.StatusCode)" -ForegroundColor Green
        
        # Parse response
        try {
            $responseData = $authResponse.Content | ConvertFrom-Json
            Write-Host "Response received:" -ForegroundColor Green
            
            if ($responseData.text) {
                Write-Host "  Response: $($responseData.text)" -ForegroundColor White
            }
            
            if ($responseData.code) {
                Write-Host "  Code: $($responseData.code)" -ForegroundColor White
            }
            
            # Check for success indicators
            if ($authResponse.StatusCode -eq 200 -and $responseData.text -eq "Success") {
                Write-Host "Event successfully sent to Splunk!" -ForegroundColor Green
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
                
                # Try to read error response
                $errorStream = $_.Exception.Response.GetResponseStream()
                if ($errorStream) {
                    $reader = New-Object System.IO.StreamReader($errorStream)
                    $errorContent = $reader.ReadToEnd()
                    if ($errorContent) {
                        Write-Host "Error Response: $errorContent" -ForegroundColor Red
                    }
                }
            } catch {
                Write-Host "Could not read error details" -ForegroundColor Red
            }
        }
    }
} else {
    Write-Host "`nSkipping HEC token test (no token provided)" -ForegroundColor Cyan
    Write-Host "To test with HEC token, use:" -ForegroundColor White
    Write-Host ".\test.ps1 -ApiUrl '$ApiUrl' -HecToken (ConvertTo-SecureString 'your-hec-token' -AsPlainText -Force)" -ForegroundColor White
}

# Test 4: Invalid HEC token format
Write-Host "`nTest 4: Invalid HEC token format (should fail)" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Gray

try {
    $response = Invoke-WebRequest -Uri $ApiUrl -Method POST `
        -ContentType "application/json" `
        -Headers @{
            "Authorization" = "Splunk invalid-token-format"
            "Origin" = "https://test-origin.com"
        } `
        -Body $testEvent `
        -UseBasicParsing
        
    Write-Host "Unexpected success with invalid HEC token!" -ForegroundColor Red
    
} catch {
    $response = $_.Exception.Response
    if ($response -and ($response.StatusCode -eq 401 -or $response.StatusCode -eq 403)) {
        Write-Host "Correctly rejected invalid HEC token ($($response.StatusCode))" -ForegroundColor Green
    } else {
        Write-Host "Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Test 5: Test with malformed JSON (should fail)
Write-Host "`nTest 5: Malformed JSON payload (should fail)" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Gray

$malformedJson = '{"incomplete": "json"'

try {
    $response = Invoke-WebRequest -Uri $ApiUrl -Method POST `
        -ContentType "application/json" `
        -Headers @{
            "Authorization" = "Splunk test-token"
            "Origin" = "https://test-origin.com"
        } `
        -Body $malformedJson `
        -UseBasicParsing
        
    Write-Host "Unexpected success with malformed JSON!" -ForegroundColor Red
    
} catch {
    $response = $_.Exception.Response
    if ($response -and $response.StatusCode -eq 400) {
        Write-Host "Correctly rejected malformed JSON (400)" -ForegroundColor Green
    } else {
        Write-Host "Response: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Test 6: Test endpoint health/info
Write-Host "`nTest 6: Health check endpoint" -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Gray

$healthUrl = $ApiUrl + "/health"
try {
    $healthResponse = Invoke-WebRequest -Uri $healthUrl -Method GET -UseBasicParsing
    Write-Host "Health endpoint Status: $($healthResponse.StatusCode)" -ForegroundColor Green
    
    if ($VerboseOutput) {
        Write-Host "Health Response: $($healthResponse.Content)" -ForegroundColor White
    }
    
} catch {
    Write-Host "Health endpoint not available or failed: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`nTesting Complete!" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Gray
Write-Host "Monitoring:" -ForegroundColor Yellow
Write-Host "  CloudWatch Logs: /aws/lambda/{function-name}" -ForegroundColor White
Write-Host "  API Gateway Logs: Check API Gateway console" -ForegroundColor White
Write-Host "  Splunk Search: index=$TestIndex source=$TestSource" -ForegroundColor White

if ($VerboseOutput) {
    Write-Host "`nTest Event Sent:" -ForegroundColor Cyan
    Write-Host $testEvent -ForegroundColor White
}