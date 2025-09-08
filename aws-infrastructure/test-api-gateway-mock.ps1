param(
    [string]$TestMode = "mock"
)

# Test the Splunk API Gateway endpoint
# Replace with your actual API Gateway URL from the deployment output

$apiEndpoint = "https://your-api-id.execute-api.us-east-1.amazonaws.com/prod/telemetry"

# Test 1: Simple test event (using newline-delimited JSON format)
$testEvent1 = @{
    event = @{
        eventType = "api_gateway_test"
        message = "Testing API Gateway connection"
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        testNumber = 1
        testMode = $TestMode
    }
    sourcetype = "json:outlook_email_assistant"
    source = "api_gateway_test"
    index = "main"
}

# Convert to newline-delimited JSON (Splunk HEC format)
$testData1 = ($testEvent1 | ConvertTo-Json -Depth 3 -Compress)

Write-Host "Testing API Gateway endpoint: $apiEndpoint" -ForegroundColor Green
Write-Host "Test 1: Simple event (Mode: $TestMode)" -ForegroundColor Yellow

try {
    $response1 = Invoke-RestMethod -Uri $apiEndpoint -Method POST -Body $testData1 -ContentType "application/json" -Verbose
    Write-Host "✓ Test 1 SUCCESS:" -ForegroundColor Green
    Write-Host $response1 -ForegroundColor White
} catch {
    Write-Host "✗ Test 1 FAILED:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $responseBody = $reader.ReadToEnd()
        Write-Host "Response Body: $responseBody" -ForegroundColor Red
    }
}

Write-Host ""

# Test 2: Email analysis simulation (using newline-delimited JSON format)
$testEvent2 = @{
    event = @{
        eventType = "email_analysis_test"
        emailSubject = "TEST: API Gateway Integration"
        provider = "onsite1"
        analysisType = "test_simulation"
        duration = 1250
        success = $true
        timestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        sessionId = "test_" + [System.Guid]::NewGuid().ToString()
        testMode = $TestMode
    }
    sourcetype = "json:outlook_email_assistant"
    source = "outlook_addon"
    index = "main"
}

# Convert to newline-delimited JSON (Splunk HEC format)
$testData2 = ($testEvent2 | ConvertTo-Json -Depth 3 -Compress)

Write-Host "Test 2: Email analysis simulation (Mode: $TestMode)" -ForegroundColor Yellow

try {
    $response2 = Invoke-RestMethod -Uri $apiEndpoint -Method POST -Body $testData2 -ContentType "application/json"
    Write-Host "✓ Test 2 SUCCESS:" -ForegroundColor Green
    Write-Host $response2 -ForegroundColor White
} catch {
    Write-Host "✗ Test 2 FAILED:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""

# Test 3: CORS preflight test
Write-Host "Test 3: CORS preflight (OPTIONS request)" -ForegroundColor Yellow

try {
    $response3 = Invoke-WebRequest -Uri $apiEndpoint -Method OPTIONS -Verbose
    Write-Host "✓ Test 3 SUCCESS:" -ForegroundColor Green
    Write-Host "Status: $($response3.StatusCode)" -ForegroundColor White
    Write-Host "CORS Headers:" -ForegroundColor White
    $response3.Headers.GetEnumerator() | Where-Object { $_.Key -like "*Access-Control*" } | ForEach-Object { 
        Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Cyan 
    }
} catch {
    Write-Host "✗ Test 3 FAILED:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

Write-Host ""
Write-Host "Testing completed!" -ForegroundColor Green

if ($TestMode -eq "mock") {
    Write-Host "NOTE: Tests run in MOCK mode - no actual Splunk forwarding attempted." -ForegroundColor Yellow
    Write-Host "The 500 errors are expected because Splunk at 10.0.1.118:8088 is not accessible." -ForegroundColor Yellow
    Write-Host "If JSON parsing and API Gateway routing work, the infrastructure is correct." -ForegroundColor Green
} else {
    Write-Host "If all tests passed, your API Gateway is working correctly." -ForegroundColor Yellow
}
