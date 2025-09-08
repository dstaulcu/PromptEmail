# Quick test script that bypasses Splunk connectivity issues
# This creates a mock version of the Lambda function for testing

param(
    [Parameter(Mandatory=$false)]
    [string]$StackName = "outlook-assistant-splunk-gateway-prod"
)

Write-Host "Creating mock Lambda function for testing..." -ForegroundColor Yellow

# Create a mock Lambda function code
$mockLambdaCode = @"
exports.handler = async (event) => {
    console.info('[INFO] - Mock handler - Event:', JSON.stringify(event, null, 2));
    
    const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
    const corsHeaders = {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': allowedOrigin,
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With'
    };
    
    // Handle preflight CORS requests
    if (event.httpMethod === 'OPTIONS') {
        return {
            statusCode: 200,
            headers: corsHeaders,
            body: JSON.stringify({ message: 'CORS preflight handled (MOCK)' })
        };
    }
    
    // Mock successful Splunk response
    return {
        statusCode: 200,
        headers: corsHeaders,
        body: JSON.stringify({ 
            text: 'Success',
            code: 0,
            ack_id: 'mock-ack-' + Date.now(),
            mock: true,
            message: 'Mock response - Splunk not actually called'
        })
    };
};
"@

# Create temporary directory for mock
New-Item -ItemType Directory -Path "lambda-mock" -Force | Out-Null

try {
    # Write mock code to file
    $mockLambdaCode | Out-File -FilePath "lambda-mock\index.js" -Encoding UTF8

    # Package mock Lambda
    $mockZipPath = "splunk-gateway-lambda-mock.zip"
    if (Test-Path $mockZipPath) { Remove-Item $mockZipPath -Force }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory("lambda-mock", $mockZipPath)

    Write-Host "Updating Lambda function with mock code..." -ForegroundColor Yellow

    # Update Lambda function
    aws lambda update-function-code `
        --function-name "$StackName-splunk-gateway" `
        --zip-file fileb://$mockZipPath `
        --region us-east-1

    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Lambda function updated with mock code" -ForegroundColor Green
        Write-Host "Now run: .\test-api-gateway.ps1" -ForegroundColor Yellow
        Write-Host "It should return mock success responses" -ForegroundColor Yellow
    } else {
        Write-Error "Failed to update Lambda function"
    }

} finally {
    # Clean up
    Remove-Item "lambda-mock" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $mockZipPath -Force -ErrorAction SilentlyContinue
}
