# BedrockAPIKey Generator Script
# This script helps you create a properly formatted BedrockAPIKey for the Outlook add-in

Write-Host "BedrockAPIKey Generator" -ForegroundColor Green
Write-Host "======================" -ForegroundColor Green
Write-Host ""

# Get user input
Write-Host "Please enter your AWS credentials:" -ForegroundColor Yellow
$accessKeyId = Read-Host "Access Key ID (AKIA...)"
$secretAccessKey = Read-Host "Secret Access Key" -AsSecureString
$secretAccessKeyPlain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secretAccessKey))

# Optional session token
$sessionToken = Read-Host "Session Token (optional, press Enter to skip)"

Write-Host ""

# Create the credentials JSON
$credentialsObj = @{
    accessKeyId = $accessKeyId
    secretAccessKey = $secretAccessKeyPlain
}

if ($sessionToken -and $sessionToken.Trim() -ne "") {
    $credentialsObj.sessionToken = $sessionToken.Trim()
}

# Convert to JSON and base64 encode
$credentialsJson = $credentialsObj | ConvertTo-Json -Compress
$credentialsBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($credentialsJson))

# Generate a simple ID (timestamp-based)
$id = "$(Get-Date -Format 'yyyyMMdd-HHmmss')"

# Create the final BedrockAPIKey
$bedrockApiKey = "BedrockAPIKey-$id`:$credentialsBase64"

Write-Host "Generated BedrockAPIKey:" -ForegroundColor Green
Write-Host $bedrockApiKey -ForegroundColor Cyan
Write-Host ""
Write-Host "Copy this entire string and paste it into the API Key field for the bedrock1 provider in your Outlook add-in." -ForegroundColor Yellow
Write-Host ""

# Also save to clipboard if possible
try {
    $bedrockApiKey | Set-Clipboard
    Write-Host "✓ BedrockAPIKey has been copied to your clipboard!" -ForegroundColor Green
} catch {
    Write-Host "Note: Could not copy to clipboard automatically. Please copy the key manually." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Format explanation:" -ForegroundColor Magenta
Write-Host "BedrockAPIKey-[timestamp]:[base64-encoded-credentials-json]" -ForegroundColor White
Write-Host ""
Write-Host "The credentials JSON contains:" -ForegroundColor Magenta
Write-Host $credentialsJson -ForegroundColor White

# Clear sensitive variables
$secretAccessKeyPlain = $null
$secretAccessKey = $null