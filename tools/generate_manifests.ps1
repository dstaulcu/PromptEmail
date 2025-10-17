# Generate environment-specific manifests for PromptEmail Outlook Add-in
# This script generates manifests for each environment to show the unique IDs and names

param(
    [ValidateSet("Dev", "Test", "Prod", "All")]
    [string]$Environment = "All"
)

$ProjectRoot = Split-Path $PSScriptRoot -Parent
$ConfigPath = Join-Path $PSScriptRoot 'deployment-environments.json'
$TemplateePath = Join-Path $ProjectRoot 'src\manifest.template.xml'
$OutputDir = Join-Path $ProjectRoot 'manifests'

# Create output directory
if (-not (Test-Path $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path $ConfigPath)) {
    Write-Error "Configuration file not found: $ConfigPath"
    exit 1
}

if (-not (Test-Path $TemplateePath)) {
    Write-Error "Template file not found: $TemplateePath"
    exit 1
}

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$template = Get-Content $TemplateePath -Raw

$environments = if ($Environment -eq "All") { @("Dev", "Test", "Prod") } else { @($Environment) }

foreach ($env in $environments) {
    $envConfig = $config.environments.$env
    if (-not $envConfig) {
        Write-Warning "Configuration not found for environment: $env"
        continue
    }
    
    $baseUrl = "$($envConfig.publicUri.protocol)://$($envConfig.publicUri.host)"
    $hostDomain = $envConfig.publicUri.host
    
    # Replace placeholders
    $manifestContent = $template
    $manifestContent = $manifestContent.Replace('{{MANIFEST_ID}}', $envConfig.manifestId)
    $manifestContent = $manifestContent.Replace('{{DISPLAY_NAME}}', $envConfig.displayName)
    $manifestContent = $manifestContent.Replace('{{DESCRIPTION}}', $envConfig.description)
    $manifestContent = $manifestContent.Replace('{{GROUP_LABEL}}', $envConfig.groupLabel)
    $manifestContent = $manifestContent.Replace('{{BUTTON_LABEL}}', $envConfig.buttonLabel)
    $manifestContent = $manifestContent.Replace('{{BUTTON_TOOLTIP}}', $envConfig.buttonTooltip)
    $manifestContent = $manifestContent.Replace('{{BASE_URL}}', $baseUrl)
    $manifestContent = $manifestContent.Replace('{{HOST_DOMAIN}}', $hostDomain)
    
    # Write environment-specific manifest
    $outputFile = Join-Path $OutputDir "manifest-$($env.ToLower()).xml"
    $manifestContent | Set-Content $outputFile -Encoding UTF8
    
    Write-Host "Generated: $outputFile" -ForegroundColor Green
    Write-Host "  ID: $($envConfig.manifestId)" -ForegroundColor Yellow
    Write-Host "  Name: $($envConfig.displayName)" -ForegroundColor Yellow
    Write-Host "  URL: $baseUrl" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "Environment-specific manifests generated in: $OutputDir" -ForegroundColor Cyan
Write-Host ""
Write-Host "Usage Instructions:" -ForegroundColor Magenta
Write-Host "==================" -ForegroundColor Magenta
Write-Host "• Dev environment users: Use manifest-dev.xml for sideloading" -ForegroundColor White
Write-Host "• Test environment users: Use manifest-test.xml (deployed by admin)" -ForegroundColor White  
Write-Host "• Prod environment users: Use manifest-prod.xml (deployed by admin)" -ForegroundColor White
Write-Host ""
Write-Host "Each manifest has a unique ID, so all environments can coexist!" -ForegroundColor Green