# Environment Configuration Helper for PromptEmail
# Helps switch between home and work environments securely

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("home", "work", "status", "setup")]
    [string]$Action = "status",
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

$ErrorActionPreference = "Stop"

# Paths
$toolsDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $toolsDir
$configDir = Join-Path $projectRoot "src\config"

$baseDeploymentConfig = Join-Path $toolsDir "deployment-config.json"
$workDeploymentConfig = Join-Path $toolsDir "deployment-config.work.json"
$localDeploymentConfig = Join-Path $toolsDir "deployment-config.local.json"

$baseAiProviders = Join-Path $configDir "ai-providers.json"
$workAiProviders = Join-Path $configDir "ai-providers.work.json"
$workAiProvidersTemplate = Join-Path $configDir "ai-providers.work.template.json"

function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

function Get-CurrentEnvironment {
    if (Test-Path $localDeploymentConfig) {
        $config = Get-Content $localDeploymentConfig | ConvertFrom-Json
        return $config.environment
    } elseif (Test-Path $workDeploymentConfig) {
        $config = Get-Content $workDeploymentConfig | ConvertFrom-Json
        return $config.work.environment
    } else {
        $config = Get-Content $baseDeploymentConfig | ConvertFrom-Json
        return $config.default.environment
    }
}

function Show-Status {
    Write-ColorOutput "`n🔍 ENVIRONMENT STATUS" "Cyan"
    Write-ColorOutput "=====================" "Cyan"
    
    $currentEnv = Get-CurrentEnvironment
    Write-ColorOutput "Current Environment: $currentEnv" "Green"
    
    Write-ColorOutput "`n📁 Configuration Files:" "Yellow"
    
    @(
        @{ Path = $baseDeploymentConfig; Label = "Base Config"; Required = $true }
        @{ Path = $workDeploymentConfig; Label = "Work Config (git-ignored)"; Required = $false }
        @{ Path = $localDeploymentConfig; Label = "Local Override (git-ignored)"; Required = $false }
        @{ Path = $baseAiProviders; Label = "Base AI Providers"; Required = $true }
        @{ Path = $workAiProviders; Label = "Work AI Providers (git-ignored)"; Required = $false }
    ) | ForEach-Object {
        $exists = Test-Path $_.Path
        $status = if ($exists) { "✅ EXISTS" } else { if ($_.Required) { "❌ MISSING" } else { "⚪ Optional" } }
        $color = if ($exists) { "Green" } else { if ($_.Required) { "Red" } else { "Gray" } }
        Write-ColorOutput "  $($_.Label): $status" $color
    }
    
    Write-ColorOutput "`n🚨 Security Check:" "Yellow"
    $gitIgnore = Join-Path $projectRoot ".gitignore"
    if (Test-Path $gitIgnore) {
        $gitIgnoreContent = Get-Content $gitIgnore -Raw
        $hasWorkPattern = $gitIgnoreContent -match '\*\.work\.json'
        $hasLocalPattern = $gitIgnoreContent -match '\*\.local\.json'
        
        if ($hasWorkPattern -and $hasLocalPattern) {
            Write-ColorOutput "  ✅ .gitignore properly configured for security" "Green"
        } else {
            Write-ColorOutput "  ⚠️  .gitignore may need security patterns" "Yellow"
        }
    }
}

function Set-Environment {
    param([string]$Environment)
    
    Write-ColorOutput "`n🔄 SWITCHING TO $($Environment.ToUpper()) ENVIRONMENT" "Cyan"
    Write-ColorOutput "======================================" "Cyan"
    
    switch ($Environment) {
        "home" {
            # Remove work/local configs to fall back to base
            if (Test-Path $localDeploymentConfig) {
                if ($Force -or (Read-Host "Remove local config override? (y/N)") -eq 'y') {
                    Remove-Item $localDeploymentConfig
                    Write-ColorOutput "✅ Removed local config override" "Green"
                }
            }
            Write-ColorOutput "✅ Switched to HOME environment" "Green"
        }
        
        "work" {
            if (-not (Test-Path $workDeploymentConfig)) {
                Write-ColorOutput "❌ Work config not found. Run 'setup' first." "Red"
                return
            }
            
            if (-not (Test-Path $workAiProviders)) {
                Write-ColorOutput "❌ Work AI providers config not found. Run 'setup' first." "Red"
                return
            }
            
            # Create local override pointing to work config
            $localConfig = @{
                environment = "work"
                aiProvidersFile = "ai-providers.work.json"
                description = "Local override pointing to work environment"
            }
            
            $localConfig | ConvertTo-Json -Depth 3 | Out-File $localDeploymentConfig -Encoding UTF8
            Write-ColorOutput "✅ Switched to WORK environment" "Green"
        }
    }
    
    Show-Status
}

function Initialize-WorkEnvironment {
    Write-ColorOutput "`n🛠️  WORK ENVIRONMENT SETUP" "Cyan"
    Write-ColorOutput "===========================" "Cyan"
    
    # Check if work configs already exist
    if ((Test-Path $workDeploymentConfig) -and (Test-Path $workAiProviders) -and -not $Force) {
        Write-ColorOutput "⚠️  Work configs already exist. Use -Force to overwrite." "Yellow"
        return
    }
    
    Write-ColorOutput "`n📋 This will create templates for:" "Yellow"
    Write-ColorOutput "  • deployment-config.work.json (deployment settings)" "White"
    Write-ColorOutput "  • ai-providers.work.json (AI provider configs)" "White"
    Write-ColorOutput "`n🚨 SECURITY REMINDER:" "Red"
    Write-ColorOutput "  • These files are git-ignored and contain classified data" "Red"
    Write-ColorOutput "  • Never commit them to any repository" "Red"
    Write-ColorOutput "  • Only edit them on secure work systems" "Red"
    
    if (-not $Force -and (Read-Host "`nContinue with setup? (y/N)") -ne 'y') {
        Write-ColorOutput "❌ Setup cancelled" "Yellow"
        return
    }
    
    # Copy templates
    if (Test-Path $workAiProvidersTemplate) {
        Copy-Item $workAiProvidersTemplate $workAiProviders
        Write-ColorOutput "✅ Created ai-providers.work.json from template" "Green"
    }
    
    # Create work deployment config from template in memory
    $workDeploymentTemplate = Join-Path $toolsDir "deployment-config.work.template.json"
    if (Test-Path $workDeploymentTemplate) {
        Copy-Item $workDeploymentTemplate $workDeploymentConfig
        Write-ColorOutput "✅ Created deployment-config.work.json from template" "Green"
    }
    
    Write-ColorOutput "`n📝 NEXT STEPS:" "Cyan"
    Write-ColorOutput "1. Edit deployment-config.work.json with your work S3 bucket and settings" "White"
    Write-ColorOutput "2. Edit ai-providers.work.json with your work AI providers and endpoints" "White"
    Write-ColorOutput "3. Run: .\switch-environment.ps1 work" "White"
    Write-ColorOutput "4. Test your deployment with work-specific values" "White"
}

# Main execution
try {
    switch ($Action) {
        "status" { Show-Status }
        "home" { Set-Environment "home" }
        "work" { Set-Environment "work" }
        "setup" { Initialize-WorkEnvironment }
    }
} catch {
    Write-ColorOutput "`n❌ ERROR: $($_.Exception.Message)" "Red"
    exit 1
}