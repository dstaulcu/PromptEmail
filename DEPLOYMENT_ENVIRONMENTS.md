# Multi-Environment Deployment Solution

## Problem Solved

Previously, all environments (Dev, Test, Prod) used the same manifest ID, causing conflicts where users could only install one version of the add-in. This prevented developers from testing dev versions while test/prod versions were deployed by Exchange admins.

## Solution

Each environment now has a **unique manifest ID** and **distinct display name**, allowing all environments to coexist on the same user's Outlook installation.

## Environment Configuration

| Environment | Manifest ID | Display Name | Use Case |
|------------|-------------|--------------|----------|
| **Dev** | `...789DEV` | "Prompt Email (Dev)" | Developer sideloading & testing |
| **Test** | `...789TST` | "Prompt Email (Test)" | Admin-deployed testing |
| **Prod** | `...789PRD` | "Prompt Email" | Admin-deployed production |

## How It Works

### 1. Template-Based Generation
- `src/manifest.template.xml` - Single template with placeholders
- `tools/deployment-environments.json` - Environment-specific configuration
- Deployment script generates environment-specific manifests automatically

### 2. Automatic Deployment
```powershell
# Deploys with environment-specific manifest
.\tools\deploy_web_assets.ps1 -Environment Dev
.\tools\deploy_web_assets.ps1 -Environment Test  
.\tools\deploy_web_assets.ps1 -Environment Prod
```

### 3. Manual Manifest Generation
```powershell
# Generate all environment manifests for review
.\tools\generate_manifests.ps1

# Generate specific environment manifest
.\tools\generate_manifests.ps1 -Environment Dev
```

## Developer Workflow

### Before (Conflicted)
1. ❌ Exchange admin deploys Test manifest → blocks Dev manifest
2. ❌ Developer can't sideload Dev version (same ID conflict)
3. ❌ Must disable Test to use Dev version

### After (Coexistent) 
1. ✅ Exchange admin deploys Test manifest (`...789TST`)
2. ✅ Developer sideloads Dev manifest (`...789DEV`) 
3. ✅ Both versions coexist peacefully in Outlook
4. ✅ User sees both "Prompt Email (Test)" and "Prompt Email (Dev)"

## Exchange Admin Instructions

### For Test Environment
1. Run: `.\tools\deploy_web_assets.ps1 -Environment Test`
2. Deploy generated `public/manifest.xml` to Exchange
3. Manifest will show as "Prompt Email (Test)" with test URLs

### For Production Environment  
1. Run: `.\tools\deploy_web_assets.ps1 -Environment Prod`
2. Deploy generated `public/manifest.xml` to Exchange
3. Manifest will show as "Prompt Email" with production URLs

## Developer Instructions

### For Development Testing
1. Ensure Test version is deployed by admin (doesn't conflict)
2. Run: `.\tools\deploy_web_assets.ps1 -Environment Dev`  
3. Sideload `public/manifest.xml` in Outlook
4. You'll see both "Prompt Email (Test)" and "Prompt Email (Dev)"

## Benefits

- 🎯 **No Conflicts**: Each environment has unique ID
- 🔄 **Parallel Testing**: Dev and Test can run simultaneously  
- 👥 **User Choice**: Users can enable/disable environments independently
- 🚀 **Seamless Development**: Developers can test without affecting production users
- 📊 **Clear Identification**: Environment-specific names prevent confusion

## File Structure

```
src/
  manifest.template.xml        # Template with {{PLACEHOLDERS}}
  manifest.xml                 # Legacy file (still supported)
tools/
  deployment-environments.json # Environment configurations
  deploy_web_assets.ps1       # Main deployment script  
  generate_manifests.ps1      # Manifest generation utility
manifests/                    # Generated environment manifests
  manifest-dev.xml
  manifest-test.xml
  manifest-prod.xml
```

This solution enables true multi-environment coexistence! 🎉