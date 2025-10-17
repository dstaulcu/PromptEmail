# Environment Configuration System

This document describes the secure environment configuration system for PromptEmail that separates home/public configurations from site-specific configurations.

## 🔒 Security Model

- **Public configs**: Committed to GitHub (home development values)
- **Mock site configs**: Git-ignored mock values for testing site structure at home
- **Real site configs**: Only exist on actual site systems with site-specific values
- **Template files**: Safe templates committed to GitHub for easy setup

### Three-Tier Security:
1. **Home System**: Mock site values for UI/structure testing (no sensitive data)
2. **Site System**: Real site-specific values (never leave site environment)
3. **GitHub**: Only public/home values (completely safe)

## 📁 File Structure

```
tools/
├── deployment-config.json                    # ✅ Public base config
├── deployment-config.work.template.json      # ✅ Template for site setup
├── deployment-config.work.json               # 🚫 Git-ignored (mock at home, real at site)
├── deployment-config.local.json              # 🚫 Git-ignored, local overrides
└── switch-environment.ps1                    # ✅ Environment switching script

src/config/
├── ai-providers.json                         # ✅ Public default providers
├── ai-providers.work.template.json           # ✅ Template for site providers
└── ai-providers.work.json                    # 🚫 Git-ignored (mock at home, real at site)
```

## 🚀 Quick Start

### First-Time Site Setup

1. **Setup site environment** (only needed once):
   ```powershell
   .\tools\switch-environment.ps1 setup
   ```

2. **Edit site configs**:
   - **At Home**: Edit with mock/placeholder values for testing UI structure
   - **At Site**: Edit with real site-specific values for actual deployment
   - Files: `tools\deployment-config.work.json` & `src\config\ai-providers.work.json`

3. **Switch to site environment**:
   ```powershell
   .\tools\switch-environment.ps1 work
   ```

### Daily Usage

- **Switch to home** (for GitHub commits):
  ```powershell
  .\tools\switch-environment.ps1 home
  ```

- **Switch to work/site** (for site development):
  ```powershell
  .\tools\switch-environment.ps1 work
  ```

- **Check current status**:
  ```powershell
  .\tools\switch-environment.ps1 status
  ```

## 🛡️ Security Features

### Automatic Git Protection
The `.gitignore` includes patterns that prevent sensitive files from being committed:
- `*.work.json` - Site-specific configurations
- `*.local.json` - Local overrides
- `*.site.json` - Any site-specific files

### Configuration Priority
The system loads configs in this priority order:
1. `deployment-config.local.json` (highest priority, local overrides)
2. `deployment-config.work.json` (work environment)
3. `deployment-config.json` (default/fallback)

### Template System
- Templates are safe to commit (contain no sensitive data)
- Real site configs are created from templates locally
- Templates include security reminders and proper structure

## 🔄 Workflow Examples

### Home Development → GitHub
```bash
# Switch to home environment
.\tools\switch-environment.ps1 home

# Verify no sensitive data
.\tools\switch-environment.ps1 status

# Commit and push
git add .
git commit -m "Feature update"
git push origin main
```

### Site Development
```bash
# Pull latest from GitHub
git pull origin main

# Switch to site environment
.\tools\switch-environment.ps1 work

# Deploy with site-specific configs
.\tools\deploy_web_assets.ps1 -Environment work
```

## 🛠️ Advanced Usage

### Custom Local Overrides
Create `deployment-config.local.json` for any local-specific overrides:
```json
{
  "environment": "development",
  "aiProvidersFile": "ai-providers.json",
  "s3Bucket": "my-personal-test-bucket",
  "description": "Personal development overrides"
}
```

### Environment Detection in Scripts
Your deployment scripts can detect the current environment:
```powershell
# In deploy scripts
$configPath = "tools\deployment-config.local.json"
if (-not (Test-Path $configPath)) {
    $configPath = "tools\deployment-config.work.json"
}
if (-not (Test-Path $configPath)) {
    $configPath = "tools\deployment-config.json"
}

$config = Get-Content $configPath | ConvertFrom-Json
$s3Bucket = $config.s3Bucket
```

## ⚠️ Security Reminders

1. **Never commit site configs**: Always verify with `git status` before committing
2. **Use secure systems**: Only edit site configs on approved site systems
3. **Regular audits**: Periodically run `switch-environment.ps1 status` to verify setup
4. **Clean transitions**: Always switch to home environment before GitHub operations

## 🤝 Benefits

- ✅ **Zero manual file editing**: Automated environment switching
- ✅ **Security by design**: Sensitive data never touches GitHub
- ✅ **Easy transitions**: Single command switches between environments
- ✅ **Template-driven**: Easy setup for new team members
- ✅ **Status verification**: Always know which environment you're in
- ✅ **Fail-safe**: System falls back to safe defaults if site configs missing