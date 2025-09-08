# Developer Guide - PromptEmail Outlook Add-in

This guide provides comprehensive setup instructions and development workflows for the PromptEmail Outlook Add-in project.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Development Environment Setup](#development-environment-setup)
- [Project Structure](#project-structure)
- [Development Workflow](#development-workflow)
- [Build & Deployment](#build--deployment)
- [Testing & Debugging](#testing--debugging)
- [Code Standards](#code-standards)
- [Troubleshooting](#troubleshooting)

## Prerequisites

### Required Software

| Software | Version | Purpose | Installation |
|----------|---------|---------|--------------|
| **Node.js** | 16.x or later | JavaScript runtime and package management | [Download](https://nodejs.org/) |
| **npm** | 8.x or later | Package manager (included with Node.js) | Comes with Node.js |
| **PowerShell** | 5.1+ or Core 7+ | Build scripts and deployment tools | Windows built-in or [PowerShell Core](https://github.com/PowerShell/PowerShell) |
| **Git** | Latest | Version control | [Download](https://git-scm.com/) |
| **VS Code** | Latest | Recommended IDE | [Download](https://code.visualstudio.com/) |
| **Microsoft Outlook** | Office 365/2019+ | Target application for add-in | Office 365 subscription |

### Optional but Recommended

| Software | Purpose |
|----------|---------|
| **AWS CLI** | S3 deployment and cloud resources |
| **Office Add-in Debugger** | Advanced debugging capabilities |


### Development Machine Requirements

- **OS**: Windows 10/11 (primary), macOS (limited), Linux (limited)
- **Memory**: 8GB RAM minimum, 16GB recommended
- **Storage**: 2GB free space for dependencies and builds
- **Network**: Internet connectivity for package downloads and S3 deployment

## Development Environment Setup

### 1. Clone and Install Dependencies

```bash
# Clone the repository
git clone <repository-url>
cd outlook-email-assistant

# Install npm dependencies
npm install

# Verify installation
npm run build
```

### 2. VS Code Configuration

Install recommended VS Code extensions:

```json
{
  "recommendations": [
    "shd101wyy.markdown-preview-enhanced",
    "ms-vscode.powershell",
    "shd101wyy.markdown-preview-enhanced"
  ]
}
```

### 3. PowerShell Execution Policy

Enable PowerShell script execution:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
```

### 4. Configure Development Environment

Set up your development environment:

```powershell
# Navigate to tools directory
cd tools

# Set development environment registry key
.\outlook_installer.ps1 -SetEnvironmentRegistry Dev

# Verify configuration
.\outlook_installer.ps1 -ShowEnvironmentRegistry
```

## Project Structure

```
outlook-email-assistant/
├── src/                          # Source code
│   ├── taskpane/                 # Main task pane UI and logic
│   │   ├── taskpane.js             # Main application entry point
│   │   └── taskpane.html           # Task pane HTML template
│   ├── commands/                 # Ribbon button commands
│   │   ├── commands.js             # Command handlers
│   │   └── commands.html           # Commands HTML template
│   ├── services/                 # Core business logic
│   │   ├── AIService.js            # AI provider integrations
│   │   ├── EmailAnalyzer.js        # Email analysis engine
│   │   ├── ClassificationDetector.js # Security classification
│   │   ├── Logger.js               # Telemetry and logging
│   │   └── SettingsManager.js      # User preferences
│   ├── ui/                       # User interface components
│   │   ├── UIController.js         # UI state management
│   │   └── AccessibilityManager.js # Accessibility features
│   ├── assets/                   # Static assets
│   │   ├── css/                  # Stylesheets
│   │   ├── icons/                # Application icons
│   │   └── source/               # Source images
│   └── config/                   # Configuration files
│       ├── ai-providers.json       # AI provider settings
│       └── telemetry.json          # Telemetry configuration
├── public/                       # Built output and static files
├── tools/                        # Build and deployment scripts
│   ├── deploy_web_assets.ps1       # Main deployment script
│   ├── outlook_installer.ps1       # End-user installer
│   ├── outlook_addin_diagnostics.ps1 # Debugging utilities
│   └── deployment-environments.json # Environment configuration
├── docs/                         # Documentation and examples
├── package.json                     # Node.js project configuration
├── webpack.config.js               # Build configuration
└── README.md                       # Project overview
```

### Key Components

#### Core Services (`src/services/`)

- **`AIService.js`**: Handles communication with AI providers (OpenAI, Ollama, custom endpoints)
- **`EmailAnalyzer.js`**: Analyzes email content, extracts metadata, and processes responses
- **`ClassificationDetector.js`**: Detects security classifications in email headers/content
- **`Logger.js`**: Manages telemetry collection and Windows event logging
- **`SettingsManager.js`**: Persists user settings with Office 365 roaming support

#### User Interface (`src/ui/`, `src/taskpane/`)

- **`UIController.js`**: Centralized UI state management and event handling
- **`AccessibilityManager.js`**: Keyboard navigation and screen reader support
- **`taskpane.js`**: Main application logic and Office.js integration

## Development Workflow

### Deployment Process

1. **Initiate Development Build & Deployment to s3**
   ```powershell
   .\tools\deploy_web_assets.ps1 -Environment Dev
   ```

2. **Load Add-in in Outlook**
   ```powershell
   # Show Usage
   .\outlook_installer.ps1 -help

   # Install development version
   .\outlook_installer.ps1 -Environment Dev
   ```

3. **Make Code Changes**
   - Edit files in `src/` directory as appropriate
   - Refresh task pane in Outlook (Ctrl+R) to see changes

4. **Debug Issues**
   ```powershell
   # Clear Office cache if needed
   .\tools\outlook_cache_clear.ps1
   
   # Run comprehensive diagnostics
   .\tools\outlook_addin_diagnostics.ps1
   ```

### Branch Management

```bash
# Create feature branch
git checkout -b feature/your-feature-name

# Regular commits with descriptive messages
git commit -m "feat: add email sentiment analysis"

# Push and create pull request
git push origin feature/your-feature-name
```

### Code Style Guidelines

- **JavaScript**: ES6+ syntax, async/await preferred over Promises
- **Error Handling**: Always wrap async operations in try-catch blocks
- **Logging**: Use `Logger.js` for telemetry, console for development debugging

## Build & Deployment

### Development Build

```bash
# Development build with source maps
npm run dev

# Production build (optimized)
npm run build
```

### Environment Deployment

Deploy to different environments using the PowerShell deployment script:

```powershell
# Deploy to Development environment
.\tools\deploy_web_assets.ps1 -Environment Dev

# Deploy to Test environment  
.\tools\deploy_web_assets.ps1 -Environment Test

# Deploy to Production
.\tools\deploy_web_assets.ps1 -Environment Prod -IncrementVersion patch
```

### Version Management

The project uses semantic versioning:

```powershell
# Patch version (bug fixes): 1.1.20 → 1.1.21
.\tools\deploy_web_assets.ps1 -Environment Prod -IncrementVersion patch

# Minor version (new features): 1.1.20 → 1.2.0
.\tools\deploy_web_assets.ps1 -Environment Prod -IncrementVersion minor

# Major version (breaking changes): 1.1.20 → 2.0.0
.\tools\deploy_web_assets.ps1 -Environment Prod -IncrementVersion major
```

### Build Outputs

| File | Purpose | Location |
|------|---------|----------|
| `taskpane.bundle.js` | Main application logic | `public/` |
| `commands.bundle.js` | Ribbon command handlers | `public/` |
| `taskpane.html` | Task pane UI template | `public/` |
| `manifest.xml` | Add-in configuration | `public/` |
| `styles.css` | Compiled stylesheets | `public/` |

## Testing & Debugging

### Manual Testing Checklist

- [ ] **UI Functionality**: All buttons and inputs work correctly
- [ ] **AI Integration**: Test with different providers (OpenAI, Ollama, custom)
- [ ] **Email Analysis**: Verify tone detection, classification, and response generation
- [ ] **Settings Persistence**: Check settings save and load correctly
- [ ] **Accessibility**: Test keyboard navigation and screen reader compatibility
- [ ] **Error Handling**: Verify graceful error handling and user feedback

### Debugging Tools

#### PowerShell Diagnostic Script
```powershell
# Comprehensive system diagnostics
.\tools\outlook_addin_diagnostics.ps1

# Options include:
# 1. Office Add-in Registry Check
# 2. Office Settings Verification  
# 3. Network Connectivity Test
# 4. Cache and Temp Files Cleanup
```

#### Console Debugging
```javascript
// Use structured logging for debugging
console.group('Email Analysis');
console.info('[INFO] - Email content:', emailText);
console.info('[INFO] - Classification result:', classificationResult);
console.groupEnd();
```

### Performance Testing

Monitor performance using browser tools:

1. **Network Tab**: Check API response times
2. **Performance Tab**: Identify JavaScript bottlenecks  
4. **Console**: Review error logs and warnings

## Code Standards

### JavaScript Conventions

```javascript
// Use const/let, avoid var
const apiEndpoint = 'https://api.example.com';
let userPreferences = {};

// Async/await over Promises
async function analyzeEmail(emailText) {
    try {
        const result = await aiService.analyze(emailText);
        return result;
    } catch (error) {
        Logger.logError('Email analysis failed', error);
        throw error;
    }
}

```

### Error Handling

```javascript
// Consistent error handling pattern
try {
    const result = await riskyOperation();
    return result;
} catch (error) {
    Logger.logError('Operation failed', {
        operation: 'riskyOperation',
        error: error.message,
        stack: error.stack
    });
    
    // Show user-friendly message
    UIController.showError('An error occurred. Please try again.');
    return null;
}
```

## 🚨 Troubleshooting

### Common Development Issues

#### Build Errors

**Problem**: `npm run build` fails with module errors
```bash
# Solution: Clean install dependencies
rm -rf node_modules package-lock.json
npm install
```

**Problem**: Webpack build hanging or slow
```bash
# Solution: Clear webpack cache
npm run build -- --no-cache
```

#### Add-in Loading Issues

**Problem**: Add-in doesn't appear in Outlook
```powershell
# Solution: Check Office settings and registry
.\tools\outlook_addin_diagnostics.ps1
# Select option 2: Office Settings Check
```

**Problem**: Changes not reflected in Outlook
```powershell
# Solution: Clear cache and restart
.\tools\outlook_cache_clear.ps1
# Restart Outlook completely
```

#### Development Environment Issues

**Problem**: PowerShell execution policy errors
```powershell
# Run as Administrator
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Problem**: AWS deployment fails
```bash
# Check AWS CLI configuration
aws configure list
aws s3 ls # Test connectivity
```

### Performance Issues

**Problem**: Slow AI response times
- Check network connectivity to AI provider
- Verify API endpoint configuration
- Monitor API rate limits and quotas

**Problem**: UI freezing during operations
- Move heavy operations to web workers
- Add progress indicators for long-running tasks
- Implement proper error boundaries

### Getting Help

1. **Check Documentation**: Review README.md and this developer guide
2. **Run Diagnostics**: Use `outlook_addin_diagnostics.ps1` for system checks
3. **Check Logs**: Review browser console and Windows Event Logs
4. **Community Support**: Check Office Add-in documentation and forums
5. **Issue Reporting**: Create detailed bug reports with reproduction steps

---

## Contributing

When contributing to the project:

1. **Follow coding standards** outlined in this guide
2. **Write descriptive commit messages** using conventional commit format
3. **Test thoroughly** before submitting pull requests
4. **Update documentation** for any new features or changes
5. **Ensure backwards compatibility** when modifying public APIs

For questions or support, refer to the project maintainers and documentation resources.
