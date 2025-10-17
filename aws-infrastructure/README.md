# AWS Infrastructure - PromptEmail

This directory contains AWS CLI-based deployment scripts for the PromptEmail Outlook Add-in infrastructure. The infrastructure has been organized into separate services for better maintainability.

## 📁 Directory Structure

```
aws-infrastructure/
├── bedrock/                    # AWS Bedrock API Proxy
│   ├── deploy.ps1             # AWS CLI deployment script
│   ├── test.ps1               # Comprehensive testing script
│   └── lambda/                # Lambda function code
│       ├── index.js           # Bedrock proxy with user credentials
│       └── package.json       # Node.js dependencies
├── hosting/                    # S3 Static Website Hosting
│   ├── deploy.ps1             # S3 bucket configuration script
│   └── test.ps1               # Hosting validation script
├── splunk/                     # Splunk HEC API Proxy
│   ├── deploy.ps1             # AWS CLI deployment script
│   ├── test.ps1               # Comprehensive testing script
│   └── lambda/                # Lambda function code
│       ├── index.js           # Splunk HEC proxy
│       └── package.json       # Node.js package config
├── shared/                     # Shared resources
│   ├── api-gateway-resource-policy.template.json  # IP restriction template
│   └── README.md              # Usage instructions
└── README.md                   # This file
```

## 🚀 Quick Start

### Prerequisites
1. **AWS CLI Configuration**: Run `aws configure` with your credentials
2. **Environment Configuration**: Ensure `../tools/deployment-environments.json` contains your target environments
3. **PowerShell**: Scripts require PowerShell 5.1+ or PowerShell Core

### Deploy S3 Hosting
```powershell
cd aws-infrastructure/hosting
.\deploy.ps1 -Environment Dev
.\test.ps1 -Environment Dev
```

### Deploy Bedrock API Proxy
```powershell
cd aws-infrastructure/bedrock
.\deploy.ps1 -Environment Dev
.\test.ps1 -Environment Dev -Credentials "AKIA...:secret..."
```

### Deploy Splunk HEC Proxy
```powershell
cd aws-infrastructure/splunk
.\deploy.ps1 -Environment Dev -SplunkHecUrl "https://splunk.company.com:8088/services/collector/event"
.\test.ps1 -Environment Dev
```

### Get Help
All deployment scripts support comprehensive help:
```powershell
.\deploy.ps1 -Help
```

## 🔧 Services Overview

### 🌐 S3 Static Website Hosting
**Purpose**: Hosts the Outlook Add-in static files (HTML, CSS, JavaScript) with public web access.

**Features**:
- Multi-environment support (Dev, Test, Prod)
- Static website hosting configuration
- Public read access policies
- Bucket accessibility validation

**Configuration**: Uses `../../tools/deployment-environments.json` for environment settings

### 🤖 Bedrock API Proxy
**Purpose**: Enables browser-based applications to access AWS Bedrock with user-provided credentials while resolving CORS limitations.

**Features**:
- User credential authentication (no shared keys)
- Multiple credential formats support (direct, BedrockAPIKey, ABSK)
- Complete CORS handling
- Comprehensive error handling and logging

**Endpoints**: `/bedrock` (POST, OPTIONS)

### 📊 Splunk HEC Proxy
**Purpose**: Securely forwards telemetry data from browser applications to Splunk HTTP Event Collector.

**Features**:
- Secure HEC token management in Lambda environment variables
- Event batching and transformation to HEC format
- CORS handling for browser compatibility
- Telemetry pipeline optimization

**Endpoints**: `/telemetry` (POST, OPTIONS)

## 📋 Deployment Features

### Enhanced PowerShell Scripts
- **🎯 Environment-Based Configuration**: All scripts use centralized `deployment-environments.json`
- **📖 Comprehensive Help**: Every script supports `-Help` parameter with detailed usage examples
- **🚫 Auto-Help Display**: Invalid parameters automatically trigger help display
- **🔍 Parameter Validation**: Required parameters are validated with clear error messages
- **🔒 Secure Credential Handling**: Bedrock scripts use SecureString for sensitive data

### AWS CLI Advantages
- **🔍 Transparency**: Each resource creation is explicit and visible
- **🛠️ Easy Debugging**: Individual steps can be troubleshot separately
- **⚡ Fast Deployments**: No CloudFormation overhead
- **🎯 Flexible Updates**: Modify components without full stack updates
- **📊 Built-in Validation**: Checks existing resources before creating

### Deployment Capabilities
- **Resource Detection**: Automatically detects and updates existing resources
- **IAM Management**: Creates roles and policies automatically
- **Progress Tracking**: Clear, colored output with status indicators
- **Error Handling**: Graceful handling of common deployment scenarios
- **Deployment Info**: Each deployment saves configuration to JSON files for reference

## 🧪 Testing Suite

Both services include comprehensive test scripts that validate:

### Bedrock Tests
- ✅ CORS preflight requests
- ✅ Authentication without credentials (should fail)
- ✅ Authentication with valid credentials
- ✅ Invalid credential format handling
- ✅ Error response formatting

### Splunk Tests
- ✅ CORS preflight requests
- ✅ Single telemetry event processing
- ✅ Batch telemetry event processing
- ✅ Invalid JSON payload handling
- ✅ Empty payload processing

## 🔐 Security Features

### Bedrock Proxy
- User-provided credentials (no shared secrets)
- Bearer token authentication
- Comprehensive credential format validation
- Proper error handling without exposing sensitive data

### Splunk Proxy
- HEC tokens stored in Lambda environment variables (encrypted at rest)
- No credentials in client-side code
- Origin-based CORS restrictions
- Event data validation and transformation

## 📖 Documentation

- **Bedrock Guide**: `../../docs/BEDROCK_API_PROXY.md`
- **Splunk Guide**: `../../docs/SPLUNK_HEC_API_PROXY.md`

## 🔧 Configuration

### Environment Configuration
All deployment scripts use the centralized configuration file: `../tools/deployment-environments.json`

**Required Configuration Structure**:
```json
{
  "environments": {
    "Dev": {
      "region": "us-east-1",
      "bucketName": "promptemail-dev-12345",
      "domainSuffix": "dev"
    },
    "Test": {
      "region": "us-west-2", 
      "bucketName": "promptemail-test-12345",
      "domainSuffix": "test"
    },
    "Prod": {
      "region": "us-east-1",
      "bucketName": "promptemail-prod-12345", 
      "domainSuffix": "prod"
    }
  }
}
```

### Script Parameters

**Bedrock Deploy Script**:
```powershell
.\deploy.ps1 -Environment <Dev|Test|Prod> [-Help]
```

**Splunk Deploy Script**:
```powershell
.\deploy.ps1 -Environment <Dev|Test|Prod> -SplunkHecUrl <url> [-AllowedOrigin <origin>] [-Help]
```

**Hosting Deploy Script**:
```powershell
.\deploy.ps1 -Environment <Dev|Test|Prod> [-DryRun] [-Help]
```

### IP-Based Access Restrictions
Both services can use the shared API Gateway resource policy template:
```powershell
# 1. Copy and customize the template
cp shared/api-gateway-resource-policy.template.json my-policy.json
# Edit my-policy.json with your actual values

# 2. Apply IP restrictions
aws apigateway put-rest-api-policy `
    --rest-api-id YOUR_API_ID `
    --policy file://my-policy.json
```

### Environment Variables

**Bedrock**: No environment variables needed (user credentials via headers)

**Splunk**:
- `SPLUNK_HEC_TOKEN`: Your Splunk HTTP Event Collector token
- `SPLUNK_HEC_URL`: Base Splunk URL (e.g., https://splunk.company.com:8088)
- `ALLOWED_ORIGIN`: CORS allowed origin (default: *)

## 🔄 Migration from CloudFormation

This infrastructure has been migrated from CloudFormation to AWS CLI for better maintainability and transparency. Key improvements:

- **Individual Service Management**: Deploy and update services independently
- **Faster Iterations**: No CloudFormation stack dependencies
- **Better Debugging**: Clear visibility into each resource creation step
- **Flexible Configuration**: Easy parameter modification without template changes

## 💡 Tips

### Development Workflow
1. Deploy with your target environment (Dev, Test, Prod)
2. Use the `-Help` parameter to understand script options
3. Monitor CloudWatch logs: `/aws/lambda/{function-name}`
4. Test thoroughly using the provided test scripts
5. Review `deployment-info.json` files for deployment details

### Cost Optimization
- Use the minimum required Lambda memory (256MB is sufficient)
- Monitor API Gateway usage for rate limiting needs
- Consider regional deployment for reduced latency

### Monitoring
- **CloudWatch Logs**: All Lambda functions log to CloudWatch
- **API Gateway Logs**: Enable if detailed request logging is needed
- **Deployment Info**: Each deployment saves configuration to `deployment-info.json`

## 🆘 Troubleshooting

### Getting Help
All scripts provide comprehensive help when called with `-Help` or when invalid parameters are provided:
```powershell
.\deploy.ps1 -Help
```

### Common Issues
1. **AWS CLI not configured**: Run `aws configure` first
2. **IAM permissions**: Ensure your AWS user has Lambda, API Gateway, and IAM permissions
3. **Environment not found**: Check that your target environment exists in `deployment-environments.json`
4. **Property name mismatch**: Ensure configuration uses `region` (not `awsRegion`) property
5. **Missing required parameters**: Scripts will automatically display help with examples
6. **CORS errors**: Check that origins match your application domain

### PowerShell Script Issues
- **Syntax errors**: All scripts have been updated for PowerShell compatibility
- **Unicode characters**: All non-ASCII characters have been removed
- **Parameter validation**: Scripts validate required parameters and show auto-help

### Getting Help
- Check CloudWatch logs for Lambda errors
- Use test scripts to validate deployments
- Review deployment-info.json for configuration details
- Use `-Help` parameter on any script for detailed usage
- Refer to service-specific documentation for detailed troubleshooting
