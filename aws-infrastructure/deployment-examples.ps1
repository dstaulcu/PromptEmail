# Example deployment configuration for different environments

# Production deployment
.\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-prod" `
    -SplunkHecToken "YOUR_PRODUCTION_HEC_TOKEN" `
    -SplunkHecUrl "https://splunk-prod.your-company.com:8088" `
    -Region "us-east-1" `
    -Environment "prod" `
    -AllowedOrigin "https://your-account-outlook-assistant-prod.s3.us-east-1.amazonaws.com"

# Development deployment
.\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-dev" `
    -SplunkHecToken "YOUR_DEV_HEC_TOKEN" `
    -SplunkHecUrl "https://splunk-dev.your-company.com:8088" `
    -Region "us-east-1" `
    -Environment "dev" `
    -AllowedOrigin "*"

# Test deployment with specific origin
.\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-test" `
    -SplunkHecToken "YOUR_TEST_HEC_TOKEN" `
    -SplunkHecUrl "https://splunk-test.company.com:8088" `
    -Region "us-west-2" `
    -Environment "test" `
    -AllowedOrigin "https://localhost:3000"
