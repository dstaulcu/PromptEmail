# Splunk HEC API Gateway Setup

This directory contains the AWS infrastructure code to create an API Gateway with Lambda functions that proxy Splunk HEC (HTTP Event Collector) requests. This eliminates the need to store Splunk credentials in the client application and resolves CORS issues.

**Note**: This assumes you already have a Splunk instance (Enterprise, Cloud, or other) with HEC enabled. Most enterprises will use their existing Splunk infrastructure.

## Architecture

```
Browser App → API Gateway → Lambda Function → Your Existing Splunk HEC
```

### Benefits

- **Security**: Splunk HEC tokens stored securely in AWS Lambda environment variables
- **CORS**: Proper CORS handling without browser restrictions  
- **Scalability**: AWS Lambda automatically scales based on demand
- **Enterprise Integration**: Works with existing Splunk Enterprise, Splunk Cloud, or other HEC endpoints
- **Monitoring**: CloudWatch logs for debugging and monitoring
- **Cost**: Pay-per-request pricing model

## Files

- `deploy-splunk-gateway.ps1` - PowerShell script to deploy the entire stack
- `cloudformation-template.yaml` - CloudFormation template defining AWS resources
- `lambda/index.js` - Lambda function code to proxy Splunk requests
- `lambda/package.json` - Node.js package definition
- `test-api-gateway.ps1` - Test script to verify deployment
- `deployment-examples.ps1` - Example deployment commands
## Quick Setup

1. **Get Splunk HEC details** from your Splunk administrator:
   - HEC Token (e.g., `12345678-1234-1234-1234-123456789012`)
   - HEC URL (e.g., `https://splunk.company.com:8088` or `https://input-prd-your-realm.cloud.splunkcoud.com:8088`)

2. **Deploy the API Gateway**:
   ```powershell
   .\deploy-splunk-gateway.ps1 `
       -SplunkHecToken "YOUR_HEC_TOKEN" `
       -SplunkHecUrl "https://your-splunk.com:8088"
   ```

3. **Test the deployment**:
   ```powershell
   .\test-api-gateway.ps1
   ```

4. **Update your application** with the API Gateway URL from the deployment output.

## Detailed Usage

### Deployment Parameters

```powershell
.\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-prod" `
    -SplunkHecToken "your-hec-token-here" `
    -SplunkHecUrl "https://splunk.company.com:8088" `
    -Region "us-east-1" `
    -Environment "prod" `
    -AllowedOrigin "https://your-s3-bucket.s3.amazonaws.com"
```

### Parameters Explained

- `StackName`: CloudFormation stack name (default: outlook-assistant-splunk-gateway)
- `SplunkHecToken`: Your Splunk HTTP Event Collector token
- `SplunkHecUrl`: Base URL of your Splunk instance (e.g., https://splunk.company.com:8088)
- `Region`: AWS region for deployment (default: us-east-1)
- `Environment`: Environment name for the API Gateway stage (default: prod)
- `AllowedOrigin`: CORS allowed origin for your application

### API Endpoints

After deployment, you'll get an API Gateway URL like:
```
https://abc123def4.execute-api.us-east-1.amazonaws.com/prod
```

**Telemetry Endpoint**: `POST /telemetry`
- Accepts JSON payload in Splunk HEC format
- Automatically adds authentication headers
- Handles CORS preflight requests

### Usage in Your Application

Replace direct Splunk HEC calls with calls to your API Gateway:

```javascript
// Old way (direct to Splunk)
const response = await fetch('https://splunk.company.com:8088/services/collector/event', {
    method: 'POST',
    headers: {
        'Authorization': 'Splunk your-token-here',
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(telemetryData)
});

// New way (via API Gateway)
const response = await fetch('https://your-api-id.execute-api.us-east-1.amazonaws.com/prod/telemetry', {
    method: 'POST',
    headers: {
        'Content-Type': 'application/json'
    },
    body: JSON.stringify(telemetryData)
});
```

## Monitoring and Troubleshooting

### CloudWatch Logs

- **Lambda Logs**: `/aws/lambda/your-stack-name-splunk-gateway`
- **API Gateway Logs**: `/aws/apigateway/your-stack-name`

### Common Issues

1. **CORS Errors**: Ensure `AllowedOrigin` parameter matches your application domain
2. **403 Forbidden**: Check that your Splunk HEC token is valid and enabled
3. **Timeout Errors**: Verify Splunk URL is accessible from AWS (check security groups/NACLs)

### Testing

You can test the endpoint using curl:

```bash
curl -X POST https://your-api-id.execute-api.us-east-1.amazonaws.com/prod/telemetry \
  -H "Content-Type: application/json" \
  -d '{
    "event": {
      "eventType": "test",
      "message": "API Gateway test"
    },
    "sourcetype": "outlook-assistant",
    "source": "api-test"
  }'
```

## Security Considerations

- HEC token is stored in Lambda environment variables (encrypted at rest)
- API Gateway has no authentication by default - add API keys or Cognito if needed
- CORS is configured to allow only specified origins
- Lambda function runs with minimal IAM permissions

## Cost Estimation

Typical costs for moderate usage:
- **Lambda**: ~$0.20 per 1M requests
- **API Gateway**: ~$3.50 per 1M requests
- **CloudWatch Logs**: ~$0.50 per GB ingested

For most telemetry use cases, costs should be under $10/month.

## Cleanup

To remove all resources:

```powershell
aws cloudformation delete-stack --stack-name your-stack-name --region us-east-1
```
