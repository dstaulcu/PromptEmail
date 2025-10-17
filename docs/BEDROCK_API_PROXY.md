# AWS Bedrock API Proxy Implementation Guide

## Overview

This guide documents the complete implementation of a secure AWS API Gateway + Lambda proxy for AWS Bedrock, enabling browser-based applications (like Office Add-ins) to access Bedrock services with user-provided credentials. This solution solves CORS limitations while implementing proper authentication and security controls.

## Architecture

```
Office Add-in (Browser)
    ↓ HTTPS with User AWS Credentials
API Gateway (/bedrock endpoint)
    ↓ Bearer Token Authentication
Lambda Function (Bedrock Proxy)
    ↓ User AWS Credentials → AWS SDK
AWS Bedrock Runtime API
```

## Key Features Implemented

### ✅ **User Credential Authentication**
- Lambda uses **user-provided AWS credentials** instead of Lambda IAM role
- Supports multiple credential formats:
  - Direct: `AKIA...:secretAccessKey`
  - Direct with session token: `AKIA...:secretAccessKey:sessionToken`
  - BedrockAPIKey: `BedrockAPIKey-id:base64EncodedJSON`
  - ABSK: `ABSK...` (double-encoded format)

### ✅ **CORS Resolution**
- Complete CORS header implementation for browser compatibility
- Proper preflight (OPTIONS) request handling
- Cache control headers to prevent stale responses

### ✅ **Security Controls**
- IP-based access restrictions available
- API key-based usage plans implemented
- Request/response logging for monitoring

### ✅ **Error Handling & Debugging**
- Comprehensive error handling with proper CORS headers
- Detailed CloudWatch logging for troubleshooting
- Timeout optimization (increased from 3s to 30s)

---

## Implementation Details

### Lambda Function (bedrock-proxy-function)

**Location**: `aws-infrastructure/lambda/index.js`

**Key Components:**

1. **Credential Parsing**
   ```javascript
   // Parses Authorization: Bearer <token> header
   // Supports multiple credential formats
   const parseUserCredentials = (bearerToken) => {
       // Handle ABSK double-encoded format
       if (bearerToken.startsWith('ABSK')) {
           const originalFormat = Buffer.from(bearerToken, 'base64').toString('utf-8');
           return parseUserCredentials(originalFormat);
       }
       
       // Handle BedrockAPIKey format
       if (bearerToken.startsWith('BedrockAPIKey')) {
           const colonIndex = bearerToken.indexOf(':');
           const base64Part = bearerToken.substring(colonIndex + 1);
           const credentialsJson = Buffer.from(base64Part, 'base64').toString('utf-8');
           return JSON.parse(credentialsJson);
       }
       
       // Handle direct format: accessKeyId:secretAccessKey[:sessionToken]
       if (bearerToken.includes(':')) {
           const parts = bearerToken.split(':');
           return {
               accessKeyId: parts[0].trim(),
               secretAccessKey: parts[1].trim(),
               sessionToken: parts.length === 3 ? parts[2].trim() : undefined
           };
       }
   }
   ```

2. **CORS Headers**
   ```javascript
   const corsHeaders = {
       "Content-Type": "application/json",
       "Access-Control-Allow-Origin": "*",
       "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
       "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With, X-Amz-Date, X-Api-Key, X-Amz-Security-Token",
       "Access-Control-Max-Age": "86400",
       "Access-Control-Allow-Credentials": "false"
   };
   ```

3. **Bedrock Client with User Credentials**
   ```javascript
   const bedrockRuntime = new BedrockRuntimeClient({
       region: "us-east-1",
       credentials: {
           accessKeyId: awsCredentials.accessKeyId,
           secretAccessKey: awsCredentials.secretAccessKey,
           ...(awsCredentials.sessionToken && { sessionToken: awsCredentials.sessionToken })
       }
   });
   ```

### Client-Side Integration

**Location**: `src/services/AIService.js`

**Key Features:**

1. **Credential Format Support**
   ```javascript
   // Handle direct format: "AKIA...:secretkey"
   if (config.apiKey && !bearerToken && config.apiKey.includes(':') && config.apiKey.startsWith('AKIA')) {
       const parts = config.apiKey.split(':');
       if (parts.length >= 2 && parts.length <= 3) {
           awsCredentials = {
               accessKeyId: parts[0].trim(),
               secretAccessKey: parts[1].trim(),
               sessionToken: parts.length === 3 ? parts[2].trim() : undefined
           };
           bearerToken = config.apiKey; // Set as bearer token for Lambda
       }
   }
   ```

2. **CORS-Compliant Requests**
   ```javascript
   const fetchOptions = {
       method: 'POST',
       headers: {
           'Authorization': `Bearer ${bearerToken}`,
           'Content-Type': 'application/json'
       },
       body: requestBody,
       mode: 'cors',
       credentials: 'omit'
   };
   ```

### API Gateway Configuration

**Endpoint**: `https://{api-gateway-id}.execute-api.{region}.amazonaws.com/{stage}/bedrock`

**Example**: `https://abc123defg.execute-api.us-east-1.amazonaws.com/dev/bedrock`

**Stages**: 
- `dev` - Primary development endpoint
- `dev2` - Alternative endpoint for cache-busting during development

### Security Implementation

#### Option 1: IP-Based Resource Policy (Recommended)
```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Principal": "*",
      "Action": "execute-api:Invoke",
      "Resource": "arn:aws:execute-api:{region}:{account-id}:{api-id}/*/*/*",
      "Condition": {
        "IpAddress": {
          "aws:SourceIp": ["{your-home-ip}/32"]
        }
      }
    }
  ]
}
```

**Example**:
```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Principal": "*",
      "Action": "execute-api:Invoke",
      "Resource": "arn:aws:execute-api:us-east-1:123456789012:abc123defg/*/*/*",
      "Condition": {
        "IpAddress": {
          "aws:SourceIp": ["203.0.113.25/32"]
        }
      }
    }
  ]
}
```

**Application**: AWS Console → API Gateway → Resource Policy

#### Option 2: API Key-Based Usage Plan
- **API Key**: `{your-generated-api-key}`
- **Usage Plan**: `{usage-plan-name}`
- **Limits**: 1000 requests/day, 50 req/sec rate limit

**Example**:
- **API Key**: `AbCdEf123456789GhIjKlMnOpQrStUvWxYz`
- **Usage Plan**: `bedrock-proxy-plan`

---

## Configuration Guide

### 1. AWS Infrastructure Setup

#### Deploy Lambda Function
```bash
# Zip the lambda code
cd aws-infrastructure/lambda
Compress-Archive -Path * -DestinationPath ..\bedrock-proxy.zip -Force

# Update Lambda function
aws lambda update-function-code --function-name {your-lambda-function-name} --zip-file fileb://..\bedrock-proxy.zip

# Set timeout to 30 seconds
aws lambda update-function-configuration --function-name {your-lambda-function-name} --timeout 30
```

#### Deploy API Gateway
```bash
# Create new deployment
aws apigateway create-deployment --rest-api-id {your-api-gateway-id} --stage-name dev --description "Updated CORS headers for Lambda"
```

### 2. Client Configuration

#### AI Provider Configuration (`src/config/ai-providers.json`)
```json
{
  "bedrock1": {
    "label": "AWS Bedrock via API Gateway",
    "baseUrl": "https://{api-gateway-id}.execute-api.{region}.amazonaws.com/{stage}/bedrock",
    "defaultModel": "anthropic.claude-3-sonnet-20240229-v1:0",
    "apiFormat": "bedrock",
    "helpUrl": "https://aws.amazon.com/bedrock/",
    "helpText": "AWS Bedrock with user credentials. Format: 'accessKeyId:secretAccessKey' or 'BedrockAPIKey-id:base64EncodedCredentials'. User must have bedrock:InvokeModel permissions.",
    "blockedClassifications": ["confidential"],
    "models": [
      "anthropic.claude-3-sonnet-20240229-v1:0",
      "anthropic.claude-3-haiku-20240307-v1:0"
    ]
  }
}
```

### 3. User Credential Setup

#### Required AWS Permissions
```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": [
        "bedrock:InvokeModel"
      ],
      "Resource": "*"
    }
  ]
}
```

#### Credential Formats Supported

**Direct Format (Recommended)**:
```
AKIA{AccessKeyId}:{SecretAccessKey}
```
**Example**: `AKIAIOSFODNN7EXAMPLE:wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY`

**With Session Token**:
```
AKIA{AccessKeyId}:{SecretAccessKey}:{SessionToken}
```

**BedrockAPIKey Format**:
```
BedrockAPIKey-{timestamp}:{base64EncodedCredentials}
```
**Example**: `BedrockAPIKey-20241012:eyJhY2Nlc3NLZXlJZCI6IkFLSUEuLi4iLCJzZWNyZXRBY2Nlc3NLZXkiOiIuLi4ifQ==`

**Generate BedrockAPIKey** (using provided script):
```powershell
.\tools\generate-bedrock-apikey.ps1
```

---

## Troubleshooting Guide

### Common Issues and Solutions

#### 1. "No 'Access-Control-Allow-Origin' header is present"
**Cause**: Browser caching of failed preflight requests
**Solution**: 
- Clear Outlook cache: `.\tools\outlook_cache_clear.ps1`
- Use cache-busting parameters in development
- Verify Lambda returns CORS headers on all response paths

#### 2. "The security token included in the request is invalid"
**Cause**: Invalid or expired AWS credentials
**Solution**:
- Verify credentials format
- Check AWS user has `bedrock:InvokeModel` permissions
- Test credentials with AWS CLI: `aws bedrock list-foundation-models`

#### 3. "AWS credentials not configured"
**Cause**: Client-side credential parsing failure
**Solution**:
- Ensure credentials start with `AKIA` for direct format
- Verify colon separation: `accessKeyId:secretAccessKey`
- Check browser console for parsing errors

#### 4. Lambda timeout errors
**Cause**: 3-second default timeout too short for Bedrock calls
**Solution**: Increased timeout to 30 seconds (already implemented)

### Debugging Tools

#### CloudWatch Logs
Monitor Lambda execution:
```bash
aws logs get-log-events --log-group-name "/aws/lambda/{your-lambda-function-name}" --log-stream-name "LATEST_STREAM"
```

#### Client-Side Debugging
Enable verbose logging in add-in settings:
- Debug Logging: ON
- Check browser console for detailed request/response logs

#### Test Endpoints
**Manual Testing**:
```bash
# Test OPTIONS preflight
curl -i -X OPTIONS "https://{api-gateway-id}.execute-api.{region}.amazonaws.com/{stage}/bedrock" \
  -H "Origin: https://your-origin.com" \
  -H "Access-Control-Request-Headers: content-type,authorization"

# Test POST with credentials
curl -i -X POST "https://{api-gateway-id}.execute-api.{region}.amazonaws.com/{stage}/bedrock" \
  -H "Origin: https://your-origin.com" \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer AKIA{AccessKeyId}:{SecretAccessKey}" \
  -d '{"modelId":"anthropic.claude-3-sonnet-20240229-v1:0","input":"Hello"}'
```

---

## Deployment Process

### Build and Deploy Client
```bash
# Build webpack bundle
npm run build

# Deploy to S3
.\tools\deploy_web_assets.ps1 -Environment dev

# Clear cache
.\tools\outlook_cache_clear.ps1
```

### Update Lambda
```bash
# Update function code
cd aws-infrastructure/lambda
Compress-Archive -Path * -DestinationPath ..\bedrock-proxy.zip -Force
aws lambda update-function-code --function-name {your-lambda-function-name} --zip-file fileb://..\bedrock-proxy.zip

# Create API Gateway deployment
aws apigateway create-deployment --rest-api-id {your-api-gateway-id} --stage-name dev
```

---

## Performance and Monitoring

### Metrics to Monitor
- **Lambda Duration**: Should be < 30 seconds
- **API Gateway 4xx/5xx Errors**: Monitor authentication failures
- **CORS-related Browser Errors**: Check for preflight issues

### Optimization Notes
- **Timeout**: Set to 30 seconds for Bedrock calls
- **Memory**: 128MB sufficient for proxy operations
- **Caching**: CORS headers include cache control for browser optimization

---

## Security Best Practices

### 1. IP Restrictions
- Apply resource policy to limit access to known IP ranges
- Use AWS WAF for advanced protection

### 2. User Credential Management
- Users provide their own AWS credentials
- No shared or embedded credentials in client code
- Credentials transmitted over HTTPS only

### 3. Monitoring and Logging
- All requests logged to CloudWatch
- Failed authentication attempts tracked
- Usage patterns monitored through API Gateway metrics

### 4. Error Handling
- Sensitive information not exposed in error messages
- Consistent CORS headers on all responses including errors
- Graceful degradation for network issues

---

## Advanced Configuration

### Custom Domain Setup
To use custom domain instead of API Gateway URL:
1. Configure Route 53 domain
2. Create API Gateway custom domain mapping
3. Update SSL certificates
4. Modify client configuration

### Multi-Region Deployment
For global availability:
1. Deploy Lambda/API Gateway in multiple regions
2. Implement region selection logic in client
3. Configure Route 53 health checks
4. Update CORS policies for multiple endpoints

### Rate Limiting
Additional protection beyond usage plans:
1. Implement Lambda-level rate limiting
2. Use API Gateway throttling
3. Monitor per-user usage patterns
4. Implement circuit breaker patterns

---

## Conclusion

This implementation provides a robust, secure, and scalable solution for accessing AWS Bedrock from browser-based applications. The user credential authentication model ensures proper security while the comprehensive CORS handling enables seamless browser integration.

Key achievements:
- ✅ Resolved CORS limitations for Bedrock access
- ✅ Implemented user-credential based authentication
- ✅ Comprehensive error handling and debugging
- ✅ Production-ready security controls
- ✅ Scalable architecture for future enhancements

The solution is fully functional and ready for production use with proper monitoring and security controls in place.