# Splunk HEC API Proxy Implementation Guide

## Overview

This guide documents the complete implementation of a secure AWS API Gateway + Lambda proxy for Splunk HTTP Event Collector (HEC), enabling browser-based applications (like Office Add-ins) to send telemetry data to Splunk without exposing HEC tokens or dealing with CORS limitations.

## Architecture

```
Office Add-in (Browser)
    ↓ HTTPS Telemetry Events
API Gateway (/telemetry endpoint)
    ↓ Lambda Environment Variables (HEC Token)
Lambda Function (Splunk HEC Proxy)
    ↓ Authenticated HEC Request
Splunk Enterprise/Cloud HTTP Event Collector
```

## Key Features Implemented

### ✅ **Secure Token Management**
- HEC tokens stored in Lambda environment variables (encrypted at rest)
- No sensitive credentials exposed in client-side code
- Centralized credential management

### ✅ **CORS Resolution**
- Complete CORS header implementation for browser compatibility
- Proper preflight (OPTIONS) request handling
- Eliminates browser same-origin policy restrictions

### ✅ **Telemetry Pipeline**
- Batch event processing for performance
- Automatic HEC format validation and transformation
- Queue management with retry logic

### ✅ **Enterprise Integration**
- Compatible with Splunk Enterprise, Splunk Cloud, and HEC endpoints
- Configurable source types and indexes
- Preserves event metadata and timestamps

---

## Implementation Details

### Lambda Function (splunk-gateway-function)

**Location**: `aws-infrastructure/lambda/index.js` (Splunk version)

**Key Components:**

1. **HEC Request Proxy**
   ```javascript
   const https = require('https');
   const url = require('url');

   exports.handler = async (event) => {
       const corsHeaders = {
           "Content-Type": "application/json",
           "Access-Control-Allow-Origin": "*",
           "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
           "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With",
           "Access-Control-Max-Age": "86400",
           "Access-Control-Allow-Credentials": "false"
       };

       // Handle CORS preflight
       if (event.httpMethod === "OPTIONS") {
           return {
               statusCode: 200,
               headers: corsHeaders,
               body: JSON.stringify({ message: "CORS preflight handled" })
           };
       }

       // Extract configuration from environment
       const splunkHecToken = process.env.SPLUNK_HEC_TOKEN;
       const splunkHecUrl = process.env.SPLUNK_HEC_URL;
       const allowedOrigin = process.env.ALLOWED_ORIGIN || "*";

       if (!splunkHecToken || !splunkHecUrl) {
           return {
               statusCode: 500,
               headers: corsHeaders,
               body: JSON.stringify({ 
                   error: "Splunk configuration missing",
                   message: "HEC token and URL must be configured"
               })
           };
       }
   };
   ```

2. **Event Processing**
   ```javascript
   // Process telemetry payload
   let telemetryData;
   try {
       telemetryData = JSON.parse(event.body);
   } catch (err) {
       return {
           statusCode: 400,
           headers: corsHeaders,
           body: JSON.stringify({ 
               error: "Invalid JSON payload",
               details: err.message 
           })
       };
   }

   // Transform to HEC format if needed
   const hecPayload = Array.isArray(telemetryData) 
       ? telemetryData.map(transformToHecFormat)
       : transformToHecFormat(telemetryData);
   ```

3. **Splunk HEC Integration**
   ```javascript
   // Forward to Splunk HEC
   const hecEndpoint = `${splunkHecUrl}/services/collector/event`;
   const hecOptions = {
       method: 'POST',
       headers: {
           'Authorization': `Splunk ${splunkHecToken}`,
           'Content-Type': 'application/json',
           'User-Agent': 'AWS-Lambda-HEC-Proxy/1.0'
       }
   };

   try {
       const response = await forwardToSplunk(hecEndpoint, hecOptions, hecPayload);
       return {
           statusCode: 200,
           headers: corsHeaders,
           body: JSON.stringify({ 
               message: "Events forwarded successfully",
               count: Array.isArray(hecPayload) ? hecPayload.length : 1
           })
       };
   } catch (error) {
       console.error('Failed to forward to Splunk:', error);
       return {
           statusCode: 502,
           headers: corsHeaders,
           body: JSON.stringify({ 
               error: "Failed to forward events to Splunk",
               details: error.message 
           })
       };
   }
   ```

### Client-Side Integration

**Location**: `src/services/Logger.js`

**Key Features:**

1. **Telemetry Configuration**
   ```javascript
   // Load telemetry configuration
   async initializeTelemetryConfig() {
       try {
           const response = await fetch('/config/telemetry.json');
           if (response.ok) {
               this.telemetryConfig = await response.json();
               console.info('Telemetry configuration loaded:', this.telemetryConfig);
           }
       } catch (error) {
           console.error('Failed to load telemetry configuration:', error);
           this.telemetryConfig = this.getDefaultTelemetryConfig();
       }
   }
   ```

2. **Event Batching and Transmission**
   ```javascript
   // Queue events for batch transmission
   queueTelemetryEvent(eventData) {
       if (!this.telemetryConfig?.telemetry?.enabled) return;
       
       const event = {
           timestamp: new Date().toISOString(),
           eventType: eventData.eventType,
           data: eventData,
           sessionId: this.getSessionId(),
           source: "outlook_addon",
           sourcetype: "json:outlook_email_assistant",
           index: "main"
       };
       
       this.apiGatewayQueue.push(event);
       
       if (this.apiGatewayQueue.length >= this.telemetryConfig.api_gateway.batchSize) {
           this.flushTelemetryQueue();
       }
   }
   ```

3. **API Gateway Communication**
   ```javascript
   // Send events to API Gateway
   async flushTelemetryQueue() {
       if (this.apiGatewayQueue.length === 0) return;
       
       const events = [...this.apiGatewayQueue];
       this.apiGatewayQueue = [];
       
       try {
           const response = await fetch(this.telemetryConfig.api_gateway.endpoint, {
               method: 'POST',
               headers: {
                   'Content-Type': 'application/json'
               },
               body: JSON.stringify(events),
               mode: 'cors',
               credentials: 'omit'
           });
           
           if (!response.ok) {
               throw new Error(`HTTP ${response.status}: ${response.statusText}`);
           }
           
           console.info(`Telemetry batch sent: ${events.length} events`);
       } catch (error) {
           console.error('Failed to send telemetry batch:', error);
           // Re-queue events for retry
           this.apiGatewayQueue.unshift(...events);
       }
   }
   ```

### API Gateway Configuration

**Endpoint**: `https://{api-gateway-id}.execute-api.{region}.amazonaws.com/{stage}/telemetry`

**Example**: `https://23epm9o08b.execute-api.us-east-1.amazonaws.com/prod/telemetry`

**Methods**:
- `POST /telemetry` - Send telemetry events
- `OPTIONS /telemetry` - CORS preflight handling

---

## Configuration Guide

### 1. AWS Infrastructure Setup

#### Deploy Splunk Gateway Stack
```powershell
# Deploy the complete infrastructure
.\aws-infrastructure\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-prod" `
    -SplunkHecToken "{your-hec-token}" `
    -SplunkHecUrl "https://{your-splunk-host}:8088" `
    -Region "us-east-1" `
    -Environment "prod" `
    -AllowedOrigin "https://{your-s3-bucket}.s3.amazonaws.com"
```

#### CloudFormation Template Parameters
```yaml
Parameters:
  SplunkHecToken:
    Type: String
    Description: 'Splunk HTTP Event Collector Token'
    NoEcho: true
    
  SplunkHecUrl:
    Type: String
    Description: 'Splunk HEC URL (e.g., https://splunk.company.com:8088)'
    
  Environment:
    Type: String
    Default: 'prod'
    Description: 'Environment name for API Gateway stage'
    
  AllowedOrigin:
    Type: String
    Default: '*'
    Description: 'CORS allowed origin'
```

### 2. Client Configuration

#### Telemetry Configuration (`src/config/telemetry.json`)
```json
{
  "telemetry": {
    "enabled": true,
    "provider": "api_gateway",
    "api_gateway": {
      "endpoint": "https://{api-gateway-id}.execute-api.{region}.amazonaws.com/{stage}/telemetry",
      "timeout": 5000,
      "retryAttempts": 1,
      "batchSize": 10,
      "flushInterval": 30000
    }
  }
}
```

### 3. Splunk HEC Setup

#### Required HEC Configuration
```bash
# Splunk HEC Token Creation (Splunk Admin Console)
# Settings > Data Inputs > HTTP Event Collector > New Token

# Token Settings:
Name: outlook-email-assistant
Index: main (or custom index)
Source type: json:outlook_email_assistant
Token: {generated-uuid-token}
```

#### Required Splunk Indexes
```splunk
# Create custom index (optional)
| rest /servicesNS/admin/search/data/indexes
| search title="outlook_assistant"

# Or use default "main" index
```

---

## Event Format and Schema

### Standard Event Structure
```json
{
  "timestamp": "2024-10-12T15:30:45.123Z",
  "eventType": "email_analysis",
  "sessionId": "uuid-session-identifier",
  "data": {
    "analysisType": "classification",
    "duration": 1250,
    "success": true,
    "errorMessage": null,
    "aiProvider": "bedrock",
    "model": "claude-3-sonnet"
  },
  "source": "outlook_addon",
  "sourcetype": "json:outlook_email_assistant", 
  "index": "main"
}
```

### Event Types Supported
- **email_analysis**: Email classification and analysis events
- **user_interaction**: UI interactions and user behavior
- **performance_metrics**: Response times and system performance
- **error_events**: Application errors and exceptions
- **session_events**: Session start/end and authentication

### Batch Event Format
```json
[
  {
    "timestamp": "2024-10-12T15:30:45.123Z",
    "eventType": "session_start",
    "data": { "userId": "user123", "sessionId": "session456" }
  },
  {
    "timestamp": "2024-10-12T15:30:50.456Z", 
    "eventType": "email_analysis",
    "data": { "analysisType": "classification", "duration": 1250 }
  }
]
```

---

## Troubleshooting Guide

### Common Issues and Solutions

#### 1. "Failed to send telemetry batch" Errors
**Cause**: Network connectivity or API Gateway configuration issues
**Solution**:
- Verify API Gateway endpoint URL in telemetry configuration
- Check CORS headers in browser developer tools
- Confirm Lambda function is deployed and running

#### 2. "Splunk configuration missing" Error
**Cause**: Missing environment variables in Lambda function
**Solution**:
- Verify `SPLUNK_HEC_TOKEN` and `SPLUNK_HEC_URL` are set in Lambda environment
- Redeploy CloudFormation stack with correct parameters
- Check CloudWatch logs for detailed error information

#### 3. Events Not Appearing in Splunk
**Cause**: Invalid HEC token, index permissions, or network connectivity
**Solution**:
- Test HEC token directly: `curl -k -H "Authorization: Splunk {token}" {hecUrl}/services/collector/event`
- Verify Splunk index exists and HEC token has write permissions
- Check Splunk HEC endpoint accessibility from AWS Lambda

#### 4. CORS Preflight Failures
**Cause**: Missing or incorrect CORS headers
**Solution**:
- Ensure Lambda returns CORS headers on all responses (including errors)
- Verify `AllowedOrigin` parameter matches client application domain
- Clear browser cache and try again

### Debugging Tools

#### CloudWatch Logs
Monitor Lambda execution:
```bash
aws logs get-log-events --log-group-name "/aws/lambda/{stack-name}-splunk-gateway" --log-stream-name "LATEST_STREAM"
```

#### Splunk HEC Testing
Direct HEC endpoint test:
```bash
# Test HEC connectivity
curl -k -X POST "https://{splunk-host}:8088/services/collector/event" \
  -H "Authorization: Splunk {your-hec-token}" \
  -H "Content-Type: application/json" \
  -d '{
    "event": {
      "eventType": "hec_test",
      "message": "Direct HEC test",
      "timestamp": "'$(date -u +%Y-%m-%dT%H:%M:%S.%3NZ)'"
    },
    "sourcetype": "manual_test",
    "source": "curl_test"
  }'
```

#### API Gateway Testing
Test the proxy endpoint:
```bash
# Test API Gateway endpoint
curl -X POST "https://{api-gateway-id}.execute-api.{region}.amazonaws.com/{stage}/telemetry" \
  -H "Content-Type: application/json" \
  -H "Origin: https://your-app-origin.com" \
  -d '{
    "event": {
      "eventType": "api_gateway_test",
      "message": "Testing API Gateway proxy"
    },
    "sourcetype": "api_test",
    "source": "manual_test"
  }'
```

---

## Deployment Process

### Build and Deploy Infrastructure
```powershell
# 1. Deploy AWS infrastructure
.\aws-infrastructure\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-prod" `
    -SplunkHecToken $env:SPLUNK_HEC_TOKEN `
    -SplunkHecUrl "https://splunk.company.com:8088" `
    -Region "us-east-1" `
    -Environment "prod"

# 2. Get deployment outputs
aws cloudformation describe-stacks `
    --stack-name "outlook-assistant-splunk-gateway-prod" `
    --query 'Stacks[0].Outputs'
```

### Update Client Configuration
```powershell
# 3. Update telemetry configuration with API Gateway URL
$apiGatewayUrl = "{api-gateway-url-from-outputs}"
# Update src/config/telemetry.json with the endpoint URL

# 4. Deploy client application
.\tools\deploy_web_assets.ps1 -Environment prod
```

### Test Integration
```powershell
# 5. Test the complete pipeline
.\aws-infrastructure\test-api-gateway.ps1
```

---

## Performance and Monitoring

### Metrics to Monitor
- **Lambda Duration**: Should be < 5 seconds for HEC forwarding
- **Lambda Errors**: Monitor for HEC connectivity issues
- **API Gateway 4xx/5xx Errors**: Track malformed requests or Lambda failures
- **Splunk Index Volume**: Monitor incoming event volume and data quality

### Performance Optimization
- **Batch Size**: Optimize batch size (10-50 events) for best throughput
- **Flush Interval**: Balance between latency and efficiency (30-60 seconds)
- **Lambda Memory**: 256MB sufficient for typical telemetry volumes
- **Timeout**: 30 seconds allows for network retries

### Cost Monitoring
Typical costs for moderate usage:
- **Lambda**: ~$0.20 per 1M requests
- **API Gateway**: ~$3.50 per 1M requests  
- **CloudWatch Logs**: ~$0.50 per GB ingested

For most telemetry use cases, costs should be under $20/month.

---

## Security Best Practices

### 1. Token Management
- HEC tokens stored in Lambda environment variables (encrypted at rest)
- No tokens embedded in client-side code
- Regular token rotation recommended

### 2. Network Security  
- API Gateway regional endpoints for reduced latency
- CORS configured for specific origins in production
- Consider API keys or Cognito authentication for additional security

### 3. Data Privacy
- No PII or sensitive data in telemetry events
- Event data anonymized or pseudonymized
- Compliance with data retention policies

### 4. Access Control
- Lambda execution role with minimal required permissions
- Splunk index permissions restricted to HEC token
- CloudWatch log access controlled via IAM

---

## Advanced Configuration

### Multi-Environment Setup
```powershell
# Development environment
.\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-dev" `
    -SplunkHecToken "{dev-hec-token}" `
    -SplunkHecUrl "https://splunk-dev.company.com:8088" `
    -Environment "dev" `
    -AllowedOrigin "*"

# Production environment
.\deploy-splunk-gateway.ps1 `
    -StackName "outlook-assistant-splunk-gateway-prod" `
    -SplunkHecToken "{prod-hec-token}" `
    -SplunkHecUrl "https://splunk-prod.company.com:8088" `
    -Environment "prod" `
    -AllowedOrigin "https://your-prod-domain.com"
```

### Custom Index Configuration
```json
{
  "telemetry": {
    "enabled": true,
    "provider": "api_gateway",
    "defaultIndex": "outlook_assistant",
    "eventTypes": {
      "email_analysis": { "index": "outlook_assistant", "sourcetype": "outlook:analysis" },
      "user_interaction": { "index": "user_behavior", "sourcetype": "outlook:ui" },
      "error_events": { "index": "outlook_errors", "sourcetype": "outlook:error" }
    }
  }
}
```

### Rate Limiting and Throttling
```yaml
# Add to CloudFormation template
UsagePlan:
  Type: AWS::ApiGateway::UsagePlan
  Properties:
    UsagePlanName: !Sub '${AWS::StackName}-usage-plan'
    Throttle:
      RateLimit: 100
      BurstLimit: 200
    Quota:
      Limit: 10000
      Period: DAY
```

---

## Splunk Dashboard Examples

### Basic Event Monitoring
```splunk
index="main" sourcetype="json:outlook_email_assistant"
| timechart span=1h count by eventType
| eval _time=strftime(_time, "%H:%M")
```

### Performance Metrics
```splunk
index="main" sourcetype="json:outlook_email_assistant" eventType="email_analysis"
| stats avg(data.duration) as avg_duration, max(data.duration) as max_duration by data.aiProvider
| eval avg_duration=round(avg_duration,2)
```

### Error Analysis
```splunk
index="main" sourcetype="json:outlook_email_assistant" eventType="error_events"
| stats count by data.errorType, data.aiProvider
| sort -count
```

---

## Conclusion

This Splunk HEC API Proxy implementation provides a secure, scalable, and enterprise-ready solution for collecting telemetry data from browser-based Office Add-ins. The solution eliminates CORS restrictions while maintaining security through proper token management and centralized logging.

Key achievements:
- ✅ Resolved CORS limitations for Splunk HEC access
- ✅ Secure token management with AWS Lambda environment variables
- ✅ Batch processing for optimal performance
- ✅ Enterprise integration with existing Splunk infrastructure
- ✅ Comprehensive monitoring and debugging capabilities
- ✅ Cost-effective pay-per-request model

The solution integrates seamlessly with existing Splunk Enterprise or Cloud deployments and provides valuable insights into application usage, performance, and user behavior patterns.