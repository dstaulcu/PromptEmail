# Shared AWS Infrastructure Resources

This directory contains reusable AWS infrastructure resources and templates that can be used across multiple services.

## Files

### `api-gateway-resource-policy.template.json`
**Purpose**: Template for restricting API Gateway access to specific IP addresses.

**Usage**: 
1. Copy the template and replace the placeholders with your actual values:
   - `{region}`: Your AWS region (e.g., `us-east-1`)
   - `{account-id}`: Your AWS account ID (12 digits)  
   - `{api-gateway-id}`: Your API Gateway ID from deployment output
   - `{your-ip-address}`: Your public IP address

2. Apply the policy to your API Gateway:
   ```bash
   aws apigateway put-rest-api-policy \
     --rest-api-id YOUR_API_ID \
     --policy file://your-policy.json
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
          "aws:SourceIp": [
            "203.0.113.25/32"
          ]
        }
      }
    }
  ]
}
```

**Security Note**: Never commit files with your actual IP addresses or account IDs to version control.

## Usage Across Services

Both Bedrock and Splunk API Gateways can use this IP restriction policy for enhanced security:

- **Bedrock**: Restricts access to your Bedrock proxy
- **Splunk**: Restricts access to your telemetry endpoint
- **Hosting**: Not applicable (S3 buckets need public access for web hosting)

## Getting Your IP Address

```bash
# Get your current public IP
curl ifconfig.me

# Or use PowerShell
(Invoke-WebRequest ifconfig.me).Content.Trim()
```