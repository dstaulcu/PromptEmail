const { BedrockRuntimeClient, InvokeModelCommand } = require("@aws-sdk/client-bedrock-runtime");

exports.handler = async (event) => {
    // Keep a compact CORS header set and also provide multiValueHeaders to
    // increase the chance API Gateway will forward them exactly as given.
    const corsHeaders = {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With, X-Amz-Date, X-Api-Key, X-Amz-Security-Token",
        "Access-Control-Max-Age": "86400",
        "Access-Control-Allow-Credentials": "false"
    };

    const corsMultiValue = {
        "Content-Type": ["application/json"],
        "Access-Control-Allow-Origin": ["*"],
        "Access-Control-Allow-Methods": ["GET, POST, PUT, DELETE, OPTIONS"],
        "Access-Control-Allow-Headers": ["Content-Type, Authorization, X-Requested-With, X-Amz-Date, X-Api-Key, X-Amz-Security-Token"],
        "Access-Control-Max-Age": ["86400"],
        "Access-Control-Allow-Credentials": ["false"]
    };

    // Small debug log so we can see incoming method + origin in CloudWatch when diagnosing
    console.log(`bedrock-proxy: httpMethod=${event.httpMethod} origin=${(event.headers && (event.headers.Origin || event.headers.origin)) || 'none'}`);

    try {
        if (event.httpMethod === "OPTIONS") {
            return {
                statusCode: 200,
                headers: corsHeaders,
                multiValueHeaders: corsMultiValue,
                body: JSON.stringify({ message: "CORS preflight handled" }),
                isBase64Encoded: false
            };
        }

        // Extract user credentials from Authorization header
        const authHeader = event.headers?.Authorization || event.headers?.authorization;
        let awsCredentials = null;

        if (authHeader && authHeader.startsWith('Bearer ')) {
            const bearerToken = authHeader.substring(7); // Remove 'Bearer ' prefix
            console.log(`bedrock-proxy: processing bearer token format`);
            
            try {
                awsCredentials = parseUserCredentials(bearerToken);
                console.log(`bedrock-proxy: extracted credentials for access key: ${awsCredentials?.accessKeyId?.substring(0, 8)}...`);
            } catch (credError) {
                console.error('bedrock-proxy: failed to parse user credentials:', credError.message);
                return {
                    statusCode: 401,
                    headers: corsHeaders,
                    multiValueHeaders: corsMultiValue,
                    body: JSON.stringify({ 
                        error: "Invalid AWS credentials format", 
                        details: credError.message 
                    }),
                    isBase64Encoded: false
                };
            }
        }

        if (!awsCredentials) {
            return {
                statusCode: 401,
                headers: corsHeaders,
                multiValueHeaders: corsMultiValue,
                body: JSON.stringify({ 
                    error: "AWS credentials required",
                    message: "Please provide AWS credentials in Authorization header as Bearer token"
                }),
                isBase64Encoded: false
            };
        }

        // Create Bedrock client with user-provided credentials
        const bedrockRuntime = new BedrockRuntimeClient({
            region: "us-east-1",
            credentials: {
                accessKeyId: awsCredentials.accessKeyId,
                secretAccessKey: awsCredentials.secretAccessKey,
                ...(awsCredentials.sessionToken && { sessionToken: awsCredentials.sessionToken })
            }
        });

        let requestData = {};
        if (event.body) {
            try {
                requestData = JSON.parse(event.body);
            } catch (err) {
                // If body isn't JSON, keep it raw
                requestData = event.body;
            }
        }

        const modelId = (requestData && requestData.modelId) || "anthropic.claude-3-haiku-20240307-v1:0";

        const command = new InvokeModelCommand({
            modelId: modelId,
            body: JSON.stringify(requestData),
            contentType: "application/json",
            accept: "application/json"
        });

        const bedrockResponse = await bedrockRuntime.send(command);

        // bedrockResponse.body is a stream in many runtimes; decode safely
        let decoded = null;
        try {
            const raw = await streamToString(bedrockResponse.body);
            decoded = raw ? JSON.parse(raw) : {};
        } catch (err) {
            console.error('bedrock-proxy: failed to decode response body', err);
            decoded = { error: 'Failed to decode Bedrock response' };
        }

        return {
            statusCode: 200,
            headers: corsHeaders,
            multiValueHeaders: corsMultiValue,
            body: JSON.stringify(decoded),
            isBase64Encoded: false
        };

    } catch (error) {
        console.error('bedrock-proxy: handler error', error && error.message);
        return {
            statusCode: 500,
            headers: corsHeaders,
            multiValueHeaders: corsMultiValue,
            body: JSON.stringify({ error: error && error.message }),
            isBase64Encoded: false
        };
    }
};

// Helper: convert a stream (or string) to a string consistently
const streamToString = async (stream) => {
    if (!stream) return '';
    // If already a string
    if (typeof stream === 'string') return stream;
    // If a Uint8Array or ArrayBuffer-like
    if (stream instanceof Uint8Array) return new TextDecoder().decode(stream);
    // If stream is a readable stream with async iterator
    if (typeof stream[Symbol.asyncIterator] === 'function') {
        let chunks = '';
        for await (const chunk of stream) {
            if (typeof chunk === 'string') chunks += chunk;
            else if (chunk instanceof Uint8Array) chunks += new TextDecoder().decode(chunk);
            else chunks += JSON.stringify(chunk);
        }
        return chunks;
    }
    // Fallback: attempt JSON stringify
    try { return JSON.stringify(stream); } catch (e) { return String(stream); }
};

// Parse user credentials from Bearer token format
// Supports multiple formats:
// 1. "BedrockAPIKey-id:base64({"accessKeyId":"...", "secretAccessKey":"..."})"
// 2. "ABSK..." (double-encoded format)
// 3. "accessKeyId:secretAccessKey:sessionToken" (direct format)
const parseUserCredentials = (bearerToken) => {
    if (!bearerToken) {
        throw new Error('Bearer token is required');
    }

    // Handle ABSK double-encoded format
    if (bearerToken.startsWith('ABSK')) {
        try {
            const originalFormat = Buffer.from(bearerToken, 'base64').toString('utf-8');
            return parseUserCredentials(originalFormat); // Recursive call with decoded format
        } catch (err) {
            throw new Error('Failed to decode ABSK format credentials');
        }
    }

    // Handle BedrockAPIKey format: "BedrockAPIKey-id:base64EncodedJSON"
    if (bearerToken.startsWith('BedrockAPIKey')) {
        const colonIndex = bearerToken.indexOf(':');
        if (colonIndex === -1) {
            throw new Error('Invalid BedrockAPIKey format - missing colon separator');
        }

        const base64Part = bearerToken.substring(colonIndex + 1);
        try {
            const credentialsJson = Buffer.from(base64Part, 'base64').toString('utf-8');
            const credentials = JSON.parse(credentialsJson);
            
            if (!credentials.accessKeyId || !credentials.secretAccessKey) {
                throw new Error('Invalid credentials JSON - missing accessKeyId or secretAccessKey');
            }
            
            return {
                accessKeyId: credentials.accessKeyId,
                secretAccessKey: credentials.secretAccessKey,
                sessionToken: credentials.sessionToken || undefined
            };
        } catch (err) {
            throw new Error('Failed to parse BedrockAPIKey credentials: ' + err.message);
        }
    }

    // Handle direct format: "accessKeyId:secretAccessKey" or "accessKeyId:secretAccessKey:sessionToken"
    if (bearerToken.includes(':')) {
        const parts = bearerToken.split(':');
        if (parts.length < 2 || parts.length > 3) {
            throw new Error('Invalid direct credential format - expected accessKeyId:secretAccessKey[:sessionToken]');
        }

        return {
            accessKeyId: parts[0].trim(),
            secretAccessKey: parts[1].trim(),
            sessionToken: parts.length === 3 ? parts[2].trim() : undefined
        };
    }

    throw new Error('Unrecognized credential format - expected BedrockAPIKey, ABSK, or direct format');
};
