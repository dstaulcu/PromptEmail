const https = require('https');
const url = require('url');

/**
 * AWS Lambda function to proxy Splunk HEC requests
 * Handles CORS and forwards telemetry events to Splunk HTTP Event Collector
 */
exports.handler = async (event) => {
    // CORS headers for browser compatibility
    const corsHeaders = {
        "Content-Type": "application/json",
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With, X-Amz-Date, X-Api-Key",
        "Access-Control-Max-Age": "86400",
        "Access-Control-Allow-Credentials": "false"
    };

    console.log(`splunk-proxy: httpMethod=${event.httpMethod} origin=${(event.headers && (event.headers.Origin || event.headers.origin)) || 'none'}`);

    try {
        // Handle CORS preflight request
        if (event.httpMethod === "OPTIONS") {
            return {
                statusCode: 200,
                headers: corsHeaders,
                body: JSON.stringify({ message: "CORS preflight handled" }),
                isBase64Encoded: false
            };
        }

        // Extract configuration from environment variables
        const splunkHecToken = process.env.SPLUNK_HEC_TOKEN;
        const splunkHecUrl = process.env.SPLUNK_HEC_URL;
        const allowedOrigin = process.env.ALLOWED_ORIGIN || "*";
        const environment = process.env.ENVIRONMENT || "prod";
        
        // Optional Splunk metadata from environment variables (can be overridden by request)
        const defaultMetadata = {
            index: process.env.SPLUNK_INDEX || null,
            host: process.env.SPLUNK_HOST || null,
            source: process.env.SPLUNK_SOURCE || null,
            sourcetype: process.env.SPLUNK_SOURCETYPE || null
        };

        // Validate required configuration
        if (!splunkHecToken || !splunkHecUrl) {
            console.error('splunk-proxy: Missing required configuration');
            return {
                statusCode: 500,
                headers: corsHeaders,
                body: JSON.stringify({ 
                    error: "Splunk configuration missing",
                    message: "HEC token and URL must be configured in Lambda environment variables"
                }),
                isBase64Encoded: false
            };
        }

        // Parse request body
        let telemetryData;
        try {
            telemetryData = JSON.parse(event.body || '{}');
        } catch (err) {
            console.error('splunk-proxy: Invalid JSON payload:', err.message);
            return {
                statusCode: 400,
                headers: corsHeaders,
                body: JSON.stringify({ 
                    error: "Invalid JSON payload",
                    details: err.message 
                }),
                isBase64Encoded: false
            };
        }

        // Transform to HEC format if needed
        const hecPayload = Array.isArray(telemetryData) 
            ? telemetryData.map(event => transformToHecFormat(event, defaultMetadata))
            : transformToHecFormat(telemetryData, defaultMetadata);

        console.log(`splunk-proxy: Processing ${Array.isArray(hecPayload) ? hecPayload.length : 1} events`);

        // Forward to Splunk HEC
        const hecEndpoint = `${splunkHecUrl}/services/collector/event`;
        
        try {
            const splunkResponse = await forwardToSplunk(hecEndpoint, splunkHecToken, hecPayload);
            
            console.log(`splunk-proxy: Successfully forwarded events to Splunk`);
            
            return {
                statusCode: 200,
                headers: corsHeaders,
                body: JSON.stringify({ 
                    message: "Events forwarded successfully",
                    count: Array.isArray(hecPayload) ? hecPayload.length : 1,
                    splunkResponse: splunkResponse
                }),
                isBase64Encoded: false
            };
            
        } catch (splunkError) {
            console.error('splunk-proxy: Failed to forward to Splunk:', splunkError.message);
            return {
                statusCode: 502,
                headers: corsHeaders,
                body: JSON.stringify({ 
                    error: "Failed to forward events to Splunk",
                    details: splunkError.message 
                }),
                isBase64Encoded: false
            };
        }

    } catch (error) {
        console.error('splunk-proxy: Unexpected error:', error);
        return {
            statusCode: 500,
            headers: corsHeaders,
            body: JSON.stringify({ 
                error: "Internal server error",
                details: error.message 
            }),
            isBase64Encoded: false
        };
    }
};

/**
 * Transform event data to Splunk HEC format
 * @param {Object} eventData - The event data to transform
 * @param {Object} defaultMetadata - Default Splunk metadata from environment variables
 */
function transformToHecFormat(eventData, defaultMetadata) {
    // If already in HEC format, return as-is (but allow metadata override)
    if (eventData.event && eventData.time !== undefined) {
        // Apply metadata overrides even for pre-formatted HEC events
        const enhancedEvent = { ...eventData };
        
        // Apply default metadata if not specified in the event
        if (!enhancedEvent.index && defaultMetadata.index) {
            enhancedEvent.index = defaultMetadata.index;
        }
        if (!enhancedEvent.host && defaultMetadata.host) {
            enhancedEvent.host = defaultMetadata.host;
        }
        if (!enhancedEvent.source && defaultMetadata.source) {
            enhancedEvent.source = defaultMetadata.source;
        }
        if (!enhancedEvent.sourcetype && defaultMetadata.sourcetype) {
            enhancedEvent.sourcetype = defaultMetadata.sourcetype;
        }
        
        return enhancedEvent;
    }

    // Transform to HEC format with metadata priority:
    // 1. Event-specific metadata (highest priority)
    // 2. Default metadata from environment variables
    // 3. Hardcoded fallbacks (lowest priority)
    const hecEvent = {
        time: eventData.timestamp ? Math.floor(new Date(eventData.timestamp).getTime() / 1000) : Math.floor(Date.now() / 1000),
        event: eventData.data || eventData,
        index: eventData.index || defaultMetadata.index || null,
        source: eventData.source || defaultMetadata.source || null,
        sourcetype: eventData.sourcetype || defaultMetadata.sourcetype || null
    };

    // Add host information with priority: event > environment > none
    const hostValue = eventData.host || defaultMetadata.host;
    if (hostValue) {
        hecEvent.host = hostValue;
    }

    return hecEvent;
}

/**
 * Forward events to Splunk HEC endpoint
 */
function forwardToSplunk(hecEndpoint, hecToken, payload) {
    return new Promise((resolve, reject) => {
        const parsedUrl = url.parse(hecEndpoint);
        const payloadString = JSON.stringify(payload);
        
        const options = {
            hostname: parsedUrl.hostname,
            port: parsedUrl.port || (parsedUrl.protocol === 'https:' ? 443 : 80),
            path: parsedUrl.path,
            method: 'POST',
            headers: {
                'Authorization': `Splunk ${hecToken}`,
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(payloadString),
                'User-Agent': 'AWS-Lambda-HEC-Proxy/1.0'
            },
            timeout: 25000 // 25 second timeout
        };

        const req = https.request(options, (res) => {
            let responseBody = '';
            
            res.on('data', (chunk) => {
                responseBody += chunk;
            });
            
            res.on('end', () => {
                if (res.statusCode >= 200 && res.statusCode < 300) {
                    try {
                        const parsedResponse = JSON.parse(responseBody);
                        resolve(parsedResponse);
                    } catch (parseError) {
                        resolve({ text: responseBody, statusCode: res.statusCode });
                    }
                } else {
                    reject(new Error(`Splunk HEC responded with status ${res.statusCode}: ${responseBody}`));
                }
            });
        });

        req.on('error', (error) => {
            reject(new Error(`Network error forwarding to Splunk: ${error.message}`));
        });

        req.on('timeout', () => {
            req.destroy();
            reject(new Error('Timeout forwarding events to Splunk HEC'));
        });

        // Write payload and end request
        req.write(payloadString);
        req.end();
    });
}