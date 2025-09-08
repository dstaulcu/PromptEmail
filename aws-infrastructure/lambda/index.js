const https = require('https');
const http = require('http');
const url = require('url');

/**
 * AWS Lambda function to proxy Splunk HEC requests
 * Handles CORS and credential management
 */
exports.handler = async (event) => {
    console.info('[INFO] - Received event:', JSON.stringify(event, null, 2));
    
    const allowedOrigin = process.env.ALLOWED_ORIGIN || '*';
    const corsHeaders = {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': allowedOrigin,
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-Requested-With, X-API-Key'
    };
    
    try {
        // Handle preflight CORS requests
        if (event.httpMethod === 'OPTIONS') {
            console.info('[INFO] - Handling CORS preflight request');
            return createResponse(200, { message: 'CORS preflight handled' }, corsHeaders);
        }
        
        // Only allow POST requests for telemetry
        if (event.httpMethod !== 'POST') {
            console.info('[INFO] - Method not allowed:', event.httpMethod);
            return createResponse(405, { error: 'Method not allowed' }, corsHeaders);
        }
        
        // Validate required environment variables
        if (!process.env.SPLUNK_HEC_TOKEN || !process.env.SPLUNK_HEC_URL) {
            console.error('[ERROR] - Missing required environment variables');
            return createResponse(500, { error: 'Service configuration error' }, corsHeaders);
        }
        
        // Parse and validate request body
        let requestData;
        let parsedEvents = [];
        try {
            requestData = typeof event.body === 'string' ? event.body : JSON.stringify(event.body);
            
            // Parse each line as a separate JSON event (Splunk HEC format)
            const lines = requestData.trim().split('\n');
            for (const line of lines) {
                if (line.trim()) {
                    const eventData = JSON.parse(line);
                    
                    // Enrich with server-side information
                    if (eventData.event) {
                        // Add client IP address - try multiple sources to get real client IP
                        const clientIp = getClientIpAddress(event);
                        if (clientIp) {
                            eventData.event.client_ip_address = clientIp.ip;
                            eventData.event.client_ip_source = clientIp.source;
                        }
                        
                        // Add user agent information
                        if (event.requestContext && event.requestContext.identity && event.requestContext.identity.userAgent) {
                            eventData.event.server_user_agent = event.requestContext.identity.userAgent;
                        }
                                               
                        // Add Lambda execution context
                        //if (process.env.AWS_LAMBDA_FUNCTION_NAME) {
                        //    eventData.event.lambda_function_name = process.env.AWS_LAMBDA_FUNCTION_NAME;
                        //}
                    }
                    
                    parsedEvents.push(eventData);
                }
            }
            
            // Reconstruct the data for Splunk
            requestData = parsedEvents.map(event => JSON.stringify(event)).join('\n');
            
        } catch (parseError) {
            console.error('[ERROR] - Invalid JSON in request body:', parseError);
            return createResponse(400, { error: 'Invalid JSON in request body' }, corsHeaders);
        }
        
        // Forward to Splunk
        const splunkResponse = await forwardToSplunk(requestData);
        
        console.info('[INFO] - Splunk response status:', splunkResponse.statusCode);
        return createResponse(splunkResponse.statusCode, splunkResponse.body, corsHeaders);
        
    } catch (error) {
        console.error('[ERROR] - Handler error:', error);
        return createResponse(500, { 
            error: 'Internal server error',
            message: error.message 
        }, corsHeaders);
    }
};

/**
 * Extract the real client IP address from various sources
 * Prioritizes X-Forwarded-For and other proxy headers over direct connection IP
 * @param {Object} event - AWS Lambda event object
 * @returns {Object} Object with ip and source properties
 */
function getClientIpAddress(event) {
    const headers = event.headers || {};
    
    // Convert header names to lowercase for case-insensitive lookup
    const lowerHeaders = {};
    Object.keys(headers).forEach(key => {
        lowerHeaders[key.toLowerCase()] = headers[key];
    });
    
    // Priority order for IP address sources
    const ipSources = [
        // X-Forwarded-For is the standard header for proxied requests
        // Format: "client, proxy1, proxy2" - first IP is the original client
        {
            header: 'x-forwarded-for',
            source: 'X-Forwarded-For',
            parser: (value) => value.split(',')[0].trim()
        },
        
        // Alternative headers used by various proxies/load balancers
        {
            header: 'x-real-ip',
            source: 'X-Real-IP',
            parser: (value) => value.trim()
        },
        
        // Cloudflare
        {
            header: 'cf-connecting-ip',
            source: 'CF-Connecting-IP',
            parser: (value) => value.trim()
        },
        
        // Other common proxy headers
        {
            header: 'x-client-ip',
            source: 'X-Client-IP',
            parser: (value) => value.trim()
        },
        
        {
            header: 'x-forwarded',
            source: 'X-Forwarded',
            parser: (value) => value.split('for=')[1]?.split(';')[0]?.trim().replace(/"/g, '')
        },
        
        {
            header: 'forwarded-for',
            source: 'Forwarded-For',
            parser: (value) => value.split(',')[0].trim()
        },
        
        {
            header: 'forwarded',
            source: 'Forwarded',
            parser: (value) => value.split('for=')[1]?.split(';')[0]?.trim().replace(/"/g, '')
        }
    ];
    
    // Try each IP source in priority order
    for (const ipSource of ipSources) {
        const headerValue = lowerHeaders[ipSource.header];
        if (headerValue) {
            try {
                const extractedIp = ipSource.parser(headerValue);
                if (extractedIp && isValidIpAddress(extractedIp)) {
                    console.info(`[INFO] - Client IP found via ${ipSource.source}: ${extractedIp}`);
                    return {
                        ip: extractedIp,
                        source: ipSource.source
                    };
                }
            } catch (error) {
                console.warn(`[WARN] - Error parsing ${ipSource.source} header: ${error.message}`);
            }
        }
    }
    
    // Fallback to API Gateway's direct connection IP
    if (event.requestContext && event.requestContext.identity && event.requestContext.identity.sourceIp) {
        const sourceIp = event.requestContext.identity.sourceIp;
        console.info(`[INFO] - Using API Gateway sourceIp as fallback: ${sourceIp}`);
        return {
            ip: sourceIp,
            source: 'API-Gateway-Direct'
        };
    }
    
    console.warn('[WARN] - No client IP address could be determined');
    return null;
}

/**
 * Validate if a string is a valid IPv4 or IPv6 address
 * @param {string} ip - IP address to validate
 * @returns {boolean} True if valid IP address
 */
function isValidIpAddress(ip) {
    if (!ip) return false;
    
    // Remove any port number (e.g., "192.168.1.1:8080" -> "192.168.1.1")
    const cleanIp = ip.split(':')[0];
    
    // IPv4 regex
    const ipv4Regex = /^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$/;
    
    // IPv6 regex (simplified)
    const ipv6Regex = /^(?:[0-9a-fA-F]{1,4}:){7}[0-9a-fA-F]{1,4}$|^::1$|^::$/;
    
    // Check for private/local IPs that might indicate proxy issues
    const privateIpRanges = [
        /^127\./,          // 127.0.0.0/8
        /^10\./,           // 10.0.0.0/8
        /^172\.(1[6-9]|2[0-9]|3[01])\./,  // 172.16.0.0/12
        /^192\.168\./,     // 192.168.0.0/16
        /^169\.254\./      // 169.254.0.0/16 (link-local)
    ];
    
    const isValidFormat = ipv4Regex.test(cleanIp) || ipv6Regex.test(cleanIp);
    
    if (isValidFormat) {
        // Log warning for private IPs (might indicate configuration issue)
        const isPrivate = privateIpRanges.some(range => range.test(cleanIp));
        if (isPrivate) {
            console.warn(`[WARN] - Detected private IP address: ${cleanIp} - this might indicate a proxy configuration issue`);
        }
    }
    
    return isValidFormat;
}

/**
 * Forward request to Splunk HEC endpoint
 * @param {string} data - JSON data to send
 * @returns {Promise<Object>} Response from Splunk
 */
async function forwardToSplunk(data) {
    return new Promise((resolve, reject) => {
        try {
            const splunkUrl = new URL(process.env.SPLUNK_HEC_URL);
            
            // Ensure we're hitting the collector/event endpoint
            if (!splunkUrl.pathname.includes('/services/collector')) {
                splunkUrl.pathname = '/services/collector/event';
            }
            
            console.info('[INFO] - Forwarding to Splunk:', splunkUrl.href);
            
            const options = {
                hostname: splunkUrl.hostname,
                port: splunkUrl.port || (splunkUrl.protocol === 'https:' ? 443 : 80),
                path: splunkUrl.pathname + splunkUrl.search,
                method: 'POST',
                headers: {
                    'Authorization': `Splunk ${process.env.SPLUNK_HEC_TOKEN}`,
                    'Content-Type': 'application/json',
                    'Content-Length': Buffer.byteLength(data),
                    'User-Agent': 'OutlookEmailAssistant-Gateway/1.0'
                },
                // Handle self-signed certificates
                rejectUnauthorized: false
            };
            
            const protocol = splunkUrl.protocol === 'https:' ? https : http;
            
            const req = protocol.request(options, (res) => {
                let responseBody = '';
                
                res.on('data', (chunk) => {
                    responseBody += chunk;
                });
                
                res.on('end', () => {
                    console.info(`[INFO] - Splunk response: ${res.statusCode} - ${responseBody}`);
                    
                    let parsedBody;
                    try {
                        parsedBody = JSON.parse(responseBody);
                    } catch (e) {
                        parsedBody = { message: responseBody };
                    }
                    
                    resolve({
                        statusCode: res.statusCode,
                        body: parsedBody
                    });
                });
            });
            
            req.on('error', (error) => {
                console.error('[ERROR] - Request to Splunk failed:', error);
                reject(new Error(`Splunk request failed: ${error.message}`));
            });
            
            req.on('timeout', () => {
                console.error('[ERROR] - Request to Splunk timed out');
                req.destroy();
                reject(new Error('Request to Splunk timed out'));
            });
            
            // Set timeout
            req.setTimeout(25000); // 25 seconds (Lambda has 30s timeout)
            
            // Send the data
            req.write(data);
            req.end();
            
        } catch (error) {
            console.error('[ERROR] - Error setting up Splunk request:', error);
            reject(error);
        }
    });
}

/**
 * Create standardized API response
 * @param {number} statusCode - HTTP status code
 * @param {Object} body - Response body
 * @param {Object} headers - Additional headers
 * @returns {Object} API Gateway response format
 */
function createResponse(statusCode, body, headers = {}) {
    return {
        statusCode: statusCode,
        headers: {
            'Content-Type': 'application/json',
            ...headers
        },
        body: JSON.stringify(body, null, 2)
    };
}
