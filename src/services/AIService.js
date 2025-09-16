/**
 * AI Service for email analysis and response generation
 * Supports multiple AI providers and models
 */

export class AIService {
    /**
     * Extracts the response text from the API response data for each service
     * @param {Object} data - The response data from the API
     * @param {string} service - The AI service name
     * @returns {string} The extracted response text
     */
    extractResponseText(data, service) {
        switch (service) {
            case 'openai':
                return data.choices?.[0]?.message?.content || '';
            case 'ollama':
                // Ollama returns response in data.message.content or data.response
                const content = data.message?.content || data.response || data.text || '';
                
                // Handle empty responses from Ollama
                if (!content || content.trim() === '') {
                    if (data.done_reason === 'load') {
                        throw new Error('Model is still loading. Please try again in a moment.');
                    }
                    if (data.done && data.response === '') {
                        throw new Error('AI service returned an empty response. Please try again.');
                    }
                    throw new Error('No response content received from AI service.');
                }
                
                return content;
            default:
                // Fallback: try common OpenAI-compatible fields
                const fallbackContent = data.choices?.[0]?.message?.content || data.response || data.text || data.content || '';
                if (!fallbackContent || fallbackContent.trim() === '') {
                    throw new Error('No response content received from AI service.');
                }
                return fallbackContent;
        }
    }
    constructor(providersConfig = null) {
        // Store provider configuration from ai-providers.json
        this.providersConfig = providersConfig || {};
    }

    /**
     * Updates the provider configuration (for dynamic loading)
     * @param {Object} providersConfig - Provider configuration from ai-providers.json
     */
    updateProvidersConfig(providersConfig) {
        this.providersConfig = providersConfig || {};
        if (window.debugLog) window.debugLog('[VERBOSE] - Updated AIService provider config:', this.providersConfig);
    }

    /**
     * Gets the default model for a service from provider config or fallback
     * @param {string} service - AI service name
     * @param {Object} config - Configuration that might contain a model override
     * @returns {string} Model name to use
     */
    getDefaultModel(service, config = {}) {
        // Priority order: user config.model > provider defaultModel > hardcoded fallback
        if (config.model && config.model.trim()) {
            return config.model.trim();
        }
        
        const providerConfig = this.providersConfig[service];
        if (providerConfig && providerConfig.defaultModel) {
            return providerConfig.defaultModel;
        }
        
        // Ultimate hardcoded fallback for internal deployments
        return 'llama3:latest';
    }

    /**
     * Gets the max tokens setting for a service from provider config
     * @param {string} service - AI service name
     * @param {Object} config - Configuration that might contain overrides
     * @returns {number|undefined} Max tokens to use, or undefined to let the service decide
     */
    /**
     * Test the health/connectivity of an AI service
     * @param {Object} config - AI configuration
     * @returns {Promise<boolean>} True if service is healthy
     */
    async testConnection(config) {
        try {
            if (window.debugLog) window.debugLog('[VERBOSE] - Testing connection for service:', config.service);
            
            // Simple ping test with minimal prompt
            const testPrompt = "Hello, respond with 'OK'";
            await this.callAI(testPrompt, config, 'health-check');
            
            if (window.debugLog) window.debugLog('[VERBOSE] - Connection test passed');
            return true;
        } catch (error) {
            console.warn('[WARN] - Connection test failed:', error.message);
            return false;
        }
    }

    /**
     * Analyzes an email using AI
     * @param {Object} emailData - Email data from EmailAnalyzer
     * @param {Object} config - AI configuration
     * @returns {Promise<Object>} Analysis results
     */
    /**
     * Fetch available models from Ollama using /api/tags
     * @param {string} baseUrl - The base URL for Ollama
     * @returns {Promise<Array>} - Array of model names
     */
    static async fetchOllamaModels(baseUrl) {
        try {
            const url = `${baseUrl.replace(/\/$/, '')}/api/tags`;
            const response = await fetch(url);
            if (!response.ok) throw new Error(`Failed to fetch models: ${response.status}`);
            const data = await response.json();
            // Ollama returns { models: [{ name: ... }, ...] }
            return (data.models || []).map(m => m.name);
        } catch (err) {
            console.error('[ERROR] - Error fetching Ollama models:', err);
            return [];
        }
    }

    /**
     * Fetch available models from OpenAI-compatible API using /models
     * @param {string} baseUrl - The base URL for the API (should already include /v1)
     * @param {string} apiKey - The API key for authentication
     * @returns {Promise<Array>} - Array of model names
     */
    static async fetchOpenAICompatibleModels(baseUrl, apiKey) {
        try {
            const url = `${baseUrl.replace(/\/$/, '')}/models`;
            const headers = {};
            if (apiKey) {
                headers['Authorization'] = `Bearer ${apiKey}`;
            }
            
            const response = await fetch(url, { headers });
            
            if (!response.ok) {
                let errorMessage = `Failed to fetch models: ${response.status}`;
                
                // Provide specific error messages for common authentication issues
                if (response.status === 401) {
                    errorMessage = 'Authentication failed: Invalid or missing API key. Please check your API key in settings.';
                } else if (response.status === 403) {
                    errorMessage = 'Access forbidden: Your API key may not have permission to access models. Please verify your key has the correct permissions.';
                } else if (response.status === 404) {
                    errorMessage = 'Endpoint not found: The models endpoint may not be available. Please verify your endpoint URL is correct.';
                } else if (response.status >= 500) {
                    errorMessage = 'Server error: The API server is experiencing issues. Please try again later.';
                }
                
                throw new Error(errorMessage);
            }
            
            const data = await response.json();
            // OpenAI-compatible APIs return { data: [{ id: ... }, ...] }
            return (data.data || []).map(m => m.id);
        } catch (err) {
            console.error('[ERROR] - Error fetching OpenAI-compatible models:', err);
            // Re-throw with original error message if it's already detailed
            throw err;
        }
    }
    
    async analyzeEmail(emailData, config) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Starting email analysis...');
        if (window.debugLog) window.debugLog('[VERBOSE] - Email data:', emailData);
        if (window.debugLog) window.debugLog('[VERBOSE] - AI provider config:', config);
        
        const prompt = this.buildAnalysisPrompt(emailData);
        if (window.debugLog) window.debugLog('[VERBOSE] - Built analysis prompt:', prompt);
        
        try {
            if (window.debugLog) window.debugLog('[VERBOSE] - Calling AI for analysis...');
            const response = await this.callAI(prompt, config, 'analysis');
            if (window.debugLog) window.debugLog('[VERBOSE] - Raw analysis response:', response);
            
            const parsed = this.parseAnalysisResponse(response);
            console.info('[INFO] - Parsed LLM analysis result:', parsed);
            return parsed;
        } catch (error) {
            console.error('[ERROR] - Email analysis failed:', error);
            throw new Error('Failed to analyze email: ' + error.message);
        }
    }

    /**
     * Generates a response to an email
     * @param {Object} emailData - Original email data
     * @param {Object} analysis - Email analysis results
     * @param {Object} config - Configuration including AI and response settings
     * @returns {Promise<Object>} Generated response
     */
    async generateResponse(emailData, analysis, config) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Starting response generation...');
        if (window.debugLog) window.debugLog('[VERBOSE] - Email data:', emailData);
        if (window.debugLog) window.debugLog('[VERBOSE] - Analysis:', analysis);
        if (window.debugLog) window.debugLog('[VERBOSE] - Config:', config);
        
        // Ensure analysis is not null - provide default if missing
        if (!analysis) {
            console.warn('[WARN] - Analysis is null, providing default analysis structure');
            analysis = {
                keyPoints: ['No analysis available'],
                sentiment: 'neutral',
                responseStrategy: 'respond professionally'
            };
        }
        
        const prompt = this.buildResponsePrompt(emailData, analysis, config);
        if (window.debugLog) window.debugLog('[VERBOSE] - Built response prompt:', prompt);
        
        try {
            if (window.debugLog) window.debugLog('[VERBOSE] - Calling AI for response generation...');
            const response = await this.callAI(prompt, config, 'response');
            if (window.debugLog) window.debugLog('[VERBOSE] - Raw response generation result:', response);
            
            const parsed = this.parseResponseResult(response);
            console.info('[INFO] - Parsed LLM response result:', parsed);
            return parsed;
        } catch (error) {
            console.error('[ERROR] - Response generation failed:', error);
            throw new Error('Failed to generate response: ' + error.message);
        }
    }

    /**
     * Generates follow-up suggestions for sent emails
     * @param {Object} emailData - Original sent email data
     * @param {Object} analysis - Email analysis results
     * @param {Object} config - Configuration including AI and response settings
     * @returns {Promise<Object>} Generated follow-up suggestions
     */
    async generateFollowupSuggestions(emailData, analysis, config) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Starting follow-up suggestions generation...');
        if (window.debugLog) window.debugLog('[VERBOSE] - Email data:', emailData);
        if (window.debugLog) window.debugLog('[VERBOSE] - Analysis:', analysis);
        if (window.debugLog) window.debugLog('[VERBOSE] - Config:', config);
        
        // Ensure analysis is not null - provide default if missing
        if (!analysis) {
            console.warn('[WARN] - Analysis is null, providing default analysis structure');
            analysis = {
                keyPoints: ['Sent email content analyzed'],
                sentiment: 'neutral',
                responseStrategy: 'generate appropriate follow-up actions'
            };
        }
        
        const prompt = this.buildFollowupPrompt(emailData, analysis, config);
        if (window.debugLog) window.debugLog('[VERBOSE] - Built follow-up prompt:', prompt);
        
        try {
            if (window.debugLog) window.debugLog('[VERBOSE] - Calling AI for follow-up suggestions generation...');
            const response = await this.callAI(prompt, config, 'followup');
            if (window.debugLog) window.debugLog('[VERBOSE] - Raw follow-up suggestions result:', response);
            
            const parsed = this.parseFollowupResult(response);
            console.info('[INFO] - Parsed LLM follow-up suggestions result:', parsed);
            return parsed;
        } catch (error) {
            console.error('[ERROR] - Follow-up suggestions generation failed:', error);
            throw new Error('Failed to generate follow-up suggestions: ' + error.message);
        }
    }

    /**
     * Refines an existing response based on user feedback
     * @param {Object} currentResponse - Current response object
     * @param {string} instructions - User refinement instructions
     * @param {Object} config - AI configuration
     * @param {Object} responseSettings - Response generation settings (length, tone)
     * @returns {Promise<Object>} Refined response
     */
    async refineResponse(currentResponse, instructions, config, responseSettings = null) {
        const prompt = this.buildRefinementPrompt(currentResponse, instructions, responseSettings);
        
        try {
            const response = await this.callAI(prompt, config, 'refinement');
            return this.parseResponseResult(response);
        } catch (error) {
            console.error('[ERROR] - Response refinement failed:', error);
            throw new Error('Failed to refine response: ' + error.message);
        }
    }

    /**
     * Refines response with conversation history for maintained context
     * @param {Object} currentResponse - The current response object
     * @param {string} instructions - New refinement instructions
     * @param {Object} config - AI configuration
     * @param {Object} responseSettings - Response settings (length, tone)
     * @param {Object} originalEmailContext - Original email context
     * @param {Array} conversationHistory - Previous refinement steps
     * @returns {Promise<Object>} Refined response with maintained context
     */
    async refineResponseWithHistory(currentResponse, instructions, config, responseSettings = null, originalEmailContext = null, conversationHistory = []) {
        const prompt = this.buildRefinementPromptWithHistory(
            currentResponse, 
            instructions, 
            responseSettings, 
            originalEmailContext, 
            conversationHistory
        );
        
        try {
            const response = await this.callAI(prompt, config, 'refinement');
            return this.parseResponseResult(response);
        } catch (error) {
            console.error('[ERROR] - Response refinement with history failed:', error);
            throw new Error('Failed to refine response with history: ' + error.message);
        }
    }

    /**
     * Builds the prompt for email analysis
     * @param {Object} emailData - Email data
     * @returns {string} Analysis prompt
     */
    buildAnalysisPrompt(emailData) {
        const dateStr = emailData.date ? new Date(emailData.date).toLocaleString() : 'Compose Mode';
        return `Please analyze the following email and provide insights:

**Email Details:**
From: ${emailData.from}
Subject: ${emailData.subject}
Recipients: ${emailData.recipients}
Sent: ${dateStr}
Length: ${emailData.bodyLength} characters

**Email Content:**
${emailData.cleanBody || emailData.body}

**Analysis Request:**
Please provide a structured analysis including:

1. **Key Points**: List the main points or topics discussed (3-5 bullet points)
2. **Sentiment**: Describe the overall tone and sentiment of the email
3. **Intent**: What is the sender trying to accomplish?
4. **Urgency Level**: Rate the urgency from 1-5 and explain why
5. **Due Dates**: Carefully scan for any deadlines, due dates, meetings, deadlines, submission dates, or time-sensitive requirements. Look for phrases like "due by", "deadline", "by [date]", "needs to be completed", "meeting on", "expires", etc. Mark as urgent if within 3 days or if explicitly marked as urgent.
6. **Action Items**: What actions are requested or implied?
7. **Recommended Response Strategy**: How should this email be approached in a response?

Format your response as JSON with the following structure:
{
    "keyPoints": ["point1", "point2", "point3"],
    "sentiment": "description of sentiment and tone",
    "intent": "what the sender wants to accomplish",
    "urgencyLevel": number,
    "urgencyReason": "explanation of urgency rating",
    "dueDates": [
        {
            "date": "YYYY-MM-DD or 'unspecified'",
            "time": "HH:MM or 'unspecified'", 
            "description": "what is due or when the meeting/deadline is",
            "isUrgent": true/false
        }
    ],
    "actions": ["action1", "action2"],
    "responseStrategy": "recommended approach for responding"
}`;
    }

    /**
     * Builds the prompt for response generation
     * @param {Object} emailData - Original email data
     * @param {Object} analysis - Email analysis
     * @param {Object} config - Response configuration
     * @returns {string} Response generation prompt
     */
    buildResponsePrompt(emailData, analysis, config) {
        const lengthMap = {
            1: 'very brief (1-2 sentences)',
            2: 'brief (1 short paragraph)',
            3: 'medium length (2-3 paragraphs)',
            4: 'detailed (3-4 paragraphs)',
            5: 'very detailed (4+ paragraphs)'
        };

        const toneMap = {
            1: 'very casual and friendly',
            2: 'casual but respectful',
            3: 'professional and courteous',
            4: 'formal and business-like',
            5: 'very formal and ceremonious'
        };

        // Check if we're in a very casual tone that could benefit from creativity
        const isVeryCasualTone = config.tone === 1 || config.tone === '1';
        
        let prompt = `You are an AI email assistant helping with email-related tasks. Based on the context below, help create appropriate email content.\n\n`;
        
        if (isVeryCasualTone) {
            prompt += `Generate email content with creative freedom - be engaging, fun, and personable while still being helpful:\n\n`;
        } else {
            prompt += `Generate professional email content based on the following context:\n\n`;
        }
        
        prompt += `**Original Email:**\n` +
            `From: ${emailData.from}\n` +
            `Subject: ${emailData.subject}\n` +
            `Sent: ${emailData.date ? new Date(emailData.date).toLocaleString() : 'Compose Mode'}\n` +
            `Content: ${emailData.cleanBody || emailData.body}\n\n` +
            `**Analysis Summary:**\n` +
            `- Key Points: ${(analysis && analysis.keyPoints) ? analysis.keyPoints.join(', ') : 'Not analyzed'}\n` +
            `- Sentiment: ${(analysis && analysis.sentiment) || 'Not analyzed'}\n` +
            `- Recommended Strategy: ${(analysis && analysis.responseStrategy) || 'Not analyzed'}\n\n` +
            `**Response Requirements:**\n` +
            `- Length: ${lengthMap[config.length] || 'medium length'}\n` +
            `- Tone: ${toneMap[config.tone] || 'professional'}`;

        // Add creativity boost for very casual tone
        if (isVeryCasualTone) {
            prompt += `\n\n**CREATIVE MODE:**\n` +
                `- Feel free to be witty, playful, and engaging\n` +
                `- Use humor and personality as appropriate\n` +
                `- Don't be afraid to be creative with language and approach\n` +
                `- Keep it fun and personable while maintaining respect`;
        }

        // Note: Custom instructions removed - now handled via interactive chat

        // Check if HTML table formatting might be needed (simplified detection)
        const mightNeedHtmlTables = false; // Will be handled by chat interface
        if (mightNeedHtmlTables) {
            prompt += `\n\n**IMPORTANT - Table Formatting Instructions:**\n` +
                `- If you include any tables, charts, or structured data, format them using HTML table syntax\n` +
                `- Use proper HTML table elements: <table>, <thead>, <tbody>, <tr>, <th>, <td>\n` +
                `- Apply inline CSS styling to make tables visually appealing:\n` +
                `  - border-collapse: collapse\n` +
                `  - borders around cells: border: 1px solid #ddd\n` +
                `  - header styling: background-color: #f5f5f5; font-weight: bold\n` +
                `  - padding in cells: padding: 8px\n` +
                `  - text alignment as appropriate\n` +
                `- Do NOT use markdown table syntax (| | format) - use only HTML tables\n` +
                `- Ensure tables are properly formatted and will render well in email clients`;
        }

        prompt += `\n\n**Output Requirements:**\n` +
            `Please generate appropriate email content that:\n` +
            `1. Addresses the key points from the original email appropriately\n` +
            `2. Matches the requested tone and length\n` +
            `3. Is professional and well-structured\n` +
            `4. Includes appropriate greetings and closings when needed\n` +
            `5. Uses proper paragraph formatting with blank lines (double newlines) between paragraphs.\n\n` +
            `Return only the email content, ready to be used. Do not include subject line, email headers, or any introductory phrases. Output only the email content as it should appear.\n\n` +
            `**Note:** This could be for replying, forwarding, summarizing, or other email tasks - be flexible based on the context and user needs.`;

        return prompt;
    }

    /**
     * Builds the prompt for follow-up suggestions for sent emails
     * @param {Object} emailData - Sent email data
     * @param {Object} analysis - Email analysis results
     * @param {Object} config - Configuration including AI and response settings
     * @returns {string} Follow-up prompt
     */
    buildFollowupPrompt(emailData, analysis, config) {
        const lengthMap = {
            1: 'very brief (1-2 suggestions)',
            2: 'brief (2-3 suggestions)',
            3: 'medium (3-4 suggestions)',
            4: 'detailed (4-5 suggestions)',
            5: 'comprehensive (5+ suggestions)'
        };

        let prompt = `You are analyzing a sent email and providing follow-up suggestions.\n\n` +
            `**Sent Email Context:**\n` +
            `From: ${emailData.sender || 'Current User'}\n` +
            `To: ${emailData.from}\n` +
            `Subject: ${emailData.subject}\n` +
            `Sent: ${emailData.date ? new Date(emailData.date).toLocaleString() : 'Recently'}\n` +
            `Content: ${emailData.cleanBody || emailData.body}\n\n` +
            `**Analysis Summary:**\n` +
            `- Key Points: ${(analysis && analysis.keyPoints) ? analysis.keyPoints.join(', ') : 'Not analyzed'}\n` +
            `- Sentiment: ${(analysis && analysis.sentiment) || 'Not analyzed'}\n` +
            `- Context: ${(analysis && analysis.responseStrategy) || 'Not analyzed'}\n\n` +
            `**Suggestion Requirements:**\n` +
            `- Detail Level: ${lengthMap[config.length] || 'medium'}\n`;

        // Note: Custom instructions removed - now handled via interactive chat

        // Check if HTML table formatting might be needed (simplified detection)  
        const mightNeedHtmlTablesFollowup = false; // Will be handled by chat interface
        if (mightNeedHtmlTablesFollowup) {
            prompt += `\n\n**IMPORTANT - Table Formatting Instructions:**\n` +
                `- If you include any tables, charts, or structured data, format them using HTML table syntax\n` +
                `- Use proper HTML table elements: <table>, <thead>, <tbody>, <tr>, <th>, <td>\n` +
                `- Apply inline CSS styling to make tables visually appealing:\n` +
                `  - border-collapse: collapse\n` +
                `  - borders around cells: border: 1px solid #ddd\n` +
                `  - header styling: background-color: #f5f5f5; font-weight: bold\n` +
                `  - padding in cells: padding: 8px\n` +
                `  - text alignment as appropriate\n` +
                `- Do NOT use markdown table syntax (| | format) - use only HTML tables\n` +
                `- Ensure tables are properly formatted and will render well in email clients`;
        }

        prompt += `\n\n**Output Requirements:**\n` +
            `Based on this sent email, provide practical follow-up suggestions that consider:\n` +
            `1. What responses or reactions the recipients might have\n` +
            `2. Potential next steps or actions that might be needed\n` +
            `3. Timeline considerations for follow-up actions\n` +
            `4. Any deliverables, commitments, or expectations set in the email\n` +
            `5. Proactive steps to ensure successful outcomes\n\n` +
            `IMPORTANT: Do NOT write an email response or use salutations like "Hi [Name]" or "Dear [Name]". ` +
            `Do NOT include email signatures, greetings, or closing remarks. ` +
            `This is for the SENDER to review what they should do next after sending their email.\n\n` +
            `Format your response as actionable follow-up suggestions, not as an email to send. ` +
            `Use bullet points or numbered lists for clarity. Focus on what the SENDER should consider doing next, ` +
            `not what recipients should do. Start directly with the suggestions without any email formatting.`;

        return prompt;
    }

    /**
     * Parses follow-up suggestions result
     * @param {string} response - Raw AI response
     * @returns {Object} Parsed follow-up suggestions
     */
    parseFollowupResult(response) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Parsing follow-up suggestions response:', response);
        
        if (!response || typeof response !== 'string') {
            console.warn('[WARN] - Invalid follow-up suggestions response, using fallback');
            return {
                suggestions: 'No follow-up suggestions could be generated at this time.',
                type: 'followup'
            };
        }

        // Clean up the response
        let cleanedResponse = response.trim();
        
        // Remove any introductory phrases
        const introPatterns = [
            /^here are some follow-up suggestions?:?\s*/i,
            /^follow-up suggestions?:?\s*/i,
            /^based on.+?here are.+?:?\s*/i,
            /^suggested follow-up actions?:?\s*/i
        ];
        
        for (const pattern of introPatterns) {
            cleanedResponse = cleanedResponse.replace(pattern, '');
        }

        return {
            suggestions: cleanedResponse,
            type: 'followup',
            originalResponse: response
        };
    }

    /**
     * Builds the prompt for response refinement
     * @param {Object} currentResponse - Current response
     * @param {string} instructions - User instructions  
     * @param {Object} responseSettings - Response settings (length, tone)
     * @returns {string} Refinement prompt
     */
    /**
     * Detects if the user is requesting tables, charts, or structured data
     * @param {string} instructions - User's refinement instructions
     * @returns {boolean} True if tables/charts are requested
     */
    detectTableRequest(instructions) {
        if (!instructions) return false;
        
        const tableKeywords = [
            'table', 'chart', 'grid', 'columns', 'rows', 
            'tabular', 'spreadsheet', 'data table', 'comparison table',
            'matrix', 'schedule', 'timeline table', 'budget table',
            'organize in a table', 'format as table', 'show in table form',
            'create a table', 'make a table', 'put in table',
            'list in columns', 'structured format', 'tabulated'
        ];
        
        // Keywords for refinement requests about existing tables
        const tableRefinementKeywords = [
            'update the table', 'modify the table', 'change the table',
            'add to the table', 'remove from the table', 'adjust the table',
            'revise the table', 'improve the table', 'enhance the table',
            'expand the table', 'simplify the table', 'fix the table',
            'table formatting', 'table style', 'table appearance',
            'add columns', 'remove columns', 'add rows', 'remove rows',
            'sort the table', 'reorder the table', 'reorganize the table'
        ];
        
        const lowerInstructions = instructions.toLowerCase();
        return tableKeywords.some(keyword => lowerInstructions.includes(keyword)) ||
               tableRefinementKeywords.some(keyword => lowerInstructions.includes(keyword));
    }

    /**
     * Detects if the user is requesting creative, humorous, or entertaining content
     * @param {string} instructions - User's refinement instructions
     * @returns {boolean} True if creative/humorous content is requested
     */
    detectCreativeRequest(instructions) {
        if (!instructions) return false;
        
        const creativeKeywords = [
            'funny', 'humor', 'humorous', 'joke', 'jokes', 'amusing', 'witty',
            'clever', 'creative', 'entertaining', 'playful', 'sarcastic',
            'ironic', 'satirical', 'comedic', 'hilarious', 'laugh', 'laughter',
            'pun', 'puns', 'wordplay', 'tongue-in-cheek', 'lighthearted',
            'whimsical', 'quirky', 'zany', 'silly', 'goofy', 'ridiculous',
            'your-momma', 'yo mama', 'roast', 'roasting', 'burn', 'savage',
            'meme', 'memes', 'viral', 'trending', 'emoji spam', 'lol', 'lmao',
            'make it funny', 'add humor', 'be creative', 'be entertaining',
            'make me laugh', 'crack a joke', 'tell a joke', 'be witty',
            'more creative', 'more fun', 'more entertaining', 'spice it up',
            'jazz it up', 'make it interesting', 'add some flair', 'be bold',
            'think outside the box', 'unconventional', 'unexpected'
        ];
        
        const lowerInstructions = instructions.toLowerCase();
        return creativeKeywords.some(keyword => lowerInstructions.includes(keyword));
    }

    buildRefinementPrompt(currentResponse, instructions, responseSettings = null) {
        // Detect creative/humorous requests
        const isCreativeRequest = this.detectCreativeRequest(instructions);
        
        let settingsInstructions = '';
        
        if (responseSettings) {
            const lengthMap = {
                1: 'very brief (1-2 sentences)',
                2: 'brief (1 short paragraph)',
                3: 'medium length (2-3 paragraphs)',
                4: 'detailed (3-4 paragraphs)',
                5: 'very detailed (4+ paragraphs)'
            };

            const toneMap = {
                1: 'very casual and friendly',
                2: 'casual but respectful',
                3: 'professional and courteous',
                4: 'formal and business-like',
                5: 'very formal and ceremonious'
            };

            settingsInstructions = `
**Response Settings to Apply:**
- Length: ${lengthMap[responseSettings.length] || 'medium length'}
- Tone: ${toneMap[responseSettings.tone] || 'professional and courteous'}`;
        }

        const userInstructions = instructions.trim() 
            ? `**User's Refinement Instructions:**\n${instructions}` 
            : '';

        // Creative mode adjustments
        let creativityInstructions = '';
        if (isCreativeRequest) {
            creativityInstructions = `

**CREATIVE MODE ACTIVATED:**
- Prioritize humor, creativity, and entertainment value over professional constraints
- Feel free to be witty, clever, and amusing
- Use wordplay, jokes, and creative language as appropriate
- Don't worry about being "professional" - focus on being engaging and fun
- Be bold and creative with your language and approach`;
        }

        // Check if user is requesting tables/charts and add HTML formatting instructions
        const requiresHtmlTable = this.detectTableRequest(instructions);
        const htmlTableInstructions = requiresHtmlTable ? `

**IMPORTANT - Table Formatting Instructions:**
- If you include any tables, charts, or structured data, format them using HTML table syntax
- Use proper HTML table elements: <table>, <thead>, <tbody>, <tr>, <th>, <td>
- Apply inline CSS styling to make tables visually appealing:
  - border-collapse: collapse
  - borders around cells: border: 1px solid #ddd
  - header styling: background-color: #f5f5f5; font-weight: bold
  - padding in cells: padding: 8px
  - text alignment as appropriate
- Example format:
  <table style="border-collapse: collapse; width: 100%; margin: 10px 0;">
    <thead>
      <tr>
        <th style="border: 1px solid #ddd; padding: 8px; background-color: #f5f5f5;">Header 1</th>
        <th style="border: 1px solid #ddd; padding: 8px; background-color: #f5f5f5;">Header 2</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td style="border: 1px solid #ddd; padding: 8px;">Data 1</td>
        <td style="border: 1px solid #ddd; padding: 8px;">Data 2</td>
      </tr>
    </tbody>
  </table>
- Do NOT use markdown table syntax (| | format) - use only HTML tables
- Ensure tables are properly formatted and will render well in email clients` : '';

        return `Please refine the following email response based on the settings and feedback provided:

**Current Response:**
${currentResponse.text}
${settingsInstructions}${creativityInstructions}
${userInstructions}
${htmlTableInstructions}

**Requirements:**
- Apply the settings and user feedback while maintaining professionalism
- Adjust length and tone as specified in the settings
- Keep the overall structure and flow intact unless specifically requested to change
- Ensure the response remains appropriate for business communication
- Maintain consistency in the refined tone and style

**Output Instructions:**
Return ONLY the refined email content without any prefixes, headers, or labels such as "Refined Response:" or similar. 
Do not include any introductory text or formatting markers. 
Provide only the email body text that should be sent.`;
    }

    /**
     * Builds refinement prompt with conversation history for maintaining context
     * @param {Object} currentResponse - Current response being refined
     * @param {string} instructions - New refinement instructions
     * @param {Object} responseSettings - Response settings (length, tone)
     * @param {Object} originalEmailContext - Original email context
     * @param {Array} conversationHistory - Previous refinement steps
     * @returns {string} Refinement prompt with conversation context
     */
    buildRefinementPromptWithHistory(currentResponse, instructions, responseSettings = null, originalEmailContext = null, conversationHistory = []) {
        let settingsInstructions = '';
        
        if (responseSettings) {
            const lengthMap = {
                1: 'very brief (1-2 sentences)',
                2: 'brief (1 short paragraph)',
                3: 'medium length (2-3 paragraphs)',
                4: 'detailed (3-4 paragraphs)',
                5: 'very detailed (4+ paragraphs)'
            };

            const toneMap = {
                1: 'very casual and friendly',
                2: 'casual but respectful',
                3: 'professional and courteous',
                4: 'formal and business-like',
                5: 'very formal and ceremonious'
            };

            settingsInstructions = `
**Response Settings to Apply:**
- Length: ${lengthMap[responseSettings.length] || 'medium length'}
- Tone: ${toneMap[responseSettings.tone] || 'professional and courteous'}`;
        }

        // Build conversation history context
        let conversationContext = '';
        if (originalEmailContext) {
            conversationContext += `
**Original Email Context:**
From: ${originalEmailContext.from}
Subject: ${originalEmailContext.subject}
Content: ${originalEmailContext.content.substring(0, 500)}${originalEmailContext.content.length > 500 ? '...' : ''}`;
        }

        if (conversationHistory.length > 0) {
            conversationContext += `

**Previous Refinement Steps:**`;
            conversationHistory.forEach((step, index) => {
                // Preserve HTML tables in conversation history - they're important for context
                const preservePreviousResponse = this.shouldPreserveFullContent(step.previousResponse);
                const preserveNewResponse = this.shouldPreserveFullContent(step.newResponse);
                
                const previousResponseText = preservePreviousResponse 
                    ? step.previousResponse 
                    : `${step.previousResponse.substring(0, 200)}${step.previousResponse.length > 200 ? '...' : ''}`;
                    
                const newResponseText = preserveNewResponse 
                    ? step.newResponse 
                    : `${step.newResponse.substring(0, 200)}${step.newResponse.length > 200 ? '...' : ''}`;
                
                conversationContext += `
Step ${step.step}: "${step.userInstruction}"
Previous Response: ${previousResponseText}
Result: ${newResponseText}`;
            });
        }

        const userInstructions = instructions.trim() 
            ? `**Current Refinement Request:**\n${instructions}` 
            : '';

        // Check if user is requesting tables/charts or if current response contains tables
        const requiresHtmlTable = this.detectTableRequest(instructions) || 
                                 this.shouldPreserveFullContent(currentResponse.text);
        const htmlTableInstructions = requiresHtmlTable ? `

**IMPORTANT - Table Formatting Instructions:**
- If you include any tables, charts, or structured data, format them using HTML table syntax
- Use proper HTML table elements: <table>, <thead>, <tbody>, <tr>, <th>, <td>
- Apply inline CSS styling to make tables visually appealing:
  - border-collapse: collapse
  - borders around cells: border: 1px solid #ddd
  - header styling: background-color: #f5f5f5; font-weight: bold
  - padding in cells: padding: 8px
  - text alignment as appropriate
- Do NOT use markdown table syntax (| | format) - use only HTML tables
- Ensure tables are properly formatted and will render well in email clients` : '';

        return `You are continuing a conversation to help refine email content. Please consider the full context and history when making this refinement. Be flexible about the type of email task - this could be for replying, forwarding, summarizing, composing, or other email needs.
${conversationContext}

**Current Email Content Being Refined:**
${currentResponse.text}
${settingsInstructions}
${userInstructions}
${htmlTableInstructions}

**Requirements:**
- Consider the conversation history and previous refinements to maintain consistency
- Apply the current refinement request while preserving good elements from previous iterations
- Adjust length and tone as specified in the settings
- Build upon the conversation context rather than starting from scratch
- Ensure the content remains appropriate for business communication
- Maintain consistency in the refined tone and style
- Be flexible about the email task type (reply, forward, summary, compose, etc.)

**Output Instructions:**
Return ONLY the refined email content without any prefixes, headers, or labels such as "Refined Response:" or similar. 
Do not include any introductory text or formatting markers. 
Provide only the email body text that should be sent.`;
    }

    /**
     * Makes API call to the specified AI service
     * @param {string} prompt - The prompt to send
     * @param {Object} config - AI configuration
     * @param {string} type - Type of request (analysis, response, refinement)
     * @returns {Promise<string>} AI response text
     */
    async callAI(prompt, config, type) {
        window.debugLog(`[VERBOSE] - Starting AI call for type: ${type}`);
        if (window.debugLog) window.debugLog('[VERBOSE] - Prompt:', prompt);
        if (window.debugLog) window.debugLog('[VERBOSE] - Config:', config);
        
        const service = config.service || 'openai';
        window.debugLog(`[VERBOSE] - Using service: ${service}`);

        if (service === 'custom') {
            if (window.debugLog) window.debugLog('[VERBOSE] - Calling custom endpoint...');
            return this.callCustomEndpoint(prompt, config);
        }

        // Validate service is configured in providers config
        if (!this.providersConfig[service] && service !== 'custom') {
            console.error(`[ERROR] - AIService Service not configured: ${service}`);
            throw new Error(`AI service '${service}' is not configured in providers config`);
        }

        // Build endpoint using provider configuration and user overrides
        let endpoint = this.buildEndpoint(service, config);
        if (window.debugLog) window.debugLog('[VERBOSE] - Final endpoint:', endpoint);

        let requestBody;
        let headers;

        if (service === 'ollama') {
            requestBody = {
                model: this.getDefaultModel(service, config),
                messages: [{ role: 'user', content: prompt }],
                stream: false
            };
            headers = { 'Content-Type': 'application/json' };
            if (window.debugLog) window.debugLog('[VERBOSE] - Built Ollama request body:', requestBody);
        } else {
            // For OpenAI, onsite1, onsite2, and other providers, use OpenAI-compatible format
            requestBody = this.buildRequestBody(prompt, service, config);
            headers = this.buildHeaders(service, config);
            if (window.debugLog) window.debugLog('[VERBOSE] - Built request body:', requestBody);
        }
        if (window.debugLog) window.debugLog('[VERBOSE] - Request headers:', headers);

        // Debug: Log POST request body to console
        if (window.debugLog) window.debugLog('[VERBOSE] - Making API call to endpoint:', endpoint);
        if (window.debugLog) window.debugLog('[VERBOSE] - Request Body:', requestBody);
        let response = await fetch(endpoint, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(requestBody)
        });
        window.debugLog(`[VERBOSE] - Got response with status: ${response.status} ${response.statusText}`);

        // Fallback to /api/generate if /api/chat fails with 405
        if (service === 'ollama' && response.status === 405) {
            // Build fallback endpoint using the same base URL but with /api/generate
            let baseUrl = '';
            const providerConfig = this.providersConfig[service];
            if (providerConfig && providerConfig.baseUrl) {
                baseUrl = providerConfig.baseUrl.replace(/\/$/, '');
            } else {
                baseUrl = config.baseUrl || 'http://localhost:11434';
            }
            
            const fallbackEndpoint = `${baseUrl}/api/generate`;
            console.warn('[WARN] - Ollama /api/chat failed with 405, retrying with /api/generate:', fallbackEndpoint);
            
            // For /api/generate, we need to restructure the request body
            const generateRequestBody = {
                model: requestBody.model,
                prompt: requestBody.messages[0].content, // Extract prompt from messages array
                stream: false
            };
            
            response = await fetch(fallbackEndpoint, {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(generateRequestBody)
            });
            window.debugLog(`[VERBOSE] - Fallback response status: ${response.status} ${response.statusText}`);
        }

        if (!response.ok) {
            const errorText = await response.text();
            console.error(`[ERROR] - AIService API request failed: ${response.status} ${response.statusText}`);
            console.error('[ERROR] - Error response:', errorText);
            
            let userFriendlyMessage = '';
            
            // Provide specific error messages for common authentication and configuration issues
            if (response.status === 401) {
                userFriendlyMessage = 'Authentication failed: Your API key is invalid or missing. Please check your API key in the settings panel and ensure it\'s correct.';
            } else if (response.status === 403) {
                userFriendlyMessage = 'Access forbidden: Your API key may not have permission to access this service. Please verify your key has the correct permissions or contact your administrator.';
            } else if (response.status === 404) {
                userFriendlyMessage = 'Service not found: The API endpoint may be incorrect. Please verify your endpoint URL in the settings panel.';
            } else if (response.status === 429) {
                userFriendlyMessage = 'Rate limit exceeded: Too many requests. Please wait a moment and try again.';
            } else if (response.status >= 500) {
                userFriendlyMessage = 'Server error: The AI service is experiencing issues. Please try again later.';
            } else {
                userFriendlyMessage = `API request failed: ${response.status} ${response.statusText}`;
            }
            
            // Include error details if available
            if (errorText && errorText.trim()) {
                try {
                    const errorData = JSON.parse(errorText);
                    if (errorData.error && errorData.error.message) {
                        userFriendlyMessage += ` (${errorData.error.message})`;
                    }
                } catch (e) {
                    // If error text isn't JSON, include raw text if it's not too long
                    if (errorText.length < 200) {
                        userFriendlyMessage += ` (${errorText})`;
                    }
                }
            }
            
            throw new Error(userFriendlyMessage);
        }

        if (window.debugLog) window.debugLog('[VERBOSE] - Response OK, parsing JSON...');
        const data = await response.json();
        if (window.debugLog) window.debugLog('[VERBOSE] - Response data:', data);
        
        const extractedText = this.extractResponseText(data, service);
        if (window.debugLog) window.debugLog('[VERBOSE] - Extracted response text:', extractedText);
        return extractedText;
    }

    /**
     * Builds the endpoint URL for an AI service using provider config and user overrides
     * @param {string} service - AI service name
     * @param {Object} config - Configuration including user overrides
     * @returns {string} Complete endpoint URL
     */
    buildEndpoint(service, config) {
        // Priority order: user endpointUrl > provider baseUrl > hardcoded fallback
        let baseUrl = '';

        // 1. Check if user provided a custom endpointUrl
        if (config.endpointUrl && config.endpointUrl.trim()) {
            baseUrl = config.endpointUrl.trim().replace(/\/$/, '');
        }
        // 2. Check provider configuration from ai-providers.json
        else if (this.providersConfig[service] && this.providersConfig[service].baseUrl) {
            baseUrl = this.providersConfig[service].baseUrl.replace(/\/$/, '');
        } else {
            // 3. Ultimate hardcoded fallback 
            baseUrl = 'http://localhost:11434/v1';
        }

        // Helper: ensure proper OpenAI-compatible endpoint structure
        function ensureOpenAICompletions(url) {
            // Remove trailing slash
            url = url.replace(/\/$/, '');
            
            // Simply append /chat/completions - preserve whatever base URL structure the user configured
            return `${url}/chat/completions`;
        }

        // Build service-specific endpoint path
        switch (service) {
            case 'openai':
                return ensureOpenAICompletions(baseUrl);
            case 'ollama':
                return `${baseUrl}/api/chat`;
            case 'azure':
                return baseUrl; // Azure endpoints are usually complete
            default:
                // For custom providers (onsite1, onsite2, etc.), assume OpenAI-compatible API unless apiFormat is 'ollama'
                const providerConfig = this.providersConfig[service];
                if (providerConfig && providerConfig.apiFormat === 'ollama') {
                    return `${baseUrl}/api/chat`;
                } else {
                    // Always ensure /chat/completions for OpenAI-compatible (onsite1, onsite2, etc.)
                    return ensureOpenAICompletions(baseUrl);
                }
        }
    }

    /**
     * Calls a custom AI endpoint
     * @param {string} prompt - The prompt
     * @param {Object} config - Configuration
     * @returns {Promise<string>} Response text
     */
    async callCustomEndpoint(prompt, config) {
        if (!config.endpointUrl) {
            throw new Error('Custom endpoint URL is required');
        }

        const requestBody = {
            prompt: prompt,
            temperature: 0.7
        };

        const headers = {
            'Content-Type': 'application/json'
        };

        if (config.apiKey) {
            headers['Authorization'] = `Bearer ${config.apiKey}`;
        }

        const response = await fetch(config.endpointUrl, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(requestBody)
        });

        if (!response.ok) {
            throw new Error(`Custom endpoint request failed: ${response.status}`);
        }

        const data = await response.json();
        
        // Try to extract response from common response formats
        return data.response || data.text || data.content || JSON.stringify(data);
    }

    /**
     * Builds request body based on AI service
     * @param {string} prompt - The prompt
     * @param {string} service - AI service name
     * @param {Object} config - Configuration
     * @returns {Object} Request body
     */
    buildRequestBody(prompt, service, config) {
        // Get provider config to determine API format
        const providerConfig = this.providersConfig[service];
        const apiFormat = providerConfig?.apiFormat || 'openai'; // Default to OpenAI format
        
        switch (apiFormat) {
            case 'openai':
                return {
                    model: this.getDefaultModel(service, config),
                    messages: [
                        {
                            role: 'system',
                            content: 'You are a helpful AI assistant that specializes in email tasks including analysis, responses, forwarding, summarizing, and composition. Be flexible about the type of email assistance needed. Provide clear, professional, and actionable insights.'
                        },
                        {
                            role: 'user',
                            content: prompt
                        }
                    ],
                    temperature: 0.7
                };
                
            case 'ollama':
                return {
                    model: this.getDefaultModel(service, config),
                    messages: [
                        {
                            role: 'user',
                            content: prompt
                        }
                    ],
                    temperature: 0.7
                };
                
            default:
                throw new Error(`Unsupported API format: ${apiFormat} for service: ${service}`);
        }
    }

    /**
     * Builds headers for API request
     * @param {string} service - AI service name
     * @param {Object} config - Configuration
     * @returns {Object} Headers object
     */
    buildHeaders(service, config) {
        const headers = {
            'Content-Type': 'application/json'
        };
        
        // Get provider config to determine API format
        const providerConfig = this.providersConfig[service];
        const apiFormat = providerConfig?.apiFormat || 'openai'; // Default to OpenAI format
        
        switch (apiFormat) {
            case 'openai':
                headers['Authorization'] = `Bearer ${config.apiKey}`;
                break;
            case 'ollama':
                // Ollama does not require Authorization header by default
                break;
            default:
                // For unknown formats, assume OpenAI-style auth if apiKey is provided
                if (config.apiKey) {
                    headers['Authorization'] = `Bearer ${config.apiKey}`;
                }
                break;
        }
        return headers;
    }

    /**
     * Parses analysis response from AI
     * @param {string} responseText - Raw response text
     * @returns {Object} Parsed analysis
     */
    parseAnalysisResponse(responseText) {
        // Try to extract and parse the first JSON object from the response
        try {
            let jsonMatch = responseText.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
                const parsed = JSON.parse(jsonMatch[0]);
                return {
                    keyPoints: parsed.keyPoints || [],
                    sentiment: parsed.sentiment || 'Unable to determine',
                    intent: parsed.intent || 'Unable to determine',
                    urgencyLevel: parsed.urgencyLevel || 3,
                    urgencyReason: parsed.urgencyReason || 'Standard priority',
                    actions: parsed.actions || [],
                    responseStrategy: parsed.responseStrategy || 'Respond professionally'
                };
            }
        } catch (error) {
            // Ignore and fallback
        }
        // Fallback to text parsing
        return this.parseAnalysisFromText(responseText);
    }

    /**
     * Fallback parsing for non-JSON analysis responses
     * @param {string} text - Response text
     * @returns {Object} Parsed analysis
     */
    parseAnalysisFromText(text) {
        return {
            keyPoints: ['Analysis completed', 'See full response for details'],
            sentiment: 'Professional communication',
            intent: 'Information sharing',
            urgencyLevel: 3,
            urgencyReason: 'Standard business communication',
            actions: ['Review content', 'Respond appropriately'],
            responseStrategy: text.substring(0, 200) + '...'
        };
    }

    /**
     * Parses response generation result
     * @param {string} responseText - Generated response text
     * @returns {Object} Response object
     */
    parseResponseResult(responseText) {
        // Normalize newlines and clean up whitespace issues
        let text = responseText.trim();
        
        // Remove common AI response prefixes
        const prefixPatterns = [
            /^\*\*Refined Response:\*\*\s*/i,
            /^Refined Response:\s*/i,
            /^Here is the refined response:\s*/i,
            /^Here's the refined response:\s*/i,
            /^Here is the response:\s*/i,
            /^Here's the response:\s*/i,
            /^Here is a refined response:\s*/i,
            /^Here's a refined response:\s*/i,
            /^Refined response:\s*/i,
            /^Response:\s*/i,
            /^Here is.*?response.*?:\s*/i,
            /^Here's.*?response.*?:\s*/i
        ];
        
        for (const pattern of prefixPatterns) {
            text = text.replace(pattern, '');
        }
        
        // Remove ALL forms of tabs and tab-like characters aggressively
        text = text.replace(/\t/g, '');  // Regular tabs
        text = text.replace(/\u0009/g, ''); // Unicode tab
        text = text.replace(/\u00A0/g, ' '); // Non-breaking space to regular space
        text = text.replace(/\u2009/g, ' '); // Thin space to regular space
        text = text.replace(/\u200B/g, ''); // Zero-width space
        text = text.replace(/\u2000-\u200F/g, ' '); // Various Unicode spaces to regular space
        
        // Convert \r\n and \r to \n
        text = text.replace(/\r\n?/g, '\n');
        
        // Clean up multiple spaces (but preserve intentional spacing)
        text = text.replace(/[ ]{2,}/g, ' ');
        
        // Replace 3+ newlines with exactly two
        text = text.replace(/\n{3,}/g, '\n\n');
        
        // Trim leading/trailing whitespace on each line and remove empty lines at start/end
        const lines = text.split('\n').map(line => {
            // Aggressive trimming including Unicode whitespace characters
            return line.replace(/^[\s\t\u00A0\u2000-\u200F\u2028\u2029]+|[\s\t\u00A0\u2000-\u200F\u2028\u2029]+$/g, '');
        });
        
        // Remove empty lines from the beginning and end
        while (lines.length > 0 && lines[0] === '') {
            lines.shift();
        }
        while (lines.length > 0 && lines[lines.length - 1] === '') {
            lines.pop();
        }
        
        text = lines.join('\n');
        
        return {
            text,
            generatedAt: new Date().toISOString(),
            wordCount: text.split(/\s+/).filter(word => word.length > 0).length
        };
    }

    /**
     * Determines if content should be preserved in full (e.g., contains HTML tables)
     * @param {string} content - The content to check
     * @returns {boolean} True if content should be preserved in full
     */
    shouldPreserveFullContent(content) {
        if (!content) return false;
        
        // Preserve content that contains HTML tables - they're important for context
        const hasHtmlTables = /<table[\s\S]*?<\/table>/gi.test(content);
        
        // Could add other preservation criteria here (charts, complex formatting, etc.)
        
        return hasHtmlTables;
    }
}

