/**
 * AI Service for email analysis and response generation
 * Supports multiple AI providers and models
 */

import { PromptManager } from './PromptManager.js';

export class AIService {
    constructor(providersConfig = null) {
        this.promptManager = new PromptManager();
        
        // Store provider configuration from ai-providers.json
        this.providersConfig = providersConfig || {};
        
        // Email and prompt length management constants
        this.PROMPT_LIMITS = {
            // Conservative limits to account for model context windows and overhead
            MAX_TOTAL_PROMPT_LENGTH: 32000, // characters (~8k tokens)
            MAX_EMAIL_CONTENT_LENGTH: 20000, // characters for email body
            WARNING_EMAIL_LENGTH: 15000, // warn user at this threshold
            TRUNCATION_BUFFER: 2000, // keep this much room for prompt structure
            
            // Smart truncation settings
            PRESERVE_BEGINNING: 2000, // always keep first 2k chars
            PRESERVE_ENDING: 1000, // always keep last 1k chars
            SMART_BREAK_PATTERNS: ['\n\n', '\n', '. ', '! ', '? '] // break on these
        };
        
        // Track truncation events for user transparency
        this.lastTruncationInfo = null;
        
        // Track HTML conversion events for user transparency
        this.lastHtmlConversionInfo = null;
    }

    /**
     * Calculates approximate token count from character count
     * @param {string} text - Text to estimate tokens for
     * @returns {number} Estimated token count
     */
    estimateTokenCount(text) {
        // Rough approximation: 1 token ≈ 4 characters for English text
        return Math.ceil((text || '').length / 4);
    }

    /**
     * Checks if email content needs truncation and returns truncation info
     * @param {string} emailContent - Email body content
     * @param {number} additionalPromptLength - Length of other prompt parts
     * @returns {Object} Truncation analysis result
     */
    analyzeEmailLength(emailContent, additionalPromptLength = 0) {
        const emailLength = (emailContent || '').length;
        const totalEstimatedLength = emailLength + additionalPromptLength;
        
        const analysis = {
            emailLength,
            additionalPromptLength,
            totalEstimatedLength,
            exceedsWarningThreshold: emailLength > this.PROMPT_LIMITS.WARNING_EMAIL_LENGTH,
            requiresTruncation: totalEstimatedLength > this.PROMPT_LIMITS.MAX_TOTAL_PROMPT_LENGTH,
            estimatedTokens: this.estimateTokenCount(emailContent),
            recommendedMaxLength: this.PROMPT_LIMITS.MAX_EMAIL_CONTENT_LENGTH - additionalPromptLength
        };
        

        
        return analysis;
    }

    /**
     * Intelligently truncates email content while preserving important parts
     * @param {string} emailContent - Original email content
     * @param {number} maxLength - Maximum allowed length
     * @returns {Object} Truncation result with truncated content and metadata
     */
    truncateEmailContent(emailContent, maxLength) {
        if (!emailContent || emailContent.length <= maxLength) {
            return {
                content: emailContent,
                wasTruncated: false,
                originalLength: (emailContent || '').length,
                truncatedLength: (emailContent || '').length
            };
        }
        
        const originalLength = emailContent.length;
        const preserveStart = Math.min(this.PROMPT_LIMITS.PRESERVE_BEGINNING, Math.floor(maxLength * 0.6));
        const preserveEnd = Math.min(this.PROMPT_LIMITS.PRESERVE_ENDING, Math.floor(maxLength * 0.3));
        const ellipsisText = '\n\n[... EMAIL CONTENT TRUNCATED FOR PROCESSING ...]\n\n';
        const availableLength = maxLength - preserveStart - preserveEnd - ellipsisText.length;
        
        if (availableLength < 0) {
            // If even basic preservation exceeds limits, just take from the beginning
            const simpleContent = emailContent.substring(0, maxLength - ellipsisText.length) + ellipsisText;
            
            return {
                content: simpleContent,
                wasTruncated: true,
                originalLength,
                truncatedLength: simpleContent.length,
                preservedStart: maxLength - ellipsisText.length,
                preservedEnd: 0
            };
        }
        
        // Find good break points for the start section
        let startContent = emailContent.substring(0, preserveStart);
        for (const pattern of this.PROMPT_LIMITS.SMART_BREAK_PATTERNS) {
            const lastPatternIndex = startContent.lastIndexOf(pattern);
            if (lastPatternIndex > preserveStart * 0.8) {
                startContent = emailContent.substring(0, lastPatternIndex + pattern.length);
                break;
            }
        }
        
        // Find good break points for the end section  
        let endContent = emailContent.substring(emailContent.length - preserveEnd);
        for (const pattern of this.PROMPT_LIMITS.SMART_BREAK_PATTERNS) {
            const firstPatternIndex = endContent.indexOf(pattern);
            if (firstPatternIndex >= 0 && firstPatternIndex < preserveEnd * 0.2) {
                endContent = emailContent.substring(emailContent.length - preserveEnd + firstPatternIndex);
                break;
            }
        }
        
        const truncatedContent = startContent + ellipsisText + endContent;
        
        const result = {
            content: truncatedContent,
            wasTruncated: true,
            originalLength,
            truncatedLength: truncatedContent.length,
            preservedStart: startContent.length,
            preservedEnd: endContent.length,
            charactersRemoved: originalLength - truncatedContent.length + ellipsisText.length
        };
        
        // Store for user notification
        this.lastTruncationInfo = result;
        
        if (window.debugLog) {
            window.debugLog('AIService: Email truncated from', originalLength, 'to', result.truncatedLength, 'characters');
        }
        
        return result;
    }

    /**
     * Detects if email content contains significant HTML markup that should be converted to text
     * @param {string} content - Email content to analyze
     * @returns {Object} HTML analysis results
     */
    analyzeHtmlContent(content) {
        if (!content) {
            return {
                containsHtml: false,
                htmlDensity: 0,
                recommendConversion: false,
                estimatedSavings: 0
            };
        }
        
        // Count actual HTML tags (not URLs in angle brackets)
        // Match tags that start with a letter (HTML tags) or common patterns like </tag>
        const htmlTagMatches = content.match(/<\/?[a-zA-Z][^>]*>/g) || [];
        const htmlTagCount = htmlTagMatches.length;
        const totalLength = content.length;
        
        // Calculate HTML density (percentage of content that is HTML tags)
        const htmlTagLength = htmlTagMatches.join('').length;
        const htmlDensity = totalLength > 0 ? (htmlTagLength / totalLength) * 100 : 0;
        
        // Check for common HTML elements that indicate rich formatting
        const significantHtmlPatterns = [
            /<(div|span|p|table|tr|td|th|ul|ol|li|h[1-6]|strong|b|em|i|a)[^>]*>/gi,
            /style\s*=\s*["|'][^"']*["|']/gi,
            /class\s*=\s*["|'][^"']*["|']/gi,
            /<img[^>]*>/gi,
            /<font[^>]*>/gi
        ];
        
        let significantHtmlCount = 0;
        for (const pattern of significantHtmlPatterns) {
            const matches = content.match(pattern);
            if (matches) significantHtmlCount += matches.length;
        }
        
        // Estimate potential token savings
        const estimatedTextLength = this.convertHtmlToText(content).length;
        const estimatedSavings = Math.max(0, totalLength - estimatedTextLength);
        const savingsPercentage = totalLength > 0 ? (estimatedSavings / totalLength) * 100 : 0;
        
        // Recommend conversion if:
        // 1. HTML density > 15% OR
        // 2. Significant HTML elements > 10 OR  
        // 3. Potential savings > 20% and > 500 characters
        const recommendConversion = (
            htmlDensity > 15 ||
            significantHtmlCount > 10 ||
            (savingsPercentage > 20 && estimatedSavings > 500)
        );
        
        const analysis = {
            containsHtml: htmlTagCount > 0,
            htmlTagCount,
            htmlDensity: Math.round(htmlDensity * 10) / 10,
            significantHtmlCount,
            recommendConversion,
            estimatedSavings,
            savingsPercentage: Math.round(savingsPercentage * 10) / 10,
            originalLength: totalLength,
            estimatedTextLength
        };
        

        
        return analysis;
    }

    /**
     * Converts HTML email content to clean, readable text while preserving structure
     * @param {string} htmlContent - HTML content to convert
     * @returns {string} Clean text version
     */
    convertHtmlToText(htmlContent) {
        if (!htmlContent || typeof htmlContent !== 'string') {
            return htmlContent || '';
        }
        
        // If no HTML tags detected, return as-is
        if (!/<[^>]+>/.test(htmlContent)) {
            return htmlContent;
        }
        
        let text = htmlContent;
        
        // Convert common block elements to line breaks
        const blockElements = [
            'div', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 
            'section', 'article', 'header', 'footer', 'main'
        ];
        
        for (const element of blockElements) {
            text = text.replace(new RegExp(`</${element}[^>]*>`, 'gi'), '\n\n');
            text = text.replace(new RegExp(`<${element}[^>]*>`, 'gi'), '');
        }
        
        // Convert line break elements
        text = text.replace(/<br[^>]*>/gi, '\n');
        text = text.replace(/<hr[^>]*>/gi, '\n---\n');
        
        // Convert list elements
        text = text.replace(/<\/li>/gi, '\n');
        text = text.replace(/<li[^>]*>/gi, '• ');
        text = text.replace(/<\/(ul|ol)>/gi, '\n');
        text = text.replace(/<(ul|ol)[^>]*>/gi, '');
        
        // Convert table elements to structured text
        text = text.replace(/<\/tr>/gi, '\n');
        text = text.replace(/<\/td>/gi, ' | ');
        text = text.replace(/<\/th>/gi, ' | ');
        text = text.replace(/<(table|tbody|thead|tr|td|th)[^>]*>/gi, '');
        
        // Handle emphasis elements
        text = text.replace(/<(strong|b)[^>]*>(.*?)<\/(strong|b)>/gi, '**$2**');
        text = text.replace(/<(em|i)[^>]*>(.*?)<\/(em|i)>/gi, '*$2*');
        
        // Convert links to readable format
        text = text.replace(/<a[^>]*href\s*=\s*["']([^"']*)["'][^>]*>(.*?)<\/a>/gi, '$2 ($1)');
        
        // Remove remaining HTML tags
        text = text.replace(/<[^>]+>/g, '');
        
        // Decode HTML entities
        const htmlEntities = {
            '&amp;': '&',
            '&lt;': '<',
            '&gt;': '>',
            '&quot;': '"',
            '&#39;': "'",
            '&apos;': "'",
            '&nbsp;': ' ',
            '&copy;': '©',
            '&reg;': '®',
            '&trade;': '™',
            '&mdash;': '—',
            '&ndash;': '–',
            '&hellip;': '…'
        };
        
        for (const [entity, char] of Object.entries(htmlEntities)) {
            text = text.replace(new RegExp(entity, 'gi'), char);
        }
        
        // Clean up whitespace
        text = text.replace(/\n{3,}/g, '\n\n'); // Max 2 consecutive newlines
        text = text.replace(/[ \t]{2,}/g, ' '); // Multiple spaces to single space
        text = text.replace(/^\s+|\s+$/gm, ''); // Trim each line
        text = text.trim();
        
        return text;
    }

    /**
     * Processes email content with intelligent HTML conversion if beneficial
     * @param {string} emailContent - Original email content
     * @returns {Object} Processing result with converted content and metadata
     */
    processEmailContent(emailContent) {
        if (!emailContent) {
            return {
                content: emailContent || '',
                wasConverted: false,
                originalLength: 0,
                processedLength: 0
            };
        }
        
        const htmlAnalysis = this.analyzeHtmlContent(emailContent);
        
        if (htmlAnalysis.recommendConversion) {
            const convertedContent = this.convertHtmlToText(emailContent);
            
            const result = {
                content: convertedContent,
                wasConverted: true,
                originalLength: emailContent.length,
                processedLength: convertedContent.length,
                tokensSaved: htmlAnalysis.estimatedSavings,
                conversionReason: this.getConversionReason(htmlAnalysis)
            };
            
            // Store for user notification
            this.lastHtmlConversionInfo = result;
            

            
            return result;
        }
        
        // Clear HTML conversion info if no conversion needed
        this.lastHtmlConversionInfo = null;
        
        return {
            content: emailContent,
            wasConverted: false,
            originalLength: emailContent.length,
            processedLength: emailContent.length,
            tokensSaved: 0
        };
    }

    /**
     * Gets a human-readable reason for HTML conversion
     * @param {Object} htmlAnalysis - HTML analysis results
     * @returns {string} Conversion reason
     */
    getConversionReason(htmlAnalysis) {
        if (htmlAnalysis.htmlDensity > 15) {
            return `High HTML density (${htmlAnalysis.htmlDensity}%)`;
        }
        if (htmlAnalysis.significantHtmlCount > 10) {
            return `Many formatting elements (${htmlAnalysis.significantHtmlCount} found)`;
        }
        if (htmlAnalysis.savingsPercentage > 20) {
            return `Significant space savings (${htmlAnalysis.savingsPercentage}% reduction)`;
        }
        return 'Beneficial for AI processing';
    }

    /**
     * Updates the provider configuration (for dynamic loading)
     * @param {Object} providersConfig - Provider configuration from ai-providers.json
     */
    updateProvidersConfig(providersConfig) {
        this.providersConfig = providersConfig || {};

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
            if (window.debugLog) window.debugLog('Testing connection for service:', config.service);
            
            // Simple ping test with minimal prompt
            const testPrompt = "Hello, respond with 'OK'";
            await this.callAI(testPrompt, config, 'health-check');
            
            if (window.debugLog) window.debugLog('Connection test passed');
            return true;
        } catch (error) {
            console.warn('Connection test failed:', error.message);
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
            console.error('Error fetching Ollama models:', err);
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
            console.error('Error fetching OpenAI-compatible models:', err);
            // Re-throw with original error message if it's already detailed
            throw err;
        }
    }
    
    async analyzeEmail(emailData, config) {
        const prompt = await this.buildAnalysisPrompt(emailData);
        
        try {
            const response = await this.callAI(prompt, config, 'analysis');
            const parsed = this.parseAnalysisResponse(response);
            return parsed;
        } catch (error) {
            console.error('Email analysis failed:', error);
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
        // Ensure analysis is not null - provide default if missing
        if (!analysis) {
            console.warn('Analysis is null, providing default analysis structure');
            analysis = {
                keyPoints: ['No analysis available'],
                sentiment: 'neutral',
                responseStrategy: 'respond professionally'
            };
        }
        
        const prompt = await this.buildResponsePrompt(emailData, analysis, config, config.settingsManager);
        
        try {
            console.log('Prompt length for response generation:', prompt.length, 'characters');
            
            const response = await this.callAI(prompt, config, 'response');
            
            // Validate response before parsing
            if (!response || typeof response !== 'string' || response.trim().length === 0) {
                console.error('Invalid response from AI service:', { 
                    response, 
                    type: typeof response, 
                    length: response ? response.length : 0 
                });
                throw new Error('AI service returned empty or invalid response');
            }
            
            const parsed = this.parseResponseResult(response);
            console.info('Parsed LLM response result:', parsed);
            
            // Validate parsed result
            if (!parsed || !parsed.text || parsed.text.trim().length === 0) {
                console.error('Parsed response is invalid:', parsed);
                throw new Error('Response parsing resulted in empty content');
            }
            
            return parsed;
        } catch (error) {
            console.error('Response generation failed:', error);
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
        if (window.debugLog) window.debugLog('Starting follow-up suggestions generation...');
        if (window.debugLog) window.debugLog('Email data:', emailData);
        if (window.debugLog) window.debugLog('Analysis:', analysis);
        if (window.debugLog) window.debugLog('Config:', config);
        
        // Ensure analysis is not null - provide default if missing
        if (!analysis) {
            console.warn('Analysis is null, providing default analysis structure');
            analysis = {
                keyPoints: ['Sent email content analyzed'],
                sentiment: 'neutral',
                responseStrategy: 'generate appropriate follow-up actions'
            };
        }
        
        const prompt = await this.buildFollowupPrompt(emailData, analysis, config);
        if (window.debugLog) window.debugLog('Built follow-up prompt:', prompt);
        
        try {
            if (window.debugLog) window.debugLog('Calling AI for follow-up suggestions generation...');
            const response = await this.callAI(prompt, config, 'followup');
            if (window.debugLog) window.debugLog('Raw follow-up suggestions result:', response);
            
            const parsed = this.parseFollowupResult(response);
            console.info('Parsed LLM follow-up suggestions result:', parsed);
            return parsed;
        } catch (error) {
            console.error('Follow-up suggestions generation failed:', error);
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
        const prompt = await this.buildRefinementPrompt(currentResponse, instructions, responseSettings);
        
        try {
            const response = await this.callAI(prompt, config, 'refinement');
            return this.parseResponseResult(response);
        } catch (error) {
            console.error('Response refinement failed:', error);
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
        const prompt = await this.buildRefinementPromptWithHistory(
            currentResponse, 
            instructions, 
            responseSettings, 
            originalEmailContext, 
            conversationHistory
        );
        
        try {
            console.log('Prompt length for refinement:', prompt.length, 'characters');
            
            const response = await this.callAI(prompt, config, 'refinement');
            
            // Validate response before parsing
            if (!response || typeof response !== 'string' || response.trim().length === 0) {
                console.error('Invalid refinement response from AI service:', { 
                    response, 
                    type: typeof response, 
                    length: response ? response.length : 0 
                });
                throw new Error('AI service returned empty or invalid refinement response');
            }
            
            const parsed = this.parseResponseResult(response);
            
            // Validate parsed result
            if (!parsed || !parsed.text || parsed.text.trim().length === 0) {
                console.error('Parsed refinement response is invalid:', parsed);
                throw new Error('Refinement response parsing resulted in empty content');
            }
            
            return parsed;
        } catch (error) {
            console.error('Response refinement with history failed:', error);
            throw new Error('Failed to refine response with history: ' + error.message);
        }
    }

    /**
     * Builds the prompt for email analysis
     * @param {Object} emailData - Email data
     * @returns {Promise<string>} Analysis prompt
     */
    async buildAnalysisPrompt(emailData) {
        const dateStr = emailData.date ? new Date(emailData.date).toLocaleString() : 'Compose Mode';
        
        const variables = {
            email_from: emailData.from,
            email_to: emailData.recipients,
            email_subject: emailData.subject,
            email_date: dateStr,
            email_body: emailData.cleanBody || emailData.body,
            email_length: emailData.bodyLength
        };

        return await this.promptManager.buildPrompt('analysis', variables);
    }

    /**
     * Builds the prompt for response generation
     * @param {Object} emailData - Original email data
     * @param {Object} analysis - Email analysis
     * @param {Object} config - Response configuration
     * @returns {Promise<string>} Response generation prompt
     */
    async buildResponsePrompt(emailData, analysis, config, settingsManager = null) {
        
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

        // Prepare template variables
        let variables = {
            // Basic configuration
            isVeryCasualTone: isVeryCasualTone,
            toneDescription: toneMap[config.tone] || 'professional',
            lengthDescription: lengthMap[config.length] || 'medium length',
            
            // Email data
            emailFrom: emailData.from,
            emailSubject: emailData.subject,
            emailDate: emailData.date ? new Date(emailData.date).toLocaleString() : 'Compose Mode',
            
            // Analysis data
            keyPoints: (analysis && analysis.keyPoints) ? analysis.keyPoints.join(', ') : 'Not analyzed',
            sentiment: (analysis && analysis.sentiment) || 'Not analyzed',
            responseStrategy: (analysis && analysis.responseStrategy) || 'Not analyzed',
            
            // Dynamic sections that will be filled below
            writingStyleSection: '',
            emailContent: '',
            htmlConversionNotice: '',
            truncationNotice: '',
            creativeModeSection: '',
            styleReinforcement: ''
        };

        // WritingSamples feature has been deprecated and removed
        
        // Step 1: HTML processing and conversion (before length management)
        const rawEmailContent = emailData.cleanBody || emailData.body || '';
        const htmlProcessingResult = this.processEmailContent(rawEmailContent);
        
        // Step 2: Email length management and smart truncation
        const emailContent = htmlProcessingResult.content;
        const promptSoFar = ''; // Will be calculated with template
        const additionalPromptEstimate = 2000; // estimate for remaining prompt parts
        
        const lengthAnalysis = this.analyzeEmailLength(emailContent, promptSoFar.length + additionalPromptEstimate);
        
        let processedEmailContent = emailContent;
        
        // Add HTML conversion notice if conversion occurred
        if (htmlProcessingResult.wasConverted) {
            const savedKB = Math.round(htmlProcessingResult.tokensSaved / 1024);
            const savingsPercent = Math.round(((htmlProcessingResult.tokensSaved / htmlProcessingResult.originalLength) * 100) * 10) / 10;
            variables.htmlConversionNotice = `**NOTE: HTML email converted to text for better processing** ` +
                `(${savingsPercent}% more efficient, ${savedKB}KB saved)`;
        }
        
        // Check if truncation is needed for either total length OR email content size
        const emailContentTooLarge = emailContent.length > this.PROMPT_LIMITS.MAX_EMAIL_CONTENT_LENGTH;
        const shouldTruncate = lengthAnalysis.requiresTruncation || emailContentTooLarge;
        
        console.log('Truncation decision:', {
            emailContentLength: emailContent.length,
            maxEmailContentLength: this.PROMPT_LIMITS.MAX_EMAIL_CONTENT_LENGTH,
            emailContentTooLarge,
            totalLengthRequiresTruncation: lengthAnalysis.requiresTruncation,
            shouldTruncate
        });
        
        if (shouldTruncate) {
            const maxEmailLength = this.PROMPT_LIMITS.MAX_EMAIL_CONTENT_LENGTH - promptSoFar.length - additionalPromptEstimate;
            
            console.log('Truncation triggered:', {
                emailContentLength: emailContent.length,
                promptSoFarLength: promptSoFar.length,
                additionalPromptEstimate,
                maxEmailLength,
                willTruncate: emailContent.length > maxEmailLength
            });
            
            const truncationResult = this.truncateEmailContent(emailContent, maxEmailLength);
            
            processedEmailContent = truncationResult.content;
            
            if (truncationResult.wasTruncated) {
                // Store truncation info for UI notification
                this.lastTruncationInfo = {
                    wasTruncated: true,
                    originalLength: truncationResult.originalLength,
                    truncatedLength: truncationResult.truncatedLength,
                    preservedStart: truncationResult.preservedStart,
                    preservedEnd: truncationResult.preservedEnd,
                    charactersRemoved: truncationResult.charactersRemoved
                };
                
                variables.truncationNotice = `**NOTE: Email content was automatically shortened for processing** ` +
                    `(${truncationResult.originalLength} → ${truncationResult.truncatedLength} characters)`;
                    
                console.log('Email truncated and notification info stored:', this.lastTruncationInfo);
                    
                if (window.debugLog) {
                    window.debugLog('AIService: Email truncated for processing:', {
                        originalLength: truncationResult.originalLength,
                        truncatedLength: truncationResult.truncatedLength,
                        charactersRemoved: truncationResult.charactersRemoved
                    });
                }
            } else {
                console.log('Truncation NOT triggered despite requirement - investigating:', {
                    emailContentLength: emailContent.length,
                    truncationResult: truncationResult
                });
            }
        } else if (lengthAnalysis.exceedsWarningThreshold) {
            console.log('Email exceeds warning threshold but no truncation required:', {
                emailLength: lengthAnalysis.emailLength,
                warningThreshold: this.PROMPT_LIMITS.WARNING_EMAIL_LENGTH,
                totalEstimated: lengthAnalysis.totalEstimatedLength,
                maxTotal: this.PROMPT_LIMITS.MAX_TOTAL_PROMPT_LENGTH
            });
        } else {
            console.log('No truncation needed:', {
                emailLength: lengthAnalysis.emailLength,
                totalEstimated: lengthAnalysis.totalEstimatedLength,
                maxTotal: this.PROMPT_LIMITS.MAX_TOTAL_PROMPT_LENGTH,
                exceedsWarning: lengthAnalysis.exceedsWarningThreshold
            });
            
            if (lengthAnalysis.exceedsWarningThreshold) {
                // Log warning but don't truncate yet
                if (window.debugLog) {
                    window.debugLog('AIService: Email length exceeds warning threshold:', {
                        length: lengthAnalysis.emailLength,
                        threshold: this.PROMPT_LIMITS.WARNING_EMAIL_LENGTH,
                        totalEstimated: lengthAnalysis.totalEstimatedLength
                    });
                }
            }
        }

        // Set the processed email content
        variables.emailContent = processedEmailContent;
        
        // Set the processed email content
        variables.emailContent = processedEmailContent;

        // Add creativity boost for very casual tone
        if (isVeryCasualTone) {
            variables.creativeModeSection = `**CREATIVE MODE:**
- Feel free to be witty, playful, and engaging
- Use humor and personality as appropriate
- Don't be afraid to be creative with language and approach
- Keep it fun and personable while maintaining respect`;
        }

        // Build the prompt using external template
        const prompt = await this.promptManager.buildPrompt('response', variables, 'default');

        if (window.debugLog) {
            // Add comprehensive prompt length monitoring
            const promptLengthMetrics = {
                totalLength: prompt.length,
                estimatedTokens: this.estimateTokenCount(prompt),
                originalEmailLength: (emailData.cleanBody || emailData.body || '').length,
                processedEmailLength: processedEmailContent ? processedEmailContent.length : 0,
                wasTruncated: this.lastTruncationInfo?.wasTruncated || false,
                exceedsWarning: prompt.length > this.PROMPT_LIMITS.WARNING_EMAIL_LENGTH,
                nearMaxLimit: prompt.length > (this.PROMPT_LIMITS.MAX_TOTAL_PROMPT_LENGTH * 0.8)
            };
            
            window.debugLog('AIService: Prompt length analysis:', promptLengthMetrics);
            
            if (promptLengthMetrics.nearMaxLimit) {
                window.debugLog('AIService: Prompt length is near maximum limit and may cause issues');
            }
        }

        return prompt;
    }

    /**
     * Gets information about the last email truncation that occurred
     * @returns {Object|null} Truncation information or null if no truncation occurred
     */
    getLastTruncationInfo() {
        return this.lastTruncationInfo;
    }

    /**
     * Clears the stored truncation information
     */
    clearTruncationInfo() {
        this.lastTruncationInfo = null;
    }

    /**
     * Gets information about the last HTML conversion that occurred
     * @returns {Object|null} HTML conversion information or null if no conversion occurred
     */
    getLastHtmlConversionInfo() {
        return this.lastHtmlConversionInfo;
    }

    /**
     * Clears the stored HTML conversion information
     */
    clearHtmlConversionInfo() {
        this.lastHtmlConversionInfo = null;
    }

    /**
     * Builds the prompt for follow-up suggestions for sent emails
     * @param {Object} emailData - Sent email data
     * @param {Object} analysis - Email analysis results
     * @param {Object} config - Configuration including AI and response settings
     * @returns {string} Follow-up prompt
     */
    async buildFollowupPrompt(emailData, analysis, config) {
        const lengthMap = {
            1: 'very brief (1-2 suggestions)',
            2: 'brief (2-3 suggestions)',
            3: 'medium (3-4 suggestions)',
            4: 'detailed (4-5 suggestions)',
            5: 'comprehensive (5+ suggestions)'
        };

        // Prepare template variables
        let variables = {
            // Basic configuration
            lengthDescription: lengthMap[config.length] || 'medium',
            
            // Email data
            emailSender: emailData.sender || 'Current User',
            emailTo: emailData.from,
            emailSubject: emailData.subject,
            emailDate: emailData.date ? new Date(emailData.date).toLocaleString() : 'Recently',
            
            // Analysis data
            keyPoints: (analysis && analysis.keyPoints) ? analysis.keyPoints.join(', ') : 'Not analyzed',
            sentiment: (analysis && analysis.sentiment) || 'Not analyzed',
            context: (analysis && analysis.responseStrategy) || 'Not analyzed',
            
            // Dynamic sections that will be filled below
            emailContent: '',
            htmlConversionNotice: '',
            truncationNotice: ''
        };
        
        // Step 1: HTML processing and conversion (before length management) 
        const rawEmailContent = emailData.cleanBody || emailData.body || '';
        const htmlProcessingResult = this.processEmailContent(rawEmailContent);
        
        // Step 2: Email length management and smart truncation for follow-up prompts
        const emailContent = htmlProcessingResult.content;
        const promptSoFar = ''; // Will be calculated with template
        const additionalPromptEstimate = 1500; // estimate for remaining prompt parts (shorter than response prompts)
        
        const lengthAnalysis = this.analyzeEmailLength(emailContent, promptSoFar.length + additionalPromptEstimate);
        
        let processedEmailContent = emailContent;
        
        // Add HTML conversion notice if conversion occurred
        if (htmlProcessingResult.wasConverted) {
            const savedKB = Math.round(htmlProcessingResult.tokensSaved / 1024);
            const savingsPercent = Math.round(((htmlProcessingResult.tokensSaved / htmlProcessingResult.originalLength) * 100) * 10) / 10;
            variables.htmlConversionNotice = `**NOTE: HTML email converted to text for better processing** ` +
                `(${savingsPercent}% more efficient, ${savedKB}KB saved)`;
        }
        
        // Check if truncation is needed for either total length OR email content size
        const emailContentTooLarge = emailContent.length > this.PROMPT_LIMITS.MAX_EMAIL_CONTENT_LENGTH;
        const shouldTruncate = lengthAnalysis.requiresTruncation || emailContentTooLarge;
        
        console.log('Follow-up truncation decision:', {
            emailContentLength: emailContent.length,
            maxEmailContentLength: this.PROMPT_LIMITS.MAX_EMAIL_CONTENT_LENGTH,
            emailContentTooLarge,
            totalLengthRequiresTruncation: lengthAnalysis.requiresTruncation,
            shouldTruncate
        });
        
        if (shouldTruncate) {
            const maxEmailLength = this.PROMPT_LIMITS.MAX_EMAIL_CONTENT_LENGTH - promptSoFar.length - additionalPromptEstimate;
            const truncationResult = this.truncateEmailContent(emailContent, maxEmailLength);
            
            processedEmailContent = truncationResult.content;
            
            if (truncationResult.wasTruncated) {
                variables.truncationNotice = `**NOTE: Email content was automatically shortened for processing** ` +
                    `(${truncationResult.originalLength} → ${truncationResult.truncatedLength} characters)`;
                    
                if (window.debugLog) {
                    window.debugLog('AIService: Email truncated for followup processing:', {
                        originalLength: truncationResult.originalLength,
                        truncatedLength: truncationResult.truncatedLength,
                        charactersRemoved: truncationResult.charactersRemoved
                    });
                }
            }
        }

        // Set the processed email content
        variables.emailContent = processedEmailContent;

        // Build the prompt using external template
        const prompt = await this.promptManager.buildPrompt('followup', variables, 'default');

        return prompt;
    }

    /**
     * Parses follow-up suggestions result
     * @param {string} response - Raw AI response
     * @returns {Object} Parsed follow-up suggestions
     */
    parseFollowupResult(response) {
        if (window.debugLog) window.debugLog('Parsing follow-up suggestions response:', response);
        
        if (!response || typeof response !== 'string') {
            console.warn('Invalid follow-up suggestions response, using fallback');
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

    async buildRefinementPrompt(currentResponse, instructions, responseSettings = null) {
        // Detect creative/humorous requests
        const isCreativeRequest = this.detectCreativeRequest(instructions);
        
        // Prepare template variables
        let variables = {
            // Basic data
            currentResponse: currentResponse.text,
            userInstructions: instructions.trim() || '',
            
            // Settings
            hasSettings: !!responseSettings,
            lengthSetting: '',
            toneSetting: '',
            
            // Dynamic sections
            creativitySection: '',
            htmlTableSection: ''
        };
        
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

            variables.lengthSetting = lengthMap[responseSettings.length] || 'medium length';
            variables.toneSetting = toneMap[responseSettings.tone] || 'professional and courteous';
        }

        // Creative mode adjustments
        if (isCreativeRequest) {
            variables.creativitySection = `**CREATIVE MODE ACTIVATED:**
- Prioritize humor, creativity, and entertainment value over professional constraints
- Feel free to be witty, clever, and amusing
- Use wordplay, jokes, and creative language as appropriate
- Don't worry about being "professional" - focus on being engaging and fun
- Be bold and creative with your language and approach`;
        }

        // Check if user is requesting tables/charts and add HTML formatting instructions
        const requiresHtmlTable = this.detectTableRequest(instructions);
        if (requiresHtmlTable) {
            variables.htmlTableSection = `**IMPORTANT - Table Formatting Instructions:**
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
- Ensure tables are properly formatted and will render well in email clients
- CRITICAL: Use minimal spacing around tables - only TWO line breaks maximum before tables, never more`;
        }

        // Build the prompt using external template
        return await this.promptManager.buildPrompt('refinement', variables, 'simple_prompt');
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
    async buildRefinementPromptWithHistory(currentResponse, instructions, responseSettings = null, originalEmailContext = null, conversationHistory = []) {
        // Prepare template variables
        let variables = {
            // Basic data
            currentResponse: currentResponse.text,
            userInstructions: instructions.trim() || '',
            
            // Settings
            hasSettings: !!responseSettings,
            lengthSetting: '',
            toneSetting: '',
            
            // Context data
            hasOriginalContext: !!originalEmailContext,
            originalEmailFrom: originalEmailContext?.from || '',
            originalEmailSubject: originalEmailContext?.subject || '',
            originalEmailContent: originalEmailContext ? 
                (originalEmailContext.content.substring(0, 500) + (originalEmailContext.content.length > 500 ? '...' : '')) : '',
            
            // Conversation history
            hasConversationHistory: conversationHistory.length > 0,
            conversationSteps: '',
            
            // Dynamic sections
            htmlTableSection: ''
        };
        
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

            variables.lengthSetting = lengthMap[responseSettings.length] || 'medium length';
            variables.toneSetting = toneMap[responseSettings.tone] || 'professional and courteous';
        }

        // Build conversation history steps
        if (conversationHistory.length > 0) {
            let stepsText = '';
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
                
                stepsText += `\nStep ${step.step}: "${step.userInstruction}"\nPrevious Response: ${previousResponseText}\nResult: ${newResponseText}`;
            });
            variables.conversationSteps = stepsText;
        }

        // Check if user is requesting tables/charts or if current response contains tables
        const requiresHtmlTable = this.detectTableRequest(instructions) || 
                                 this.shouldPreserveFullContent(currentResponse.text);
        if (requiresHtmlTable) {
            variables.htmlTableSection = `**IMPORTANT - Table Formatting Instructions:**
- If you include any tables, charts, or structured data, format them using HTML table syntax
- Use proper HTML table elements: <table>, <thead>, <tbody>, <tr>, <th>, <td>
- Apply inline CSS styling to make tables visually appealing:
  - border-collapse: collapse
  - borders around cells: border: 1px solid #ddd
  - header styling: background-color: #f5f5f5; font-weight: bold
  - padding in cells: padding: 8px
  - text alignment as appropriate
- Do NOT use markdown table syntax (| | format) - use only HTML tables
- Ensure tables are properly formatted and will render well in email clients
- CRITICAL: Use minimal spacing around tables - only TWO line breaks maximum before tables, never more`;
        }

        // Build the prompt using external template
        return await this.promptManager.buildPrompt('refinement', variables, 'with_history_prompt');
    }

    /**
     * Makes API call to the specified AI service
     * @param {string} prompt - The prompt to send
     * @param {Object} config - AI configuration
     * @param {string} type - Type of request (analysis, response, refinement)
     * @returns {Promise<string>} AI response text
     */
    async callAI(prompt, config, type) {
        const service = config.service || 'openai';

        if (service === 'custom') {
            return this.callCustomEndpoint(prompt, config);
        }

        // Validate service is configured in providers config
        if (!this.providersConfig[service] && service !== 'custom') {
            console.error(`AIService Service not configured: ${service}`);
            throw new Error(`AI service '${service}' is not configured in providers config`);
        }

        // Build endpoint using provider configuration and user overrides
        let endpoint = this.buildEndpoint(service, config);

        // Get provider config to determine API format
        const providerConfig = this.providersConfig[service];
        const apiFormat = providerConfig?.apiFormat || 'openai';

        let requestBody;
        let headers;

        if (apiFormat === 'ollama') {
            requestBody = {
                model: this.getDefaultModel(service, config),
                messages: [{ role: 'user', content: prompt }],
                stream: false
            };
            headers = { 'Content-Type': 'application/json' };
        } else if (apiFormat === 'bedrock') {
            // Build Bedrock-specific request and handle AWS authentication
            const bedrockResult = await this.callBedrock(prompt, config, endpoint);
            return bedrockResult; // Bedrock handles its own request/response cycle
        } else {
            // For OpenAI, onsite1, onsite2, and other providers, use OpenAI-compatible format
            requestBody = this.buildRequestBody(prompt, service, config);
            headers = this.buildHeaders(service, config);
        }
        let response = await fetch(endpoint, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(requestBody)
        });

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
            console.warn('Ollama /api/chat failed with 405, retrying with /api/generate:', fallbackEndpoint);
            
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
            window.debugLog(`Fallback response status: ${response.status} ${response.statusText}`);
        }

        if (!response.ok) {
            const errorText = await response.text();
            console.error(`AIService API request failed: ${response.status} ${response.statusText}`);
            console.error('Error response:', errorText);
            
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

        const data = await response.json();
        const extractedText = this.extractResponseText(data, service);
        return extractedText;
    }

    /**
     * Extracts response text from different AI service response formats
     * @param {Object} data - Response data from AI service
     * @param {string} service - AI service name for format-specific extraction
     * @returns {string} Extracted response text
     */
    extractResponseText(data, service) {
        
        if (service === 'ollama') {
            // Ollama format: { response: "text" } or { message: { content: "text" } }
            if (data.response) {
                return data.response;
            }
            if (data.message && data.message.content) {
                return data.message.content;
            }
        } else if (service === 'openai' || service === 'onsite1' || service === 'custom') {
            // OpenAI compatible format: { choices: [{ message: { content: "text" } }] }
            if (data.choices && data.choices.length > 0) {
                const choice = data.choices[0];
                if (choice.message && choice.message.content) {
                    return choice.message.content;
                }
                if (choice.text) {
                    return choice.text;
                }
            }
            
            // Fallback for different formats
            if (data.response) {
                return data.response;
            }
            if (data.text) {
                return data.text;
            }
            if (data.content) {
                return data.content;
            }
        }
        
        // Generic fallbacks
        if (data.response) return data.response;
        if (data.text) return data.text;
        if (data.content) return data.content;
        
        // Last resort - stringify the data
        console.warn('Could not extract response text, returning JSON string');
        return JSON.stringify(data);
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
            
            // For Bedrock, ensure no OpenAI-style paths are included
            if (service === 'bedrock1' || service === 'bedrock') {
                baseUrl = baseUrl.replace(/\/chat\/completions$/, '');
            }
        }
        // 2. Check provider configuration from ai-providers.json
        else if (this.providersConfig[service] && this.providersConfig[service].baseUrl) {
            baseUrl = this.providersConfig[service].baseUrl.replace(/\/$/, '');
            
            // For Bedrock, ensure no OpenAI-style paths are included
            if (service === 'bedrock1' || service === 'bedrock') {
                baseUrl = baseUrl.replace(/\/chat\/completions$/, '');
            }
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
            case 'bedrock':
            case 'bedrock1':
                // For Bedrock, ensure clean base URL with no OpenAI artifacts
                baseUrl = baseUrl.replace(/\/chat\/completions$/, '');
                // For Bedrock, we'll build the endpoint dynamically based on model
                return baseUrl; // Base URL only, model-specific path added in request
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
                
            case 'bedrock':
                // Bedrock uses its own request format, but this method shouldn't be called for Bedrock
                // since callAI handles Bedrock separately. This is here for completeness.
                throw new Error(`Bedrock requests should be handled by callBedrock, not buildRequestBody`);
                
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
        
        // Debug logging
        if (window.debugLog) {
            window.debugLog(`buildHeaders for service: ${service}`, {
                'config.apiKey': config.apiKey ? '[HIDDEN]' : '[EMPTY]',
                'config.service': config.service,
                'config.endpointUrl': config.endpointUrl || '[EMPTY]'
            });
        }
        
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
            case 'bedrock':
                // Bedrock uses AWS Signature V4 authentication, handled in callBedrock
                // This method shouldn't be called for Bedrock
                throw new Error(`Bedrock headers should be handled by callBedrock, not buildHeaders`);
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
            responseStrategy: 'Respond professionally with appropriate tone'
        };
    }

    /**
     * Parses response generation result
     * @param {string} responseText - Generated response text
     * @returns {Object} Response object
     */
    parseResponseResult(responseText) {
        // Input validation
        if (!responseText || typeof responseText !== 'string') {
            console.error('parseResponseResult received invalid input:', { responseText, type: typeof responseText });
            return {
                text: 'Response processing error: Invalid input received',
                generatedAt: new Date().toISOString(),
                wordCount: 0
            };
        }
        
        // Normalize newlines and clean up whitespace issues
        let text = responseText.trim();
        
        // If after trimming we have no content, return error response
        if (!text) {
            console.error('parseResponseResult received empty text after trimming');
            return {
                text: 'Response processing error: Empty content received',
                generatedAt: new Date().toISOString(),
                wordCount: 0
            };
        }
        
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
        
        // Final validation before returning
        if (!text || text.trim().length === 0) {
            console.error('parseResponseResult: Final text is empty after processing');
            text = 'Response processing error: Content became empty during processing';
        }
        
        const result = {
            text,
            generatedAt: new Date().toISOString(),
            wordCount: text.split(/\s+/).filter(word => word.length > 0).length
        };
        
        // Log successful parsing for debugging
        console.log('parseResponseResult success:', {
            originalLength: responseText?.length || 0,
            finalLength: result.text.length,
            wordCount: result.wordCount
        });
        
        return result;
    }

    /**
     * AWS Bedrock integration
     */
    async callBedrock(prompt, config, baseEndpoint) {
        window.debugLog('Starting Bedrock AI call');
        window.debugLog('Base endpoint received:', baseEndpoint);
        window.debugLog('Config received:', JSON.stringify(config, null, 2));
        
        try {
            const model = this.getDefaultModel('bedrock', config);
            const region = config.region || 'us-east-1';
            
            // Extract AWS credentials from config.apiKey (base64 encoded format)
            let awsCredentials = {};
            let bearerToken = null;
            
            if (config.apiKey && (config.apiKey.startsWith('ABSK') || config.apiKey.startsWith('BedrockAPIKey'))) {
                try {
                    let credentialsJson = null;
                    
                    if (config.apiKey.startsWith('ABSK')) {
                        // The credential was double-encoded. First decode to get original format
                        try {
                            const originalFormat = atob(config.apiKey);
                            
                            // Now parse the original format: "BedrockAPIKey-ID:BASE64_ENCODED_CREDENTIALS"
                            const colonIndex = originalFormat.indexOf(':');
                            if (colonIndex > 0) {
                                const base64Part = originalFormat.substring(colonIndex + 1);
                                
                                // Set Bearer token to the full credential string for Bedrock authentication
                                const cleanOriginalFormat = originalFormat.replace(/[\u0000-\u001F\u007F]/g, '');
                                bearerToken = cleanOriginalFormat;
                                
                                // Also try decoding it to see if it contains AWS credentials
                                try {
                                    credentialsJson = atob(base64Part);
                                } catch (innerDecodeError) {
                                    // Bearer token approach will be used
                                }
                            }
                        } catch (outerDecodeError) {
                            console.error('Failed to decode credential:', outerDecodeError.message);
                        }
                    } else {
                        // Original BedrockAPIKey format - parse directly for Bearer token
                        const colonIndex = config.apiKey.indexOf(':');
                        if (colonIndex > 0) {
                            bearerToken = config.apiKey;
                        }
                    }
                    
                    if (credentialsJson) {
                        // Check if this is binary format - AWS keys are often stored as binary
                        if (credentialsJson.length === 44 && !credentialsJson.includes(':')) {
                            // Detected binary AWS credentials format (44 bytes)
                            const accessKeyId = 'BedrockAPIKey-5jf0-at-293354421824';
                            const secretAccessKey = bearerToken;
                            
                            awsCredentials = {
                                accessKeyId: accessKeyId,
                                secretAccessKey: secretAccessKey
                            };
                        } else {
                            // Check if it's JSON format or key:secret format
                            try {
                                awsCredentials = JSON.parse(credentialsJson);
                            } catch (jsonError) {
                                // Try key:secret format
                                if (credentialsJson.includes(':')) {
                                    const parts = credentialsJson.split(':');
                                    if (parts.length === 2) {
                                        awsCredentials = {
                                            accessKeyId: parts[0].trim(),
                                            secretAccessKey: parts[1].trim()
                                        };
                                    }
                                }
                            }
                        }
                        
                        if (!(awsCredentials && awsCredentials.accessKeyId && awsCredentials.secretAccessKey)) {
                            console.error('Failed to extract valid credentials from decoded string');
                        }
                    }
                    
                } catch (parseError) {
                    console.error('Failed to parse AWS credentials:', parseError);
                }
            }
            
            // Handle direct format: "AKIA...:secretkey" or "AKIA...:secretkey:sessiontoken"
            if (config.apiKey && !bearerToken && (!awsCredentials.accessKeyId || !awsCredentials.secretAccessKey)) {
                if (config.apiKey.includes(':') && config.apiKey.startsWith('AKIA')) {
                    const parts = config.apiKey.split(':');
                    if (parts.length >= 2 && parts.length <= 3) {
                        awsCredentials = {
                            accessKeyId: parts[0].trim(),
                            secretAccessKey: parts[1].trim(),
                            sessionToken: parts.length === 3 ? parts[2].trim() : undefined
                        };
                        bearerToken = config.apiKey; // Set as bearer token for Lambda
                        window.debugLog('Parsed direct format AWS credentials');
                    }
                }
            }
            
            // Fallback to legacy providerConfigs structure if available
            const legacyAwsConfig = config.providerConfigs?.bedrock || {};
            
            // Use credentials from either source
            const awsConfig = {
                accessKeyId: awsCredentials.accessKeyId || legacyAwsConfig.accessKeyId,
                secretAccessKey: awsCredentials.secretAccessKey || legacyAwsConfig.secretAccessKey,
                sessionToken: awsCredentials.sessionToken || legacyAwsConfig.sessionToken
            };
            
            // Check if we have Bearer token or AWS credentials
            if (!bearerToken && (!awsConfig.accessKeyId || !awsConfig.secretAccessKey)) {
                throw new Error('AWS credentials not configured. Please set Access Key ID and Secret Access Key in settings, or provide a Bearer token.');
            }
            
            window.debugLog('AWS credentials configured successfully');
            
            // For CORS proxy, use the base endpoint directly (model ID goes in request body)
            // For direct AWS Bedrock, build model-specific endpoint
            const endpoint = baseEndpoint.includes('execute-api') ? 
                baseEndpoint : // CORS proxy endpoint
                `${baseEndpoint}/model/${model}/invoke`;
            
            window.debugLog('Full endpoint URL being called:', endpoint);
            
            // Build request body based on model family
            const requestBody = this.buildBedrockRequest(prompt, model);
            
            let headers = {};
            
            if (bearerToken) {
                // Use Bearer token authentication
                headers = {
                    'Authorization': `Bearer ${bearerToken}`,
                    'Content-Type': 'application/json'
                };
            } else {
                // Use AWS Signature V4 authentication
                const signedRequest = await this.createAwsSignature(endpoint, requestBody, {
                    accessKeyId: awsConfig.accessKeyId,
                    secretAccessKey: awsConfig.secretAccessKey,
                    sessionToken: awsConfig.sessionToken,
                    region: region
                }, 'POST');
                headers = signedRequest.headers;
            }
            
            window.debugLog('Request headers being sent:', JSON.stringify(headers, null, 2));
            window.debugLog('Making Bedrock API call to:', endpoint);
            
            const fetchOptions = {
                method: 'POST',
                headers: headers,
                body: JSON.stringify(requestBody),
                mode: 'cors'
            };
            
            // For CORS proxy endpoints, ensure proper CORS handling
            if (baseEndpoint.includes('execute-api')) {
                fetchOptions.credentials = 'omit'; // Don't send credentials for CORS proxy
                window.debugLog('Using CORS proxy mode with credentials omitted');
            }
            
            window.debugLog('About to make fetch request to:', endpoint);
            window.debugLog('Fetch options:', JSON.stringify(fetchOptions, null, 2));
            
            const response = await fetch(endpoint, fetchOptions);
            
            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Bedrock API error ${response.status}: ${errorText}`);
            }
            
            const data = await response.json();
            return this.extractBedrockResponse(data, model);
            
        } catch (error) {
            console.error('Bedrock AI call failed:', error);
            throw new Error(`Bedrock request failed: ${error.message}`);
        }
    }
    
    buildBedrockRequest(prompt, model) {
        const modelFamily = model.split('.')[0];
        
        switch (modelFamily) {
            case 'anthropic':
                return {
                    anthropic_version: "bedrock-2023-05-31",
                    max_tokens: 4000,
                    messages: [
                        {
                            role: "user",
                            content: prompt
                        }
                    ]
                };
                
            case 'amazon':
                return {
                    inputText: prompt,
                    textGenerationConfig: {
                        maxTokenCount: 4000,
                        temperature: 0.7,
                        topP: 0.9
                    }
                };
                
            case 'ai21':
                return {
                    prompt: prompt,
                    maxTokens: 4000,
                    temperature: 0.7
                };
                
            case 'cohere':
                return {
                    prompt: prompt,
                    max_tokens: 4000,
                    temperature: 0.7
                };
                
            default:
                throw new Error(`Unsupported Bedrock model family: ${modelFamily}`);
        }
    }
    
    extractBedrockResponse(data, model) {
        const modelFamily = model.split('.')[0];
        
        switch (modelFamily) {
            case 'anthropic':
                return data.content?.[0]?.text || data.completion || '';
                
            case 'amazon':
                return data.results?.[0]?.outputText || '';
                
            case 'ai21':
                return data.completions?.[0]?.data?.text || '';
                
            case 'cohere':
                return data.generations?.[0]?.text || '';
                
            default:
                console.error('Unknown Bedrock model response format:', data);
                return JSON.stringify(data);
        }
    }
    
    async createAwsSignature(endpoint, body, awsConfig, method = 'POST') {
        const url = new URL(endpoint);
        const { region, accessKeyId, secretAccessKey, sessionToken } = awsConfig;
        const service = url.hostname.includes('bedrock-runtime') ? 'bedrock' : 'bedrock';
        
        const timestamp = new Date().toISOString().replace(/[:\-]|\.\d{3}/g, '');
        const date = timestamp.substring(0, 8);
        
        console.log('AWS signature timestamp:', timestamp);
        console.log('AWS signature date:', date);
        console.log('Timestamp length:', timestamp.length);
        console.log('Timestamp format check:', /^\d{8}T\d{6}Z$/.test(timestamp));
        
        const headers = {
            'X-Amz-Date': timestamp,
            'Host': url.host
        };
        
        // Only add Content-Type for POST requests
        if (method === 'POST') {
            headers['Content-Type'] = 'application/json';
        }
        
        if (sessionToken) {
            headers['X-Amz-Security-Token'] = sessionToken;
        }
        
        // Create canonical request
        const payloadHash = method === 'GET' ? await this.sha256('') : await this.sha256(JSON.stringify(body));
        const canonicalHeaders = Object.keys(headers)
            .sort()
            .map(key => `${key.toLowerCase()}:${headers[key]}`)
            .join('\n');
        const signedHeaders = Object.keys(headers)
            .map(key => key.toLowerCase())
            .sort()
            .join(';');
        
        const canonicalRequest = [
            method,
            url.pathname,
            '',
            canonicalHeaders,
            '',
            signedHeaders,
            payloadHash
        ].join('\n');
        
        // Create string to sign
        const credentialScope = `${date}/${region}/${service}/aws4_request`;
        const canonicalRequestHash = await this.sha256(canonicalRequest);
        const stringToSign = [
            'AWS4-HMAC-SHA256',
            timestamp,
            credentialScope,
            canonicalRequestHash
        ].join('\n');
        
        // Calculate signature
        const signingKey = await this.getSignatureKey(secretAccessKey, date, region, service);
        console.log('Signing key type:', signingKey instanceof Uint8Array ? 'Uint8Array' : typeof signingKey);
        console.log('Signing key length:', signingKey ? signingKey.length : 'null');
        
        const signatureBytes = await this.hmacSha256Binary(signingKey, stringToSign);
        console.log('Signature bytes type:', signatureBytes instanceof Uint8Array ? 'Uint8Array' : typeof signatureBytes);
        console.log('Signature bytes length:', signatureBytes ? signatureBytes.length : 'null');
        
        const signature = Array.from(signatureBytes)
            .map(b => b.toString(16).padStart(2, '0'))
            .join('');
        console.log('Signature conversion completed');
        console.log('Final signature length:', signature.length);
        console.log('Expected signature length: 64 (SHA256 = 32 bytes = 64 hex chars)');
        
        // Create authorization header
        const authHeader = `AWS4-HMAC-SHA256 Credential=${accessKeyId}/${credentialScope}, SignedHeaders=${signedHeaders}, Signature=${signature}`;
        headers['Authorization'] = authHeader;
        
        console.log('Authorization header parts:');
        console.log('AccessKeyId:', accessKeyId);
        console.log('AccessKeyId length:', accessKeyId ? accessKeyId.length : 'null');
        console.log('Credential scope:', credentialScope);
        console.log('Signed headers:', signedHeaders);
        console.log('Signature length:', signature ? signature.length : 'null');
        console.log('Signature (first 20 chars):', signature ? signature.substring(0, 20) : 'null');
        console.log('Authorization header length:', authHeader.length);
        console.log('Authorization header (first 100 chars):', authHeader.substring(0, 100));
        
        // Check for invalid characters in the signature
        if (signature) {
            const validHexPattern = /^[a-f0-9]+$/i;
            const isValidHex = validHexPattern.test(signature);
            console.log('Signature is valid hex:', isValidHex);
            if (!isValidHex) {
                console.error('Signature contains non-hex characters!');
            }
        }

        return { headers };
    }
    
    async sha256(message) {
        const encoder = new TextEncoder();
        const data = encoder.encode(message);
        const hashBuffer = await crypto.subtle.digest('SHA-256', data);
        return Array.from(new Uint8Array(hashBuffer))
            .map(b => b.toString(16).padStart(2, '0'))
            .join('');
    }
    
    async hmacSha256(key, message) {
        const encoder = new TextEncoder();
        const keyData = typeof key === 'string' ? encoder.encode(key) : key;
        const messageData = encoder.encode(message);
        
        const cryptoKey = await crypto.subtle.importKey(
            'raw', keyData, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
        );
        
        const signature = await crypto.subtle.sign('HMAC', cryptoKey, messageData);
        return Array.from(new Uint8Array(signature))
            .map(b => b.toString(16).padStart(2, '0'))
            .join('');
    }
    
    async hmacSha256Binary(key, message) {
        const encoder = new TextEncoder();
        const keyData = typeof key === 'string' ? encoder.encode(key) : key;
        const messageData = encoder.encode(message);
        
        const cryptoKey = await crypto.subtle.importKey(
            'raw', keyData, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
        );
        
        const signature = await crypto.subtle.sign('HMAC', cryptoKey, messageData);
        return new Uint8Array(signature);
    }
    
    async getSignatureKey(key, date, region, service) {
        const kDate = await this.hmacSha256Binary(`AWS4${key}`, date);
        const kRegion = await this.hmacSha256Binary(kDate, region);
        const kService = await this.hmacSha256Binary(kRegion, service);
        return await this.hmacSha256Binary(kService, 'aws4_request');
    }

    /**
     * Fetch available models from Bedrock API
     * @param {Object} config - Configuration with AWS credentials
     * @returns {Promise<Array>} List of available models
     */
    async fetchBedrockModels(config) {
        window.debugLog('Fetching Bedrock models...');
        
        try {
            const region = config.region || 'us-east-1';
            
            // Extract AWS credentials from config.apiKey (base64 encoded format)
            let awsCredentials = {};
            if (config.apiKey && config.apiKey.startsWith('ABSK')) {
                try {
                    // Decode the base64 credentials format: "ABSKBedrockAPIKey-<id>-at-<timestamp>:<base64EncodedJSON>"
                    const parts = config.apiKey.split(':');
                    if (parts.length >= 2) {
                        const credentialsJson = atob(parts[1]);
                        awsCredentials = JSON.parse(credentialsJson);
                        window.debugLog('Successfully parsed AWS credentials for model fetching');
                    }
                } catch (parseError) {
                    console.error('Failed to parse AWS credentials for model fetching:', parseError);
                }
            }
            
            // Fallback to legacy providerConfigs structure if available
            const legacyAwsConfig = config.providerConfigs?.bedrock || {};
            
            // Use credentials from either source
            const awsConfig = {
                accessKeyId: awsCredentials.accessKeyId || legacyAwsConfig.accessKeyId,
                secretAccessKey: awsCredentials.secretAccessKey || legacyAwsConfig.secretAccessKey,
                sessionToken: awsCredentials.sessionToken || legacyAwsConfig.sessionToken
            };
            
            if (!awsConfig.accessKeyId || !awsConfig.secretAccessKey) {
                console.warn('AWS credentials not configured for model fetching');
                // Return static models if no credentials
                return this.getStaticBedrockModels();
            }
            
            // Build foundation models endpoint (different from runtime endpoint)
            const baseUrl = `https://bedrock.${region}.amazonaws.com`;
            const endpoint = `${baseUrl}/foundation-models`;
            
            // Create AWS signature for foundation models API
            const signedRequest = await this.createAwsSignature(endpoint, '', {
                accessKeyId: awsConfig.accessKeyId,
                secretAccessKey: awsConfig.secretAccessKey,
                sessionToken: awsConfig.sessionToken,
                region: region
            }, 'GET');
            
            window.debugLog('Fetching from Bedrock foundation models API:', endpoint);
            
            const response = await fetch(endpoint, {
                method: 'GET',
                headers: signedRequest.headers
            });
            
            if (!response.ok) {
                const errorText = await response.text();
                console.warn(`Bedrock models API error ${response.status}: ${errorText}`);
                // Fallback to static models
                return this.getStaticBedrockModels();
            }
            
            const data = await response.json();
            window.debugLog('Bedrock models response:', data);
            
            // Filter and format models
            const models = data.modelSummaries
                ?.filter(model => {
                    // Filter for text generation models that are ACTIVE
                    return model.modelLifecycle?.status === 'ACTIVE' && 
                           (model.outputModalities?.includes('TEXT') || 
                            model.inferenceTypes?.includes('ON_DEMAND'));
                })
                .map(model => ({
                    id: model.modelId,
                    name: model.modelName || model.modelId,
                    provider: model.providerName || model.modelId.split('.')[0]
                }))
                .sort((a, b) => a.name.localeCompare(b.name)) || [];
            
            window.debugLog(`Found ${models.length} available Bedrock models`);
            return models;
            
        } catch (error) {
            console.error('Error fetching Bedrock models:', error);
            // Fallback to static models
            return this.getStaticBedrockModels();
        }
    }
    
    /**
     * Get static list of common Bedrock models as fallback
     * @returns {Array} Static list of models
     */
    getStaticBedrockModels() {
        return [
            { id: 'anthropic.claude-3-5-sonnet-20241022-v2:0', name: 'Claude 3.5 Sonnet v2', provider: 'Anthropic' },
            { id: 'anthropic.claude-3-sonnet-20240229-v1:0', name: 'Claude 3 Sonnet', provider: 'Anthropic' },
            { id: 'anthropic.claude-3-haiku-20240307-v1:0', name: 'Claude 3 Haiku', provider: 'Anthropic' },
            { id: 'anthropic.claude-instant-v1', name: 'Claude Instant', provider: 'Anthropic' },
            { id: 'amazon.titan-text-express-v1', name: 'Titan Text Express', provider: 'Amazon' },
            { id: 'amazon.titan-text-lite-v1', name: 'Titan Text Lite', provider: 'Amazon' },
            { id: 'ai21.j2-ultra-v1', name: 'Jurassic-2 Ultra', provider: 'AI21 Labs' },
            { id: 'ai21.j2-mid-v1', name: 'Jurassic-2 Mid', provider: 'AI21 Labs' },
            { id: 'cohere.command-text-v14', name: 'Command', provider: 'Cohere' },
            { id: 'cohere.command-light-text-v14', name: 'Command Light', provider: 'Cohere' },
            { id: 'meta.llama2-70b-chat-v1', name: 'Llama 2 70B Chat', provider: 'Meta' },
            { id: 'meta.llama2-13b-chat-v1', name: 'Llama 2 13B Chat', provider: 'Meta' }
        ];
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

