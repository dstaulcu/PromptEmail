/**
 * Prompt Manager Service
 * Handles loading and processing of external prompt templates
 */

export class PromptManager {
    constructor() {
        this.prompts = null;
        this.loadPromise = this.loadPrompts();
    }

    /**
     * Load prompts from external configuration file
     * @returns {Promise<Object>} Loaded prompts configuration
     */
    async loadPrompts() {
        if (this.prompts) {
            return this.prompts;
        }

        try {
            // Load prompts from config file
            const response = await fetch('./config/prompts.json');
            if (!response.ok) {
                throw new Error(`Failed to load prompts: ${response.status} ${response.statusText}`);
            }
            
            this.prompts = await response.json();
            return this.prompts;
        } catch (error) {
            console.error('Failed to load prompt templates:', error);
            // Return fallback prompts if loading fails
            return this.getFallbackPrompts();
        }
    }

    /**
     * Get a prompt template by category and type
     * @param {string} category - Prompt category (analysis, response, followup, etc.)
     * @param {string} type - Prompt type (system_prompt, user_prompt, etc.)
     * @returns {Promise<string>} Prompt template
     */
    async getPrompt(category, type = 'user_prompt') {
        await this.loadPromise;
        
        if (!this.prompts || !this.prompts[category] || !this.prompts[category][type]) {
            console.warn(`Prompt not found: ${category}.${type}, using fallback`);
            return this.getFallbackPrompt(category, type);
        }

        return this.prompts[category][type];
    }

    /**
     * Build a complete prompt with variable substitution
     * @param {string} category - Prompt category
     * @param {Object} variables - Variables for substitution
     * @param {string} promptType - Type of prompt (default: 'user_prompt')
     * @returns {Promise<string>} Complete prompt with variables substituted
     */
    async buildPrompt(category, variables = {}, promptType = 'user_prompt') {
        const template = await this.getPrompt(category, promptType);
        return this.substituteVariables(template, variables);
    }

    /**
     * Get both system and user prompts for a category
     * @param {string} category - Prompt category
     * @param {Object} variables - Variables for substitution
     * @returns {Promise<Object>} Object with system and user prompts
     */
    async getPromptPair(category, variables = {}) {
        const systemPrompt = await this.buildPrompt(category, variables, 'system_prompt');
        const userPrompt = await this.buildPrompt(category, variables, 'user_prompt');
        
        return {
            system: systemPrompt,
            user: userPrompt
        };
    }

    /**
     * Substitute variables in a template string with support for conditional blocks
     * @param {string} template - Template string with {{variable}} placeholders and {{#var}}conditional{{/var}} blocks
     * @param {Object} variables - Variables to substitute
     * @returns {string} Template with variables substituted
     */
    substituteVariables(template, variables = {}) {
        if (!template || typeof template !== 'string') {
            return template;
        }

        let result = template;

        // Add common variables
        const now = new Date();
        const commonVars = {
            current_date: now.toLocaleDateString(),
            current_time: now.toLocaleTimeString(),
            ...variables
        };

        // Handle conditional blocks {{#var}}content{{/var}} - show if var is truthy
        result = result.replace(/\{\{#(\w+)\}\}(.*?)\{\{\/\1\}\}/gs, (match, varName, content) => {
            const value = commonVars[varName];
            return (value && value !== '' && value !== 'false') ? content : '';
        });

        // Handle inverted conditional blocks {{^var}}content{{/var}} - show if var is falsy
        result = result.replace(/\{\{\^(\w+)\}\}(.*?)\{\{\/\1\}\}/gs, (match, varName, content) => {
            const value = commonVars[varName];
            return (!value || value === '' || value === 'false') ? content : '';
        });

        // Replace all {{variable}} placeholders
        for (const [key, value] of Object.entries(commonVars)) {
            if (value !== undefined && value !== null) {
                const placeholder = new RegExp(`\\{\\{${key}\\}\\}`, 'g');
                result = result.replace(placeholder, String(value));
            }
        }

        // Clean up any remaining unreplaced variables (optional)
        result = result.replace(/\{\{[^}]+\}\}/g, '');

        return result;
    }

    /**
     * Format email analysis for prompt inclusion
     * @param {Object} analysis - Email analysis object
     * @returns {string} Formatted analysis summary
     */
    formatAnalysisForPrompt(analysis) {
        if (!analysis) return 'No analysis available';

        const parts = [];
        
        if (analysis.keyPoints && analysis.keyPoints.length > 0) {
            parts.push(`Key Points: ${analysis.keyPoints.join(', ')}`);
        }
        
        if (analysis.intent) {
            parts.push(`Intent: ${analysis.intent}`);
        }
        
        if (analysis.sentiment) {
            parts.push(`Sentiment: ${analysis.sentiment}`);
        }
        
        if (analysis.urgencyLevel) {
            parts.push(`Urgency: ${analysis.urgencyLevel}/5 - ${analysis.urgencyReason || 'No reason provided'}`);
        }
        
        if (analysis.actions && analysis.actions.length > 0) {
            parts.push(`Recommended Actions: ${analysis.actions.join(', ')}`);
        }

        return parts.join('\n');
    }

    /**
     * Format writing samples for prompt inclusion
     * @param {Array} writingSamples - Array of writing samples
     * @returns {string} Formatted writing samples section
     */
    formatWritingSamplesForPrompt(writingSamples) {
        if (!writingSamples || writingSamples.length === 0) {
            return '';
        }

        const formattedSamples = writingSamples
            .map((sample, index) => `Sample ${index + 1}: ${sample}`)
            .join('\n\n');

        return `**Writing Style Examples:**\n${formattedSamples}\n\nPlease match this writing style and tone in your response.`;
    }

    /**
     * Format conversation history for prompt inclusion
     * @param {Array} conversationHistory - Array of conversation steps
     * @returns {string} Formatted conversation history
     */
    formatConversationHistory(conversationHistory) {
        if (!conversationHistory || conversationHistory.length === 0) {
            return 'No previous conversation history';
        }

        return conversationHistory
            .map(step => `Step ${step.step}: User requested "${step.userInstruction}" → AI provided updated response`)
            .join('\n');
    }

    /**
     * Get fallback prompts in case external loading fails
     * @returns {Object} Fallback prompts configuration
     */
    getFallbackPrompts() {
        return {
            analysis: {
                user_prompt: 'Please analyze the following email and provide insights about key points, sentiment, urgency, and recommended actions.'
            },
            response: {
                user_prompt: 'Help me compose an appropriate email response based on the provided context and requirements.'
            },
            followup: {
                user_prompt: 'Suggest follow-up actions for this sent email to ensure effective communication.'
            },
            refinement: {
                simple_prompt: 'Please refine the email content according to the user\'s request while maintaining professionalism.',
                with_history_prompt: 'Refine the email content considering the conversation history and user\'s latest request.'
            },
            chat_assistant: {
                system_prompt: 'You are a helpful AI assistant that specializes in email tasks. Provide clear, professional, and actionable insights.'
            }
        };
    }

    /**
     * Get fallback prompt for specific category and type
     * @param {string} category - Prompt category
     * @param {string} type - Prompt type
     * @returns {string} Fallback prompt
     */
    getFallbackPrompt(category, type) {
        const fallbacks = this.getFallbackPrompts();
        return fallbacks[category]?.[type] || 'Please help with this email task.';
    }

    /**
     * Get available prompt categories
     * @returns {Promise<Array>} Array of available categories
     */
    async getAvailableCategories() {
        await this.loadPromise;
        if (!this.prompts) return [];
        
        return Object.keys(this.prompts).filter(key => !key.startsWith('_'));
    }

    /**
     * Get metadata about prompts configuration
     * @returns {Promise<Object>} Metadata object
     */
    async getMetadata() {
        await this.loadPromise;
        return this.prompts?._metadata || {};
    }
}