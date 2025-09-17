// PromptEmail Taskpane JavaScript
// Main application logic for the email analysis interface

import '../assets/css/taskpane.css';
import { EmailAnalyzer } from '../services/EmailAnalyzer';
import { AIService } from '../services/AIService';
import { ClassificationDetector } from '../services/ClassificationDetector';
import { Logger } from '../services/Logger';
import { SettingsManager } from '../services/SettingsManager';
import { AccessibilityManager } from '../ui/AccessibilityManager';
import { UIController } from '../ui/UIController';

class TaskpaneApp {
    async fetchDefaultProvidersConfig() {
        // Fetch ai-providers.json from public config directory
        try {
            const response = await fetch('/config/ai-providers.json');
            if (!response.ok) throw new Error('Failed to fetch ai-providers.json');
            return await response.json();
        } catch (e) {
            console.warn('[WARN] - Could not load ai-providers.json:', e);
            return {};
        }
    }

    async fetchTaskpaneResourcesConfig() {
        // Fetch taskpane-resources.json from public config directory
        try {
            const response = await fetch('/config/taskpane-resources.json');
            if (!response.ok) throw new Error('Failed to fetch taskpane-resources.json');
            return await response.json();
        } catch (e) {
            console.warn('[WARN] - Could not load taskpane-resources.json:', e);
            return {};
        }
    }

    /**
     * Dynamically populate the resources dropdown from taskpane-resources.json
     */
    populateResourcesDropdown() {
        const dropdownMenu = document.getElementById('help-dropdown-menu');
        if (!dropdownMenu) {
            console.warn('[WARN] - Could not find help-dropdown-menu element');
            return;
        }

        // Always start with fallback content to ensure dropdown is never empty
        dropdownMenu.innerHTML = `
            <a href="#" class="dropdown-item" onclick="window.open('https://github.com/dstaulcu/PromptEmail/wiki', '_blank'); return false;">
                📖 Documentation
            </a>
            <a href="#" class="dropdown-item" onclick="window.open('https://github.com/dstaulcu/PromptEmail/issues', '_blank'); return false;">
                🐛 Issues
            </a>
            <a href="#" class="dropdown-item" onclick="window.open('https://github.com/dstaulcu/PromptEmail', '_blank'); return false;">
                💾 Source Code
            </a>
        `;

        // If we have config with resources, replace with dynamic content
        if (this.taskpaneResourcesConfig?.resources && Array.isArray(this.taskpaneResourcesConfig.resources)) {
            console.log('[DEBUG] - Replacing fallback with dynamic resources:', this.taskpaneResourcesConfig.resources);
            
            // Clear fallback content
            dropdownMenu.innerHTML = '';

            // Create dropdown items from configuration
            this.taskpaneResourcesConfig.resources.forEach((resource, index) => {
                const link = document.createElement('a');
                link.id = `resource-link-${index}`;
                link.href = '#';
                link.className = 'dropdown-item';
                link.target = '_blank';
                link.rel = 'noopener noreferrer';
                link.setAttribute('aria-label', resource.ariaLabel || resource.name);
                link.innerHTML = `${resource.icon} ${resource.name}`;
                
                // Add click handler
                link.addEventListener('click', (event) => this.openResourceLink(event, resource));
                
                dropdownMenu.appendChild(link);
            });

            console.debug('[DEBUG] - Populated resources dropdown with', this.taskpaneResourcesConfig.resources.length, 'dynamic items');
        } else {
            console.log('[DEBUG] - Using fallback content for resources dropdown');
        }
    }

    /**
     * Handle resource link clicks with error handling and fallbacks
     */
    async openResourceLink(event, resource) {
        event.preventDefault();
        
        try {
            window.open(resource.url, '_blank', 'noopener,noreferrer');
        } catch (error) {
            console.error(`[ERROR] - Error opening ${resource.name}:`, error);
            
            // Fallback: copy URL to clipboard if available
            if (navigator.clipboard && navigator.clipboard.writeText) {
                try {
                    await navigator.clipboard.writeText(resource.url);
                    this.showInfoDialog(resource.name, 
                        `Could not open ${resource.name}. The URL has been copied to your clipboard:\n\n${resource.url}`);
                } catch (clipboardError) {
                    console.error('[ERROR] - Could not copy to clipboard:', clipboardError);
                    this.showInfoDialog('Error', 
                        `Could not open ${resource.name}. Please visit:\n\n${resource.url}`);
                }
            } else {
                this.showInfoDialog('Error', 
                    `Could not open ${resource.name}. Please visit:\n\n${resource.url}`);
            }
        }
    }

    /**
     * Determine the AI provider and allowed providers based on user's email domain
     * @param {Object} userProfile - User profile containing email address
     * @returns {Object} {defaultProvider: string, allowedProviders: string[]}
     * @throws {Error} If user email cannot be determined
     */
    getProvidersByDomain(userProfile) {
        if (!this.defaultProvidersConfig || !this.defaultProvidersConfig._config) {
            throw new Error('Provider configuration not loaded');
        }

        const config = this.defaultProvidersConfig._config;
        const emailAddress = userProfile?.emailAddress;

        if (!emailAddress) {
            throw new Error('User email address is required for provider selection but was not available');
        }

        // Extract domain from email address
        const domain = emailAddress.toLowerCase().split('@')[1];
        if (!domain) {
            throw new Error(`Invalid email format: ${emailAddress}`);
        }

        // Check if domain has a specific provider mapping
        if (config.domainBasedProviders && config.domainBasedProviders[domain]) {
            const allowedProviders = config.domainBasedProviders[domain];
            const defaultProvider = allowedProviders[0]; // First item is default
            console.info(`[INFO] - Domain-based provider filtering: ${domain} -> default: ${defaultProvider}, allowed: [${allowedProviders.join(', ')}]`);
            return { defaultProvider, allowedProviders };
        }

        // Use default providers for unmapped domains
        const allowedProviders = config.defaultProviders || ['ollama'];
        const defaultProvider = allowedProviders[0];
        console.warn(`[WARN] - No mapping for domain ${domain}, using default providers: default: ${defaultProvider}, allowed: [${allowedProviders.join(', ')}]`);
        return { defaultProvider, allowedProviders };
    }

    /**
     * Filter dropdown options based on allowed providers for user's domain
     * @param {string[]} allowedProviders - Array of allowed provider keys
     */
    filterProviderDropdown(allowedProviders) {
        const modelServiceSelect = document.getElementById('model-service');
        if (!modelServiceSelect || !allowedProviders) return;

        // Store all original options if not already stored
        if (!this.allProviderOptions) {
            this.allProviderOptions = Array.from(modelServiceSelect.options).map(option => ({
                value: option.value,
                text: option.text
            }));
        }

        // Clear and repopulate with only allowed providers
        modelServiceSelect.innerHTML = '';
        this.allProviderOptions
            .filter(option => allowedProviders.includes(option.value))
            .forEach(option => {
                const optionElement = document.createElement('option');
                optionElement.value = option.value;
                optionElement.text = option.text;
                modelServiceSelect.appendChild(optionElement);
            });

        window.debugLog(`[VERBOSE] - Filtered provider dropdown to: [${allowedProviders.join(', ')}]`);
    }

    /**
     * Apply domain-based provider filtering when user context is available
     */
    async applyDomainBasedProviderFiltering() {
        try {
            // Get user email from current email context - try multiple possible locations
            let userEmail = null;
            let userProfile = null;

            if (this.currentEmail?.context) {
                userEmail = this.currentEmail.context.userEmail || 
                           this.currentEmail.context.userProfile?.emailAddress;
                
                // Create a userProfile object if we have the email
                if (userEmail) {
                    userProfile = {
                        emailAddress: userEmail,
                        ...(this.currentEmail.context.userProfile || {})
                    };
                }
            }
            
            if (!userEmail) {
                console.warn('[WARN] - No user email available in email context, skipping domain filtering');
                return;
            }
            
            const providerConfig = this.getProvidersByDomain(userProfile);
            const { defaultProvider, allowedProviders } = providerConfig;
            
            console.info(`[INFO] - Applying domain-based filtering for ${userEmail}: default=${defaultProvider}, allowed=[${allowedProviders.join(', ')}]`);
            
            // Filter dropdown to only show allowed providers
            this.filterProviderDropdown(allowedProviders);
            
            // Check if current selection needs to be changed
            const modelServiceSelect = document.getElementById('model-service');
            const currentSelection = modelServiceSelect?.value;
            
            window.debugLog(`[VERBOSE] - Current provider selection: '${currentSelection}', domain default: '${defaultProvider}'`);
            
            // Check if user has explicitly chosen a different provider for this domain
            const settings = this.settingsManager.getSettings();
            const domainChoice = settings[`domain-choice-${userEmail.split('@')[1]}`];
            
            window.debugLog(`[VERBOSE] - Domain filtering debug state:`, {
                userEmail,
                currentSelection,
                domainChoice,
                savedModelService: settings['model-service'],
                userActiveChoice: settings['user-active-provider-choice'],
                defaultProvider,
                allowedProviders
            });
            
            if (!currentSelection || !allowedProviders.includes(currentSelection)) {
                // Switch to domain default if:
                // 1. No current selection (first run/new profile), OR
                // 2. Current selection is not allowed for this domain
                const reason = !currentSelection ? 'no provider selected' : `current provider '${currentSelection}' not allowed for domain`;
                console.info(`[INFO] - Setting domain default provider '${defaultProvider}' (${reason})`);
                
                if (modelServiceSelect) {
                    modelServiceSelect.value = defaultProvider;
                    
                    // Save the new selection and mark domain choice
                    settings['model-service'] = defaultProvider;
                    settings[`domain-choice-${userEmail.split('@')[1]}`] = defaultProvider;
                    await this.settingsManager.saveSettings(settings);
                    
                    // Load provider settings for the new provider
                    await this.loadProviderSettings(defaultProvider);
                    
                    // Trigger change event to update related UI
                    modelServiceSelect.dispatchEvent(new Event('change'));
                }
            } else if (!domainChoice && currentSelection !== defaultProvider) {
                // First time seeing this domain and current selection is not the domain default
                // Switch to domain default for proper governance
                console.info(`[INFO] - First time domain filtering for ${userEmail.split('@')[1]}, setting default provider '${defaultProvider}' (governance policy)`);
                
                if (modelServiceSelect) {
                    modelServiceSelect.value = defaultProvider;
                    
                    // Save the new selection and mark domain choice
                    settings['model-service'] = defaultProvider;
                    settings[`domain-choice-${userEmail.split('@')[1]}`] = defaultProvider;
                    await this.settingsManager.saveSettings(settings);
                    
                    // Load provider settings for the new provider
                    await this.loadProviderSettings(defaultProvider);
                    
                    // Trigger change event to update related UI
                    modelServiceSelect.dispatchEvent(new Event('change'));
                }
            } else {
                // User has made a choice for this domain before
                if (domainChoice && allowedProviders.includes(domainChoice)) {
                    // Check if the current selection matches the domain choice
                    if (currentSelection !== domainChoice) {
                        // The current UI selection differs from stored domain choice
                        console.info(`[INFO] - UI shows '${currentSelection}', stored domain choice is '${domainChoice}'`);
                        
                        // If current selection is valid for this domain, respect it and update preferences
                        if (allowedProviders.includes(currentSelection)) {
                            console.info(`[INFO] - Respecting user's current valid selection '${currentSelection}' - updating preferences`);
                            settings['model-service'] = currentSelection;
                            settings[`domain-choice-${userEmail.split('@')[1]}`] = currentSelection;
                            settings['user-active-provider-choice'] = currentSelection;
                            await this.settingsManager.saveSettings(settings);
                            // Load provider settings for the current selection to ensure clean configuration
                            await this.loadProviderSettings(currentSelection);
                            // Also update the UI dropdown to ensure consistency
                            if (modelServiceSelect && modelServiceSelect.value !== currentSelection) {
                                modelServiceSelect.value = currentSelection;
                            }
                            window.debugLog(`[VERBOSE] - Updated all preferences to user's current selection: '${currentSelection}'`);
                        } else {
                            // Current selection is not allowed for this domain - use domain default
                            console.info(`[INFO] - Current selection '${currentSelection}' not allowed for domain, using domain choice '${domainChoice}'`);
                            if (modelServiceSelect) {
                                modelServiceSelect.value = domainChoice;
                                settings['model-service'] = domainChoice;
                                await this.settingsManager.saveSettings(settings);
                                await this.loadProviderSettings(domainChoice);
                                modelServiceSelect.dispatchEvent(new Event('change'));
                            }
                        }
                    } else {
                        // Current selection matches domain choice - all good
                        window.debugLog(`[VERBOSE] - User's current selection '${currentSelection}' matches domain choice '${domainChoice}'`);
                    }
                } else {
                    // No valid domain choice stored, or stored choice is no longer allowed
                    if (allowedProviders.includes(currentSelection)) {
                        // Current selection is valid - make it the new domain choice
                        console.info(`[INFO] - Setting user's current valid selection '${currentSelection}' as new domain choice`);
                        settings['model-service'] = currentSelection;
                        settings[`domain-choice-${userEmail.split('@')[1]}`] = currentSelection;
                        settings['user-active-provider-choice'] = currentSelection;
                        await this.settingsManager.saveSettings(settings);
                        await this.loadProviderSettings(currentSelection);
                    } else {
                        // Current selection is not valid - use domain default
                        console.info(`[INFO] - Current selection '${currentSelection}' not valid, using domain default '${defaultProvider}'`);
                        if (modelServiceSelect) {
                            modelServiceSelect.value = defaultProvider;
                            settings['model-service'] = defaultProvider;
                            settings[`domain-choice-${userEmail.split('@')[1]}`] = defaultProvider;
                            await this.settingsManager.saveSettings(settings);
                            await this.loadProviderSettings(defaultProvider);
                            modelServiceSelect.dispatchEvent(new Event('change'));
                        }
                    }
                }
            }
            
        } catch (error) {
            console.warn('[WARN] - Could not apply domain-based provider filtering:', error);
            // Don't throw - this is not a critical failure
        }
    }

    showAnalysisSection() {
        // Show the analysis section
        const analysisSection = document.getElementById('analysis-section');
        if (analysisSection) {
            analysisSection.classList.remove('hidden');
            if (window.debugLog) window.debugLog('[VERBOSE] - Showing analysis section');
        }
    }
    
    switchToAnalysisTab() {
        // Switch to the analysis tab in the UI
        const analysisTabButton = document.querySelector('.tab-button[aria-controls="panel-analysis"]');
        if (analysisTabButton) {
            if (window.debugLog) window.debugLog('[VERBOSE] - Switching to analysis tab');
            analysisTabButton.click();
        } else {
            console.error('[ERROR] - Analysis tab button not found');
        }
    }

    clearAnalysisAndResponse() {
        // Clear analysis results
        const analysisContainer = document.getElementById('email-analysis');
        if (analysisContainer) {
            analysisContainer.innerHTML = '';
            if (window.debugLog) window.debugLog('[VERBOSE] - Cleared analysis container');
        }

        // Hide analysis section
        const analysisSection = document.getElementById('analysis-section');
        if (analysisSection) {
            analysisSection.classList.add('hidden');
        }

        // Clear chat messages
        const chatMessages = document.getElementById('chat-messages');
        if (chatMessages) {
            chatMessages.innerHTML = '';
        }

        // Clear chat input
        const chatInput = document.getElementById('chat-input');
        if (chatInput) {
            chatInput.value = '';
        }

        // Hide chat section
        const chatSection = document.getElementById('refinement-section');
        if (chatSection) {
            chatSection.classList.add('hidden');
        }

        // Reset internal state
        this.currentAnalysis = null;
        this.currentResponse = null;
        
        // Clear conversation history
        this.clearConversationHistory();
        
        if (window.debugLog) window.debugLog('[VERBOSE] - Analysis and response cleared due to AI provider change');
    }
    showResponseSection() {
        // Show the response section in the UI
        const responseSection = document.getElementById('response-section');
        if (responseSection) {
            responseSection.classList.remove('hidden');
        }
    }
    constructor() {
        // Add unique instance ID for debugging
        this.instanceId = Date.now() + '-' + Math.random().toString(36).substr(2, 9);
        if (window.debugLog) window.debugLog('[VERBOSE] - TaskpaneApp instance created:', this.instanceId);
        
    this.settingsManager = new SettingsManager();
    this.logger = new Logger(this.settingsManager);
    this.emailAnalyzer = new EmailAnalyzer();
    this.aiService = new AIService();
    this.classificationDetector = new ClassificationDetector();
    this.accessibilityManager = new AccessibilityManager();
    this.uiController = new UIController();
        
        // Create a global debug function that other modules can use
        window.debugLog = (message, ...args) => {
            if (this.logger && this.logger.debugEnabled) {
                console.debug(message, ...args);
            }
        };
        
        this.currentEmail = null;
        this.currentAnalysis = null;
        this.currentResponse = null;
        this.sessionStartTime = Date.now();
        
        // Telemetry tracking properties
        this.refinementCount = 0;
        this.hasUsedClipboard = false;
        
        // Conversation history for maintaining context across refinements
        this.conversationHistory = [];
        this.originalEmailContext = null;

        // Model selection UI elements
        this.modelServiceSelect = null;
        this.modelSelectGroup = null;
        this.modelSelect = null;
        
        // Set up session end tracking
        this.setupSessionTracking();
    }

    setupSessionTracking() {
        // Track when user navigates away or closes the taskpane
        window.addEventListener('beforeunload', () => {
            this.logSessionSummary();
        });
        
        // Track when the taskpane loses focus (user switches to another part of Outlook)
        window.addEventListener('blur', () => {
            // Log session summary with a slight delay to allow for quick focus changes
            setTimeout(() => {
                if (!document.hasFocus()) {
                    this.logSessionSummary();
                }
            }, 1000);
        });
    }

    logSessionSummary() {
        if (this.sessionSummaryLogged) return; // Prevent duplicate logging
        this.sessionSummaryLogged = true;
        
        const sessionDuration = Date.now() - this.sessionStartTime;
        this.logger.logEvent('session_summary', {
            session_duration_ms: sessionDuration,
            refinement_count: this.refinementCount,
            clipboard_used: this.hasUsedClipboard,
            email_analyzed: this.currentEmail !== null,
            response_generated: this.currentResponse !== null
        }, 'Information', this.getUserEmailForTelemetry());
    }

    async initialize() {
        try {
            // Initialize Office.js
            await this.initializeOffice();
            
            // Load user settings
            await this.settingsManager.loadSettings();
            
            // Refresh logger debug setting now that settings are loaded
            this.logger.refreshDebugSetting();
            
            // Apply accessibility settings immediately after loading
            const currentSettings = await this.settingsManager.getSettings();
            if (currentSettings['high-contrast']) {
                if (window.debugLog) window.debugLog('[VERBOSE] - Applying high contrast during initialization');
                this.toggleHighContrast(true);
            }
            
            // Load provider config before UI setup
            this.defaultProvidersConfig = await this.fetchDefaultProvidersConfig();
            
            // Load taskpane resources config
            this.taskpaneResourcesConfig = await this.fetchTaskpaneResourcesConfig();
            
            // Update AIService with provider configuration
            this.aiService.updateProvidersConfig(this.defaultProvidersConfig);
            
            // Setup UI
            await this.setupUI();
            
            // Populate resources dropdown after UI is set up
            this.populateResourcesDropdown();
            
            // Update version display
            this.updateVersionDisplay();
            
            // Setup accessibility
            this.accessibilityManager.initialize();
            
            // Initialize Splunk telemetry if enabled
            await this.initializeTelemetry();
            
            // Load current email
            await this.loadCurrentEmail();
            
            // Check if user needs initial setup (first time user or missing API key)
            await this.checkForInitialSetupNeeded();
            
            // Try automatic analysis if conditions are met
            await this.attemptAutoAnalysis();
            
            // Hide loading, show main content
            this.uiController.hideLoading();
            this.uiController.showMainContent();
            
            // Log session start
            this.logger.logEvent('session_start', {
                host: Office.context.host
            });
            
        } catch (error) {
            console.error('[ERROR] - Failed to initialize TaskpaneApp:', error);
            this.uiController.showError('Failed to initialize application. Please try again.');
        }
    }

    async initializeOffice() {
        return new Promise((resolve, reject) => {
            Office.onReady((info) => {
                if (info.host === Office.HostType.Outlook) {
                    // Cache user context immediately when Office is ready
                    this.logger.cacheUserContext();
                    resolve();
                } else {
                    reject(new Error('This add-in is designed for Outlook only'));
                }
            });
        });
    }

    async initializeTelemetry() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Initializing telemetry...');
        
        try {
            // Logger already initialized telemetry config in constructor, just check if it's ready
            // If not initialized yet, wait for it
            if (!this.logger.telemetryConfig) {
                if (window.debugLog) window.debugLog('[VERBOSE] - Waiting for telemetry config to load...');
                // Give it a moment to load
                await new Promise(resolve => setTimeout(resolve, 100));
            }
            
            // Start telemetry auto-flush if enabled
            if (this.logger.telemetryConfig?.telemetry?.enabled) {
                const provider = this.logger.telemetryConfig.telemetry.provider;
                if (provider === 'api_gateway') {
                    this.logger.startApiGatewayAutoFlush();
                    console.info(`[INFO] - ${provider} telemetry enabled and auto-flush started`);
                }
            }
            
        } catch (error) {
            console.error('[ERROR] - Failed to initialize telemetry:', error);
        }
    }

    async setupUI() {
        // Bind event listeners
        this.bindEventListeners();
        
        // Prevent password managers from interfering with API key field
        this.preventPasswordManagerInterference();
        
        // Setup tabs
        this.initializeTabs();
        
        // Model selection UI elements
        this.modelServiceSelect = document.getElementById('model-service');
        this.modelSelectGroup = document.getElementById('model-select-group');
        this.modelSelect = document.getElementById('model-select');
        
        // Populate model service dropdown from defaultProvidersConfig BEFORE loading settings
        // But don't filter by domain yet - that happens in loadSettingsIntoUI
        if (this.modelServiceSelect && this.defaultProvidersConfig) {
            this.modelServiceSelect.innerHTML = Object.entries(this.defaultProvidersConfig)
                .filter(([key, val]) => key !== 'custom' && key !== '_config')
                .map(([key, val]) => `<option value="${key}">${val.label}</option>`)
                .join('');
        }
        
        // Populate settings provider dropdown (not filtered by domain - show all providers)
        const settingsProviderSelect = document.getElementById('settings-provider-select');
        if (settingsProviderSelect && this.defaultProvidersConfig) {
            settingsProviderSelect.innerHTML = Object.entries(this.defaultProvidersConfig)
                .filter(([key, val]) => key !== 'custom' && key !== '_config')
                .map(([key, val]) => `<option value="${key}">${val.label}</option>`)
                .join('');
        }
        
        // Load settings into UI (this will now apply domain filtering and select saved model-service value)
        await this.loadSettingsIntoUI();
        // Hide AI config placeholder in main UI by default
        const aiConfigPlaceholder = document.getElementById('ai-config-placeholder');
        if (aiConfigPlaceholder) {
            aiConfigPlaceholder.classList.add('hidden');
            aiConfigPlaceholder.innerHTML = '';
        }
        if (this.modelServiceSelect && this.modelSelectGroup && this.modelSelect) {
            // Set initial baseUrl to console
            if (this.modelServiceSelect.value && this.defaultProvidersConfig && this.defaultProvidersConfig[this.modelServiceSelect.value]) {
                const baseUrl = this.defaultProvidersConfig[this.modelServiceSelect.value].baseUrl || '';
                window.debugLog(`[VERBOSE] - Model provider: ${this.modelServiceSelect.value}, base URL: ${baseUrl}`);
            }
            await this.updateModelDropdown();
        }
    }

    bindEventListeners() {
        // Main action buttons
        document.getElementById('analyze-email').addEventListener('click', () => this.analyzeEmail());
        document.getElementById('generate-response').addEventListener('click', () => this.generateResponse());
        document.getElementById('start-chat').addEventListener('click', () => this.generateResponse());
        document.getElementById('copy-final-response').addEventListener('click', () => this.copyLatestResponse());
        
        // Chat functionality buttons
        const sendChatBtn = document.getElementById('send-chat-message');
        const copyLatestBtn = document.getElementById('copy-latest-response');
        const clearChatBtn = document.getElementById('clear-chat');
        const chatInput = document.getElementById('chat-input');
        
        if (sendChatBtn) {
            sendChatBtn.addEventListener('click', () => this.sendChatMessage());
        }
        
        if (copyLatestBtn) {
            copyLatestBtn.addEventListener('click', () => this.copyLatestResponse());
        }
        
        if (clearChatBtn) {
            clearChatBtn.addEventListener('click', () => this.clearChatHistory());
        }
        
        if (chatInput) {
            // Send message on Enter (but not Shift+Enter for new lines)
            chatInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    this.sendChatMessage();
                }
            });
        }

        // Response actions (now handled in chat)
        
        // Settings
        document.getElementById('open-settings').addEventListener('click', () => this.openSettings());
        document.getElementById('close-settings').addEventListener('click', () => this.closeSettings());
        document.getElementById('reset-settings').addEventListener('click', () => this.resetSettings());
        
        // Help and navigation links
        document.getElementById('api-key-help-btn').addEventListener('click', () => this.showProviderHelp());
        document.getElementById('test-connection').addEventListener('click', () => this.testConnection());
        document.getElementById('help-dropdown-btn').addEventListener('click', () => this.toggleHelpDropdown());
        // Resource links are now dynamically populated and handled in populateResourcesDropdown()
        
        // Model service change
        document.getElementById('model-service').addEventListener('change', (e) => this.onModelServiceChange(e));
        
        // Settings provider selection change
        const settingsProviderSelect = document.getElementById('settings-provider-select');
        if (settingsProviderSelect) {
            settingsProviderSelect.addEventListener('change', (e) => this.onSettingsProviderChange(e));
        }
        
        // Model selection change (within same provider)
        if (this.modelSelect) {
            this.modelSelect.addEventListener('change', (e) => this.onModelChange(e));
        }
        
        // Settings checkboxes
        document.getElementById('high-contrast').addEventListener('change', (e) => this.toggleHighContrast(e.target.checked));
        document.getElementById('screen-reader-mode').addEventListener('change', (e) => this.toggleScreenReaderMode(e.target.checked));

        // Debug logging checkbox
        const debugCheckbox = document.getElementById('debug-logging');
        if (debugCheckbox) {
            debugCheckbox.addEventListener('change', async (e) => {
                const enabled = e.target.checked;
                const settings = this.settingsManager.getSettings();
                settings['debug-logging'] = enabled;
                await this.settingsManager.saveSettings(settings);
                this.logger.setDebugEnabled(enabled);
                // Update the global debug function
                window.debugLog = (message, ...args) => {
                    if (enabled) {
                        console.debug(message, ...args);
                    }
                };
            });
        }
        
        // Automation settings checkboxes with dependency logic
        const autoAnalysisCheckbox = document.getElementById('auto-analysis');
        const autoResponseCheckbox = document.getElementById('auto-response');
        
        if (autoAnalysisCheckbox) {
            autoAnalysisCheckbox.addEventListener('change', async (e) => {
                const enabled = e.target.checked;
                const settings = this.settingsManager.getSettings();
                
                // Prevent disabling auto-analysis if auto-response is enabled
                if (!enabled && autoResponseCheckbox && autoResponseCheckbox.checked) {
                    // Recheck the checkbox and show a message
                    e.target.checked = true;
                    await this.showInfoDialog('Dependency Required', 'Auto-analysis must be enabled when auto-response is enabled, since responses are generated based on analysis results.');
                    return;
                }
                
                settings['auto-analysis'] = enabled;
                await this.settingsManager.saveSettings(settings);
            });
        }
        
        if (autoResponseCheckbox) {
            autoResponseCheckbox.addEventListener('change', async (e) => {
                const enabled = e.target.checked;
                const settings = this.settingsManager.getSettings();
                
                // Auto-enable analysis when response is enabled
                if (enabled && autoAnalysisCheckbox && !autoAnalysisCheckbox.checked) {
                    autoAnalysisCheckbox.checked = true;
                    settings['auto-analysis'] = true;
                    await this.showInfoDialog('Auto-Analysis Enabled', 'Auto-analysis has been automatically enabled since it\'s required for auto-response generation.');
                }
                
                settings['auto-response'] = enabled;
                await this.settingsManager.saveSettings(settings);
            });
        }
        
        // Auto-save settings with special handling for provider-specific fields
        // (custom-instructions removed - now using interactive chat)
        
        // Special handling for provider-specific fields (API key and endpoint URL)
        ['api-key', 'endpoint-url'].forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                element.addEventListener('blur', () => {
                    // Only save if we're not currently loading provider settings
                    if (!this.isLoadingProviderSettings) {
                        this.saveProviderSettingsContextAware();
                    }
                });
            }
        });

        // Writing samples event listeners
        this.bindWritingSamplesEventListeners();
    }

    preventPasswordManagerInterference() {
        // Add additional attributes to prevent password managers from detecting the API key field
        const apiKeyField = document.getElementById('api-key');
        if (apiKeyField) {
            // Set multiple attributes to discourage password managers
            apiKeyField.setAttribute('autocomplete', 'off');
            apiKeyField.setAttribute('data-lpignore', 'true');
            apiKeyField.setAttribute('data-form-type', 'other');
            apiKeyField.setAttribute('data-1p-ignore', 'true');
            apiKeyField.setAttribute('data-bwignore', 'true');
            apiKeyField.setAttribute('data-dashlane-ignore', 'true');
            apiKeyField.setAttribute('data-keeper-ignore', 'true');
            
            // Also disable form submission to prevent browser password detection
            const form = apiKeyField.closest('form');
            if (form) {
                form.setAttribute('autocomplete', 'off');
                form.addEventListener('submit', (e) => {
                    e.preventDefault();
                    return false;
                });
            }
            
            // Change input type briefly to confuse password managers, then change back
            setTimeout(() => {
                const originalType = apiKeyField.type;
                apiKeyField.type = 'text';
                setTimeout(() => {
                    apiKeyField.type = originalType;
                }, 100);
            }, 500);
        }
    }

    /**
     * Bind event listeners for writing samples management
     */
    bindWritingSamplesEventListeners() {
        // Style analysis toggle
        const styleEnabledCheckbox = document.getElementById('style-analysis-enabled');
        if (styleEnabledCheckbox) {
            styleEnabledCheckbox.addEventListener('change', async (e) => {
                const enabled = e.target.checked;
                await this.settingsManager.setSetting('style-analysis-enabled', enabled);
                this.updateStyleSettingsVisibility(enabled);
                
                if (window.debugLog) {
                    window.debugLog(`[VERBOSE] - Style analysis ${enabled ? 'enabled' : 'disabled'}`);
                }
            });
        }

        // Style strength dropdown
        const styleStrengthSelect = document.getElementById('style-strength');
        if (styleStrengthSelect) {
            styleStrengthSelect.addEventListener('change', async (e) => {
                await this.settingsManager.setSetting('style-strength', e.target.value);
                
                if (window.debugLog) {
                    window.debugLog(`[VERBOSE] - Style strength set to: ${e.target.value}`);
                }
            });
        }

        // Sample input validation
        const sampleTitle = document.getElementById('sample-title');
        const sampleContent = document.getElementById('sample-content');
        const addSampleBtn = document.getElementById('add-sample');

        if (sampleTitle && sampleContent && addSampleBtn) {
            const validateSampleInput = () => {
                const titleValid = sampleTitle.value.trim().length > 0;
                const contentValid = sampleContent.value.trim().length >= 10; // Minimum 10 characters
                addSampleBtn.disabled = !(titleValid && contentValid);
                addSampleBtn.textContent = addSampleBtn.dataset.editId ? 'Update Sample' : 'Add Sample';
            };

            sampleTitle.addEventListener('input', validateSampleInput);
            sampleContent.addEventListener('input', validateSampleInput);
            
            // Add sample button
            addSampleBtn.addEventListener('click', async () => {
                await this.handleSampleSubmission();
            });
        }

        // Cancel edit button
        const cancelEditBtn = document.getElementById('cancel-sample-edit');
        if (cancelEditBtn) {
            cancelEditBtn.addEventListener('click', () => {
                this.cancelSampleEdit();
            });
        }
    }

    /**
     * Update visibility of style settings based on enabled state
     */
    updateStyleSettingsVisibility(enabled) {
        const styleSettings = document.getElementById('style-settings');
        if (styleSettings) {
            if (enabled) {
                styleSettings.classList.remove('disabled');
            } else {
                styleSettings.classList.add('disabled');
            }
        }
    }

    /**
     * Handle sample submission (add or update)
     */
    async handleSampleSubmission() {
        const sampleTitle = document.getElementById('sample-title');
        const sampleContent = document.getElementById('sample-content');
        const addSampleBtn = document.getElementById('add-sample');

        if (!sampleTitle || !sampleContent) return;

        const title = sampleTitle.value.trim();
        const content = sampleContent.value.trim();

        if (!title || content.length < 10) {
            await this.showInfoDialog('Invalid Input', 'Please provide a title and at least 10 characters of content.');
            return;
        }

        try {
            addSampleBtn.disabled = true;
            const editId = addSampleBtn.dataset.editId;

            if (editId) {
                // Update existing sample
                const success = await this.settingsManager.updateWritingSample(
                    parseInt(editId), title, content
                );
                
                if (success) {
                    this.cancelSampleEdit();
                    await this.refreshSamplesList();
                    await this.showInfoDialog('Success', 'Writing sample updated successfully.');
                } else {
                    await this.showInfoDialog('Error', 'Failed to update writing sample. Sample may no longer exist.');
                }
            } else {
                // Add new sample
                await this.settingsManager.addWritingSample(title, content);
                
                // Clear form
                sampleTitle.value = '';
                sampleContent.value = '';
                
                await this.refreshSamplesList();
                await this.showInfoDialog('Success', 'Writing sample added successfully.');
            }
        } catch (error) {
            console.error('[ERROR] - Failed to save writing sample:', error);
            await this.showInfoDialog('Error', 'Failed to save writing sample. Please try again.');
        } finally {
            addSampleBtn.disabled = false;
        }
    }

    /**
     * Cancel sample editing and reset form
     */
    cancelSampleEdit() {
        const sampleTitle = document.getElementById('sample-title');
        const sampleContent = document.getElementById('sample-content');
        const addSampleBtn = document.getElementById('add-sample');
        const cancelEditBtn = document.getElementById('cancel-sample-edit');

        if (sampleTitle) sampleTitle.value = '';
        if (sampleContent) sampleContent.value = '';
        
        if (addSampleBtn) {
            addSampleBtn.textContent = 'Add Sample';
            addSampleBtn.disabled = true;
            delete addSampleBtn.dataset.editId;
        }
        
        if (cancelEditBtn) {
            cancelEditBtn.classList.add('hidden');
        }

        // Remove editing state from all sample items
        document.querySelectorAll('.sample-item.editing').forEach(item => {
            item.classList.remove('editing');
        });
    }

    /**
     * Edit a writing sample
     */
    async editSample(sampleId) {
        const samples = this.settingsManager.getWritingSamples();
        const sample = samples.find(s => s.id === sampleId);
        
        if (!sample) {
            await this.showInfoDialog('Error', 'Sample not found.');
            return;
        }

        // Populate form
        const sampleTitle = document.getElementById('sample-title');
        const sampleContent = document.getElementById('sample-content');
        const addSampleBtn = document.getElementById('add-sample');
        const cancelEditBtn = document.getElementById('cancel-sample-edit');

        if (sampleTitle) sampleTitle.value = sample.title;
        if (sampleContent) sampleContent.value = sample.content;
        
        if (addSampleBtn) {
            addSampleBtn.textContent = 'Update Sample';
            addSampleBtn.disabled = false;
            addSampleBtn.dataset.editId = sampleId.toString();
        }
        
        if (cancelEditBtn) {
            cancelEditBtn.classList.remove('hidden');
        }

        // Highlight the sample being edited
        document.querySelectorAll('.sample-item').forEach(item => {
            item.classList.remove('editing');
        });
        
        const sampleElement = document.querySelector(`[data-sample-id="${sampleId}"]`);
        if (sampleElement) {
            sampleElement.classList.add('editing');
        }

        // Scroll to form
        const sampleInputSection = document.querySelector('.sample-input-section');
        if (sampleInputSection) {
            sampleInputSection.scrollIntoView({ behavior: 'smooth' });
        }
    }

    /**
     * Delete a writing sample
     */
    async deleteSample(sampleId) {
        const samples = this.settingsManager.getWritingSamples();
        const sample = samples.find(s => s.id === sampleId);
        
        if (!sample) {
            await this.showInfoDialog('Error', 'Sample not found.');
            return;
        }

        const confirmed = confirm(`Are you sure you want to delete the sample "${sample.title}"?`);
        if (!confirmed) return;

        try {
            const success = await this.settingsManager.deleteWritingSample(sampleId);
            
            if (success) {
                await this.refreshSamplesList();
                
                // If we were editing this sample, cancel the edit
                const addSampleBtn = document.getElementById('add-sample');
                if (addSampleBtn && addSampleBtn.dataset.editId === sampleId.toString()) {
                    this.cancelSampleEdit();
                }
                
                await this.showInfoDialog('Success', 'Writing sample deleted successfully.');
            } else {
                await this.showInfoDialog('Error', 'Failed to delete writing sample.');
            }
        } catch (error) {
            console.error('[ERROR] - Failed to delete writing sample:', error);
            await this.showInfoDialog('Error', 'Failed to delete writing sample. Please try again.');
        }
    }

    /**
     * Refresh the samples list display
     */
    async refreshSamplesList() {
        const samplesListContainer = document.getElementById('samples-list');
        if (!samplesListContainer) return;

        const samples = this.settingsManager.getWritingSamples();
        
        if (samples.length === 0) {
            samplesListContainer.innerHTML = `
                <div class="samples-empty">
                    <span class="empty-icon" aria-hidden="true">📝</span>
                    No writing samples yet. Add your first sample to help the AI learn your style.
                </div>
            `;
            return;
        }

        samplesListContainer.innerHTML = samples.map(sample => `
            <div class="sample-item" data-sample-id="${sample.id}">
                <div class="sample-header">
                    <h5 class="sample-title">${this.escapeHtml(sample.title)}</h5>
                    <div class="sample-actions-mini">
                        <button type="button" class="sample-btn edit-btn" onclick="window.taskpaneApp.editSample(${sample.id})" title="Edit sample">
                            ✏️
                        </button>
                        <button type="button" class="sample-btn delete-btn" onclick="window.taskpaneApp.deleteSample(${sample.id})" title="Delete sample">
                            🗑️
                        </button>
                    </div>
                </div>
                <div class="sample-meta">
                    <span class="sample-word-count">${sample.wordCount} words</span>
                    <span class="sample-date">${this.formatSampleDate(sample.dateAdded)}</span>
                </div>
                <div class="sample-preview">${this.escapeHtml(sample.content.substring(0, 150))}${sample.content.length > 150 ? '...' : ''}</div>
            </div>
        `).join('');
    }

    /**
     * Format date for sample display
     */
    formatSampleDate(isoString) {
        try {
            const date = new Date(isoString);
            return date.toLocaleDateString();
        } catch {
            return 'Unknown date';
        }
    }

    /**
     * Escape HTML to prevent XSS
     */
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    /**
     * Load writing style settings into UI
     */
    async loadWritingStyleSettings() {
        const styleSettings = this.settingsManager.getStyleSettings();
        
        // Set checkbox states
        const styleEnabledCheckbox = document.getElementById('style-analysis-enabled');
        if (styleEnabledCheckbox) {
            styleEnabledCheckbox.checked = styleSettings.enabled;
            this.updateStyleSettingsVisibility(styleSettings.enabled);
        }

        // Set style strength
        const styleStrengthSelect = document.getElementById('style-strength');
        if (styleStrengthSelect) {
            styleStrengthSelect.value = styleSettings.strength;
        }

        // Load samples list
        await this.refreshSamplesList();

        // Make app available globally for button onclick handlers
        window.taskpaneApp = this;
    }

    /**
     * Update version display with dynamic version from package.json and current environment
     */
    updateVersionDisplay() {
        const versionDisplay = document.getElementById('version-display');
        if (versionDisplay) {
            const version = process.env.PACKAGE_VERSION || '1.0.0';
            const environment = this.detectEnvironment();
            
            // Update text content with lowercase environment names for subtlety
            const shortEnv = environment === 'Local' ? 'local' : environment.toLowerCase();
            versionDisplay.textContent = `v${version} (${shortEnv})`;
            
            // Apply environment-specific CSS class
            versionDisplay.className = 'version'; // Reset classes
            versionDisplay.classList.add(`env-${environment.toLowerCase()}`);
        }
    }

    /**
     * Detect the current environment based on the hostname
     */
    detectEnvironment() {
        const hostname = window.location.hostname;
        
        // Check for S3 bucket hostnames to determine environment
        if (hostname.includes('-dev.s3.') || hostname.includes('dev')) {
            return 'Dev';
        } else if (hostname.includes('-test.s3.') || hostname.includes('test')) {
            return 'Test';
        } else if (hostname.includes('-prod.s3.') || hostname.includes('prod')) {
            return 'Prod';
        } else if (hostname === 'localhost' || hostname === '127.0.0.1') {
            return 'Local';
        } else {
            return 'Unknown';
        }
    }



    getCurrentSettingsSnapshot() {
        return {
            length: '2', // Brief
            tone: '3'    // Professional
        };
    }



    initializeTabs() {
        const tabButtons = document.querySelectorAll('.tab-button');
        const tabPanels = document.querySelectorAll('.tab-panel');
        
        tabButtons.forEach(button => {
            button.addEventListener('click', (e) => {
                const targetPanel = e.target.getAttribute('aria-controls');
                
                // Update buttons
                tabButtons.forEach(btn => {
                    btn.classList.remove('active');
                    btn.setAttribute('aria-selected', 'false');
                });
                
                // Update panels
                tabPanels.forEach(panel => {
                    panel.classList.remove('active');
                });
                
                // Activate current
                e.target.classList.add('active');
                e.target.setAttribute('aria-selected', 'true');
                document.getElementById(targetPanel).classList.add('active');
            });
        });
    }

    async loadCurrentEmail() {
        try {
            this.currentEmail = await this.emailAnalyzer.getCurrentEmail();
            
            // Ensure context is properly stored on currentEmail for later use
            if (this.currentEmail && this.currentEmail.context) {
                if (window.debugLog) window.debugLog('[VERBOSE] - Email context loaded:', this.currentEmail.context);
                
                // Apply domain-based provider filtering now that we have user context
                await this.applyDomainBasedProviderFiltering();
            }
            
            await this.displayEmailSummary(this.currentEmail);
        } catch (error) {
            console.error('[ERROR] - Failed to load current email:', error);
            this.uiController.showError('Failed to load email. Please select an email and try again.');
        }
    }

    async checkForInitialSetupNeeded(showSettingsIfNeeded = true) {
        try {
            const currentSettings = await this.settingsManager.getSettings();
            const selectedService = currentSettings['model-service'] || 'onsite1'; // Default provider
            
            // Check if user has an API key configured for the default/selected provider
            const providerConfigs = currentSettings['provider-configs'] || {};
            const selectedProviderConfig = providerConfigs[selectedService] || {};
            const apiKey = selectedProviderConfig['api-key'] || currentSettings['api-key'] || '';
            
            // Also check if this appears to be a first-time user (no last-updated timestamp)
            const isFirstTime = !currentSettings['last-updated'];
            
            // If no API key is set for the selected provider, or if it's a first-time user
            if (!apiKey.trim() || isFirstTime) {
                if (showSettingsIfNeeded) {
                    console.info('[INFO] - Initial setup needed - showing settings tab');
                    
                    // Show a welcome message for first-time users
                    if (isFirstTime) {
                        this.uiController.showStatus('Welcome! Please configure your AI provider settings to get started.');
                    } else {
                        this.uiController.showStatus('API key required. Please configure your API key in settings.');
                    }
                    
                    // Switch to settings tab
                    const settingsTab = document.querySelector('button[data-tab="settings"]');
                    if (settingsTab) {
                        settingsTab.click();
                        
                        // Highlight the API key field if it exists
                        setTimeout(() => {
                            const apiKeyField = document.getElementById('api-key');
                            if (apiKeyField) {
                                apiKeyField.style.borderColor = '#007bff';
                                apiKeyField.style.borderWidth = '2px';
                                apiKeyField.style.boxShadow = '0 0 0 0.2rem rgba(0, 123, 255, 0.25)';
                                apiKeyField.focus();
                                
                                // Add a helpful tooltip or message
                                let helpDiv = document.getElementById('api-key-help');
                                if (!helpDiv) {
                                    helpDiv = document.createElement('div');
                                    helpDiv.id = 'api-key-help';
                                    helpDiv.style.cssText = 'color: #007bff; font-size: 14px; margin-top: 5px; padding: 8px; background-color: #e7f3ff; border-radius: 4px; border: 1px solid #b8daff;';
                                    helpDiv.innerHTML = '💡 Enter your API key here to start using the AI assistant. You can get your API key from your AI service provider.';
                                    apiKeyField.parentNode.appendChild(helpDiv);
                                }
                                
                                // Remove highlight after user starts typing
                                const removeHighlight = () => {
                                    apiKeyField.style.borderColor = '';
                                    apiKeyField.style.borderWidth = '';
                                    apiKeyField.style.boxShadow = '';
                                    if (helpDiv) {
                                        helpDiv.remove();
                                    }
                                    apiKeyField.removeEventListener('input', removeHighlight);
                                };
                                apiKeyField.addEventListener('input', removeHighlight);
                            }
                        }, 500);
                    }
                    
                    // Log this event for analytics
                    this.logger.logEvent('initial_setup_prompted', {
                        selected_service: selectedService,
                        has_api_key: !!apiKey.trim(),
                        is_first_time: isFirstTime
                    }, 'Information', this.getUserEmailForTelemetry());
                }
                
                return true; // Indicates setup is needed
            }
            
            return false; // No setup needed
        } catch (error) {
            console.error('[ERROR] - Error checking for initial setup:', error);
            return false; // Continue normally on error
        }
    }

    async displayEmailSummary(email) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Displaying email summary:', email);
        
        // Email overview section has been removed for cleaner UI
        
        // Detect classification for console logging only (no storage)
        if (email.body) {
            const classificationResult = this.classificationDetector.detectClassification(email.body);
            if (window.debugLog) window.debugLog('[VERBOSE] - Classification detected (logging only):', classificationResult);
        }

        // Context-aware UI adaptation (works behind the scenes)
        this.adaptUIForContext(email.context);
    }

    /**
     * Adapts the UI based on email context (sent vs inbox vs compose)
     * @param {Object} context - Context information from EmailAnalyzer
     */
    adaptUIForContext(context) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Adapting UI for context:', context);
        
        if (!context) {
            console.warn('[WARN] - No context provided for UI adaptation');
            return;
        }

        // Log detailed context information for debugging
        if (window.debugLog) window.debugLog('[VERBOSE] - Context details:', {
            isSentMail: context.isSentMail,
            isInbox: context.isInbox,
            isCompose: context.isCompose,
            debugInfo: context.debugInfo
        });
        
        // Log telemetry for context detection
        this.logger.logEvent('email_context_detected', {
            context_type: context.isSentMail ? 'sent' : (context.isCompose ? 'compose' : 'inbox'),
            detection_method: context.debugInfo ? context.debugInfo.detectionMethod : 'unknown',
            email_comparison_used: context.debugInfo ? context.debugInfo.emailComparisonUsed : false
        }, 'Information', this.getUserEmailForTelemetry());
        
        try {
            // Get UI elements that need adaptation
            const analysisSection = document.getElementById('panel-analysis');
            const responseSection = document.getElementById('panel-response');
            
            // Find buttons and UI elements for context-aware behavior
            const analyzeButton = document.getElementById('analyze-email');
            const generateResponseButton = document.getElementById('generate-response');
            
            // Apply context-specific adaptations (affects button behavior)
            if (context.isCompose) {
                if (window.debugLog) window.debugLog('[VERBOSE] - Applying compose mode UI adaptations');
                this.adaptUIForComposeMode();
            } else if (context.isSentMail) {
                if (window.debugLog) window.debugLog('[VERBOSE] - Applying sent mail UI adaptations');
                this.adaptUIForSentMail();
            } else {
                if (window.debugLog) window.debugLog('[VERBOSE] - Applying inbox mail UI adaptations');
                this.adaptUIForInboxMail();
            }

        } catch (error) {
            console.error('[ERROR] - Error adapting UI for context:', error);
        }
    }

    /**
     * Adapts UI for compose mode (writing new email)
     */
    adaptUIForComposeMode() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Adapting UI for compose mode');
        
        // Hide analysis features since we're composing
        this.setElementVisibility('analyze-email', false);
        this.setElementVisibility('panel-analysis', false);
        
        // Show writing assistance features
        this.setButtonText('generate-response', '✍️ Writing Assistant');
        this.setElementVisibility('generate-response', true);
        
        // Update tab labels if they exist
        this.setElementText('tab-analysis', '📝 Composition');
        this.setElementText('tab-response', '✍️ Writing Help');
    }

    /**
     * Adapts UI for sent mail (viewing previously sent emails)
     */
    adaptUIForSentMail() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Adapting UI for sent mail');
        
        // Show analysis with different focus
        this.setButtonText('analyze-email', '📋 Analyze Sent Message');
        this.setElementVisibility('analyze-email', true);
        
        // Change response generation to follow-up suggestions
        this.setButtonText('generate-response', '📅 Follow-up Suggestions');
        this.setElementVisibility('generate-response', true);
        
        // Update tab labels
        this.setElementText('tab-analysis', '📋 Sent Analysis');
        this.setElementText('tab-response', '📅 Follow-up');
    }

    /**
     * Adapts UI for inbox mail (received emails)
     */
    adaptUIForInboxMail() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Adapting UI for inbox mail (received)');
        
        // Standard inbox functionality
        this.setButtonText('analyze-email', '🔍 Analyze Email');
        this.setElementVisibility('analyze-email', true);
        
        this.setButtonText('generate-response', '💬 Start Chat Assistant');
        this.setElementVisibility('generate-response', true);
        
        // Standard tab labels
        this.setElementText('tab-analysis', '🔍 Analysis');
        this.setElementText('tab-response', '✉️ Response');
    }

    /**
     * Gets a human-readable context label
     * @param {Object} context - Context information
     * @returns {string} Context label
     */
    getContextLabel(context) {
        if (context.isCompose) return '📝 COMPOSING';
        if (context.isSentMail) return '📤 SENT MAIL';
        if (context.isInbox) return '📥 INBOX';
        return '📧 EMAIL';
    }

    /**
     * Get CSS class for context display styling
     * @param {Object} context - Email context object
     * @returns {string} CSS class name
     */
    getContextClass(context) {
        if (context.isCompose) return 'context-compose';
        if (context.isSentMail) return 'context-sent';
        if (context.isInbox) return 'context-inbox';
        return 'context-inbox'; // default
    }

    /**
     * Helper method to set element visibility
     * @param {string} elementId - Element ID
     * @param {boolean} visible - Whether element should be visible
     */
    setElementVisibility(elementId, visible) {
        const element = document.getElementById(elementId);
        if (element) {
            element.style.display = visible ? '' : 'none';
        }
    }

    /**
     * Helper method to set button text
     * @param {string} elementId - Button element ID
     * @param {string} text - New button text
     */
    setButtonText(elementId, text) {
        const element = document.getElementById(elementId);
        if (element) {
            element.textContent = text;
        }
    }

    /**
     * Helper method to set element text content
     * @param {string} elementId - Element ID
     * @param {string} text - New text content
     */
    setElementText(elementId, text) {
        const element = document.getElementById(elementId);
        if (element) {
            element.textContent = text;
        }
    }

    async attemptAutoAnalysis() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Checking if automatic analysis should be performed...');
        
        // Check if auto-analysis is enabled in settings
        const settings = this.settingsManager.getSettings();
        const autoAnalysisEnabled = settings['auto-analysis'] || false;
        
        if (!autoAnalysisEnabled) {
            if (window.debugLog) window.debugLog('[VERBOSE] - Auto-analysis disabled in settings, skipping');
            return;
        }
        
        // Only auto-analyze if we have an email
        if (!this.currentEmail) {
            if (window.debugLog) window.debugLog('[VERBOSE] - No email available for auto-analysis');
            return;
        }

        // Skip auto-analysis if user is in initial setup mode (no API key configured)
        const needsSetup = await this.checkForInitialSetupNeeded(false);
        if (needsSetup) {
            if (window.debugLog) window.debugLog('[VERBOSE] - Initial setup needed, skipping auto-analysis');
            return;
        }

        try {
            // Get current AI provider settings
            const currentSettings = await this.settingsManager.getSettings();
            if (window.debugLog) window.debugLog('[VERBOSE] Auto-analysis settings check:', {
                fullSettings: currentSettings,
                modelService: currentSettings['model-service'],
                modelServiceType: typeof currentSettings['model-service'],
                modelServiceLength: currentSettings['model-service']?.length
            });
            
            const selectedService = currentSettings['model-service'];
            
            // Also check what the UI element shows
            const modelServiceElement = document.getElementById('model-service');
            if (window.debugLog) window.debugLog('[VERBOSE] UI element check:', {
                elementExists: !!modelServiceElement,
                elementValue: modelServiceElement?.value,
                elementType: typeof modelServiceElement?.value,
                optionsCount: modelServiceElement?.options?.length,
                selectedIndex: modelServiceElement?.selectedIndex,
                allOptions: modelServiceElement ? Array.from(modelServiceElement.options).map(opt => ({value: opt.value, text: opt.text, selected: opt.selected})) : 'N/A'
            });

            // Check for UI/settings mismatch and log it
            if (modelServiceElement && modelServiceElement.value !== selectedService) {
                console.warn(`[WARN] - UI/Settings mismatch detected: UI shows '${modelServiceElement.value}', settings show '${selectedService}'. This may cause auto-analysis to fail.`);
                window.debugLog(`[VERBOSE] - Syncing UI dropdown to match saved settings: ${selectedService}`);
                modelServiceElement.value = selectedService;
                // Trigger change event to update related UI
                modelServiceElement.dispatchEvent(new Event('change'));
            }
            
            if (!selectedService) {
                console.warn('[WARN] - No AI service configured, skipping auto-analysis');
                return;
            }

            // Check for classification (console logging only)
            const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
            if (window.debugLog) window.debugLog('[VERBOSE] - Email classification for auto-analysis:', classification);
            
            // Always proceed with auto-analysis regardless of classification
            // Classification is only logged for reference

            // Test AI service health
            const config = this.getAIConfiguration();
            const isHealthy = await this.aiService.testConnection(config);
            
            if (!isHealthy) {
                if (window.debugLog) window.debugLog('[VERBOSE] - AI service not healthy, skipping auto-analysis');
                return;
            }

            console.info('[INFO] - Connection to AI service is healthy, performing automatic analysis...');
            await this.performAnalysisWithResponse();
            
        } catch (error) {
            console.error('[ERROR] - Error during auto-analysis check:', error);
            // Don't show error to user, just skip auto-analysis
        }
    }

    async performAnalysisWithResponse() {
        const analysisStartTime = Date.now();
        let analysisEndTime, responseStartTime, responseEndTime;
        
        try {
            this.uiController.showStatus('Auto-analyzing email...');
            
            // Get AI configuration
            const config = this.getAIConfiguration();
            
            // Perform analysis
            this.currentAnalysis = await this.aiService.analyzeEmail(this.currentEmail, config);
            analysisEndTime = Date.now();
            
            // Display results
            this.displayAnalysis(this.currentAnalysis);
            this.updateWorkflowStep(3); // Show chat assistant step
            
            // Check if auto-response generation is enabled
            const settings = this.settingsManager.getSettings();
            const autoResponseEnabled = settings['auto-response'] || false;
            
            if (!autoResponseEnabled) {
                // Only analysis was performed, log and exit
                this.logger.logEvent('auto_analysis_completed', {
                    model_service: config.service,
                    model_name: config.model,
                    email_length: this.currentEmail.bodyLength,
                    auto_response_generated: false,
                    analysis_duration_ms: analysisEndTime - analysisStartTime
                }, 'Information', this.getUserEmailForTelemetry());
                
                this.uiController.showStatus('Email analyzed automatically. Click "Start Chat Assistant" to generate a response and begin refining.');
                return;
            }
            
            // Auto-generate response as well (if enabled in settings)
            console.info('[INFO] - Auto-generating response after analysis...');
            responseStartTime = Date.now();
            const responseConfig = this.getResponseConfiguration();
            
            // Check email context to determine response type
            const emailContext = this.currentEmail.context || { isSentMail: false };
            
            if (emailContext.isSentMail) {
                // Generate follow-up suggestions for sent mail
                console.info('[INFO] - Generating follow-up suggestions for sent mail...');
                this.currentResponse = await this.aiService.generateFollowupSuggestions(
                    this.currentEmail, 
                    this.currentAnalysis, 
                    { ...config, ...responseConfig }
                );
            } else {
                // Generate response for received mail
                console.info('[INFO] - Generating response for received mail...');
                this.currentResponse = await this.aiService.generateResponse(
                    this.currentEmail, 
                    this.currentAnalysis, 
                    { ...config, ...responseConfig }
                );
            }
            responseEndTime = Date.now();
            
            // Initialize conversation history for this email and response
            this.initializeConversationHistory(this.currentEmail, this.currentAnalysis);
            
            // Response will be shown in chat interface
            
            // Show analysis section for convenience
            this.showAnalysisSection();
            
            // Follow the same workflow as manual response generation
            this.showChatSection();
            this.updateWorkflowStep(4); // Move directly to step 4 (chat active) since response is auto-generated
            
            // Initialize chat with the auto-generated response
            this.initializeChatWithResponse();
            
            // Log successful auto-analysis and response generation with flattened performance metrics
            this.logger.logEvent('auto_analysis_completed', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                auto_response_generated: true,
                email_context: this.currentEmail.context ? (this.currentEmail.context.isSentMail ? 'sent' : 'inbox') : 'unknown',
                generation_type: 'standard_response',
                refinement_count: this.refinementCount,
                clipboard_used: this.hasUsedClipboard,
                // Flattened performance metrics
                analysis_duration_ms: analysisEndTime - analysisStartTime,
                response_generation_duration_ms: responseEndTime - responseStartTime,
                total_duration_ms: responseEndTime - analysisStartTime
            }, 'Information', this.getUserEmailForTelemetry());
            
            this.uiController.showStatus('Email analyzed and draft response generated automatically.');
            
        } catch (error) {
            console.error('[ERROR] - Auto-analysis failed:', error);
            this.uiController.showStatus('Automatic analysis failed. You can still analyze manually.');
        }
    }

    async analyzeEmail() {
        if (!this.currentEmail) {
            this.uiController.showError('No email selected. Please select an email first.');
            return;
        }

        // Clear any previous refinement instructions (they should be ephemeral)
        const refinementField = document.getElementById('refinement-instructions');
        if (refinementField) {
            refinementField.value = '';
        }

        // Check for classification (console logging only)
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        if (window.debugLog) window.debugLog('[VERBOSE] - Email classification check:', classification);
        
        // Always proceed with analysis - no restrictions based on classification
        await this.performAnalysis();
    }

    async performAnalysis() {
        const analysisStartTime = Date.now();
        
        try {
            this.uiController.showStatus('Analyzing email...');
            this.uiController.setButtonLoading('analyze-email', true);
            
            // Get AI configuration
            const config = this.getAIConfiguration();
            
            // Perform analysis
            this.currentAnalysis = await this.aiService.analyzeEmail(this.currentEmail, config);
            const analysisEndTime = Date.now();
            
            // Display results
            this.displayAnalysis(this.currentAnalysis);
            
            // Log successful analysis with flattened performance telemetry
            this.logger.logEvent('email_analyzed', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                recipients_count: this.currentEmail.recipients.split(',').length,
                analysis_success: true,
                refinement_count: this.refinementCount,
                clipboard_used: this.hasUsedClipboard,
                // Flattened performance metrics
                analysis_duration_ms: analysisEndTime - analysisStartTime
            }, 'Information', this.getUserEmailForTelemetry());
            
            this.uiController.showStatus('Email analysis completed successfully.');
            
            // Update workflow to show next step
            this.updateWorkflowStep(3);
            
        } catch (error) {
            console.error('[ERROR] - Analysis failed:', error);
            
            // Provide more specific error messages based on error type
            let userMessage = 'Analysis failed. Please check your configuration and try again.';
            let showSettings = false;
            
            if (error.message && error.message.includes('Authentication failed')) {
                userMessage = 'Analysis failed: Invalid or missing API key. Please check your API key in the settings panel.';
                showSettings = true;
            } else if (error.message && error.message.includes('Access forbidden')) {
                userMessage = 'Analysis failed: API key permissions issue. Please verify your key has the correct permissions.';
                showSettings = true;
            } else if (error.message && error.message.includes('Service not found')) {
                userMessage = 'Analysis failed: Service endpoint not found. Please verify your endpoint URL in settings.';
                showSettings = true;
            } else if (error.message && error.message.includes('Rate limit exceeded')) {
                userMessage = 'Analysis failed: Rate limit exceeded. Please wait a moment and try again.';
            }
            
            // Show the error with additional context
            this.uiController.showError(userMessage);
            
            // If it's a configuration issue, also provide a way to access settings
            if (showSettings) {
                // Switch to settings tab to help user fix the issue
                setTimeout(() => {
                    const settingsTab = document.querySelector('button[data-tab="settings"]');
                    if (settingsTab) {
                        settingsTab.click();
                    }
                }, 2000);
            }
        } finally {
            this.uiController.setButtonLoading('analyze-email', false);
        }
    }

    async generateResponse() {
        if (!this.currentEmail) {
            this.uiController.showError('No email to respond to. Please analyze an email first.');
            return;
        }

        // Check if this is sent mail context - handle differently
        if (this.currentEmail.context && this.currentEmail.context.isSentMail) {
            await this.generateFollowupSuggestions();
            return;
        }

        // Detect email classification for logging purposes only
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        if (window.debugLog) window.debugLog('[VERBOSE] - Email classification detected for response generation:', classification);

        try {
            this.uiController.showStatus('Starting chat assistant...');
            this.uiController.setButtonLoading('generate-response', true);
            
            // Get configuration
            const config = this.getAIConfiguration();
            const responseConfig = this.getResponseConfiguration();
            
            // Ensure we have analysis data - if not, run analysis first
            let analysisData = this.currentAnalysis;
            if (!analysisData) {
                console.warn('[WARN] - No current analysis available, running analysis first');
                this.uiController.showStatus('Analyzing email before generating response...');
                
                try {
                    // Run analysis first
                    await this.performAnalysis();
                    analysisData = this.currentAnalysis;
                    
                    if (!analysisData) {
                        // If analysis still failed, create minimal default
                        console.warn('[WARN] - Analysis failed, using default analysis');
                        analysisData = {
                            keyPoints: ['Email content needs response'],
                            sentiment: 'neutral',
                            responseStrategy: 'respond professionally and appropriately'
                        };
                    }
                } catch (analysisError) {
                    console.warn('[WARN] - Analysis failed, using default analysis:', analysisError);
                    analysisData = {
                        keyPoints: ['Email content needs response'],
                        sentiment: 'neutral',
                        responseStrategy: 'respond professionally and appropriately'
                    };
                }
                
                this.uiController.showStatus('Generating response...');
            }
            
            // Generate response
            this.currentResponse = await this.aiService.generateResponse(
                this.currentEmail, 
                analysisData,
                { ...config, ...responseConfig }
            );
            
            console.info('[INFO] - Response generated:', this.currentResponse);
            
            // Initialize conversation history with the original email context
            this.initializeConversationHistory(this.currentEmail, analysisData);
            
            // Show analysis section since we no longer have a separate response tab
            this.showAnalysisSection();
            
            // Show chat section instead of just refinement
            this.showChatSection();
            this.updateWorkflowStep(4); // Move to step 4 (chat active)
            
            // Initialize chat with welcome message and the generated response
            this.initializeChatWithResponse();
            
            this.uiController.showStatus('Chat assistant ready! Your initial response is generated. Start chatting to refine it.');
            
        } catch (error) {
            console.error('[ERROR] - Response generation failed:', error);
            
            // Provide more specific error messages based on error type
            let userMessage = 'Failed to generate response. Please try again.';
            let showSettings = false;
            
            if (error.message && error.message.includes('Authentication failed')) {
                userMessage = 'Response generation failed: Invalid or missing API key. Please check your API key in the settings panel.';
                showSettings = true;
            } else if (error.message && error.message.includes('Access forbidden')) {
                userMessage = 'Response generation failed: API key permissions issue. Please verify your key has the correct permissions.';
                showSettings = true;
            } else if (error.message && error.message.includes('Service not found')) {
                userMessage = 'Response generation failed: Service endpoint not found. Please verify your endpoint URL in settings.';
                showSettings = true;
            } else if (error.message && error.message.includes('Rate limit exceeded')) {
                userMessage = 'Response generation failed: Rate limit exceeded. Please wait a moment and try again.';
            }
            
            this.uiController.showError(userMessage);
            
            // If it's a configuration issue, provide guidance to fix it
            if (showSettings) {
                setTimeout(() => {
                    const settingsTab = document.querySelector('button[data-tab="settings"]');
                    if (settingsTab) {
                        settingsTab.click();
                    }
                }, 2000);
            }
        } finally {
            this.uiController.setButtonLoading('generate-response', false);
        }
    }

    async generateFollowupSuggestions() {
        if (!this.currentEmail) {
            this.uiController.showError('No email available for follow-up suggestions.');
            return;
        }

        try {
            this.uiController.showStatus('Generating follow-up suggestions...');
            this.uiController.setButtonLoading('generate-response', true);
            
            // Get configuration
            const config = this.getAIConfiguration();
            const responseConfig = this.getResponseConfiguration();
            
            // Ensure we have analysis data - if not, run analysis first
            let analysisData = this.currentAnalysis;
            if (!analysisData) {
                console.warn('[WARN] - No current analysis available, running analysis first');
                this.uiController.showStatus('Analyzing sent email before generating follow-up suggestions...');
                
                try {
                    await this.performAnalysis();
                    analysisData = this.currentAnalysis;
                    
                    if (!analysisData) {
                        analysisData = {
                            keyPoints: ['Sent email content analyzed'],
                            sentiment: 'neutral',
                            responseStrategy: 'generate appropriate follow-up suggestions'
                        };
                    }
                } catch (analysisError) {
                    console.warn('[WARN] - Analysis failed, using default analysis:', analysisError);
                    analysisData = {
                        keyPoints: ['Sent email content analyzed'],
                        sentiment: 'neutral', 
                        responseStrategy: 'generate appropriate follow-up suggestions'
                    };
                }
                
                this.uiController.showStatus('Generating follow-up suggestions...');
            }
            
            // Generate follow-up suggestions instead of response
            this.currentResponse = await this.aiService.generateFollowupSuggestions(
                this.currentEmail, 
                analysisData,
                { ...config, ...responseConfig }
            );
            
            console.info('[INFO] - Follow-up suggestions generated:', this.currentResponse);
            
            // Log telemetry for follow-up suggestions generation
            this.logger.logEvent('followup_suggestions_generated', {
                model_service: config.service,
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                recipients_count: this.currentEmail.recipients.split(',').length,
                suggestions_length: this.currentResponse.suggestions ? this.currentResponse.suggestions.length : 0,
                analysis_available: !!analysisData,
                generation_success: true,
                refinement_count: this.refinementCount
            }, 'Information', this.getUserEmailForTelemetry());
            
            // Show analysis section since we no longer have a separate response tab
            this.showAnalysisSection();
            
            // Follow the same workflow as manual response generation
            this.showChatSection();
            this.updateWorkflowStep(4); // Move directly to step 4 (chat active)
            
            // Initialize chat with the generated suggestions
            this.initializeChatWithResponse();
            
            this.uiController.showStatus('Follow-up suggestions generated successfully.');
            
        } catch (error) {
            console.error('[ERROR] - Follow-up suggestion generation failed:', error);
            
            // Log telemetry for failed follow-up suggestions
            this.logger.logEvent('followup_suggestions_failed', {
                error_message: error.message,
                model_service: config ? config.service : 'unknown',
                email_length: this.currentEmail ? this.currentEmail.bodyLength : 0,
                analysis_available: !!analysisData
            }, 'Error', this.getUserEmailForTelemetry());
            
            // Provide more specific error messages based on error type
            let userMessage = 'Failed to generate follow-up suggestions. Please try again.';
            let showSettings = false;
            
            if (error.message && error.message.includes('Authentication failed')) {
                userMessage = 'Follow-up generation failed: Invalid or missing API key. Please check your API key in the settings panel.';
                showSettings = true;
            } else if (error.message && error.message.includes('Access forbidden')) {
                userMessage = 'Follow-up generation failed: API key permissions issue. Please verify your key has the correct permissions.';
                showSettings = true;
            } else if (error.message && error.message.includes('Service not found')) {
                userMessage = 'Follow-up generation failed: Service endpoint not found. Please verify your endpoint URL in settings.';
                showSettings = true;
            } else if (error.message && error.message.includes('Rate limit exceeded')) {
                userMessage = 'Follow-up generation failed: Rate limit exceeded. Please wait a moment and try again.';
            }
            
            this.uiController.showError(userMessage);
            
            // If it's a configuration issue, provide guidance to fix it
            if (showSettings) {
                setTimeout(() => {
                    const settingsTab = document.querySelector('button[data-tab="settings"]');
                    if (settingsTab) {
                        settingsTab.click();
                    }
                }, 2000);
            }
        } finally {
            this.uiController.setButtonLoading('generate-response', false);
        }
    }

    async sendChatMessage() {
        const chatInput = document.getElementById('chat-input');
        const message = chatInput?.value?.trim();
        
        if (!message) {
            return;
        }

        if (!this.currentEmail || !this.currentResponse) {
            this.uiController.showError('Please generate a response first before starting a chat.');
            return;
        }

        try {
            // Add user message to chat
            this.addChatMessage('user', message);
            
            // Clear input and disable send button
            chatInput.value = '';
            this.uiController.setButtonLoading('send-chat-message', true);
            
            // Show loading indicator
            this.showChatLoading();
            
            // Get current configuration
            const config = this.getAIConfiguration();
            const responseConfig = this.getResponseConfiguration();
            
            // Store previous response for history tracking
            const previousResponse = this.currentResponse.text;
            
            // Use history-aware refinement with chat context
            this.currentResponse = await this.aiService.refineResponseWithHistory(
                this.currentResponse,
                message,
                config,
                responseConfig,
                this.originalEmailContext,
                this.conversationHistory
            );
            
            // Add this chat step to conversation history
            this.addToConversationHistory(
                message,
                previousResponse,
                this.currentResponse.text
            );
            
            // Remove loading indicator
            this.removeChatLoading();
            
            // Add AI response to chat
            this.addChatMessage('assistant', this.currentResponse.text);
            
            // Response is now only shown in chat interface
            
            // Increment refinement counter for telemetry
            this.refinementCount++;
            
            // Log chat interaction
            this.logger.logEvent('chat_message_sent', {
                refinement_count: this.refinementCount,
                message_length: message.length,
                conversation_length: this.conversationHistory.length
            }, 'Information', this.getUserEmailForTelemetry());
            
        } catch (error) {
            console.error('[ERROR] - Chat message failed:', error);
            this.removeChatLoading();
            this.addChatMessage('system', 'Sorry, I encountered an error processing your message. Please try again.');
            this.uiController.showError('Failed to process chat message. Please try again.');
        } finally {
            this.uiController.setButtonLoading('send-chat-message', false);
        }
    }

    addChatMessage(type, content) {
        const chatMessages = document.getElementById('chat-messages');
        if (!chatMessages) return;

        const messageDiv = document.createElement('div');
        messageDiv.className = `chat-message ${type}`;
        
        const timestamp = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        
        let headerText = '';
        if (type === 'user') {
            headerText = `You • ${timestamp}`;
        } else if (type === 'assistant') {
            headerText = `AI Assistant • ${timestamp}`;
        } else if (type === 'system') {
            headerText = `System • ${timestamp}`;
        }

        const messageContent = type === 'assistant' 
            ? this.renderWithHtmlTables(content)  // Render tables for AI responses
            : this.escapeHtml(content);  // Escape user messages

        messageDiv.innerHTML = `
            <div class="chat-message-header">${headerText}</div>
            <div class="chat-message-content">${messageContent}</div>
        `;

        chatMessages.appendChild(messageDiv);
        
        // Scroll to bottom
        chatMessages.scrollTop = chatMessages.scrollHeight;
        
        // Auto-focus input for user messages (but not for system messages)
        if (type !== 'system') {
            const chatInput = document.getElementById('chat-input');
            if (chatInput) {
                setTimeout(() => chatInput.focus(), 100);
            }
        }
    }

    showChatLoading() {
        const chatMessages = document.getElementById('chat-messages');
        if (!chatMessages) return;

        const loadingDiv = document.createElement('div');
        loadingDiv.className = 'chat-loading';
        loadingDiv.id = 'chat-loading-indicator';
        loadingDiv.textContent = 'AI is thinking...';

        chatMessages.appendChild(loadingDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    removeChatLoading() {
        const loadingIndicator = document.getElementById('chat-loading-indicator');
        if (loadingIndicator) {
            loadingIndicator.remove();
        }
    }

    clearChatHistory() {
        const chatMessages = document.getElementById('chat-messages');
        if (chatMessages) {
            chatMessages.innerHTML = '';
        }

        // Clear conversation history but keep original email context
        this.conversationHistory = [];
        
        // Add a system message
        this.addChatMessage('system', 'Chat history cleared. The original email response is still available.');
        
        // Log chat clear event
        this.logger.logEvent('chat_cleared', {
            messages_cleared: this.conversationHistory.length
        }, 'Information', this.getUserEmailForTelemetry());
    }

    async copyLatestResponse() {
        try {
            // Find the most recent assistant message in the chat
            const chatMessages = document.getElementById('chat-messages');
            if (!chatMessages) {
                this.uiController.showError('No chat messages found.');
                return;
            }

            const assistantMessages = chatMessages.querySelectorAll('.chat-message.assistant');
            if (assistantMessages.length === 0) {
                this.uiController.showError('No AI responses found to copy.');
                return;
            }

            // Get the last assistant message
            const lastAssistantMessage = assistantMessages[assistantMessages.length - 1];
            const messageContent = lastAssistantMessage.querySelector('.chat-message-content');
            
            if (!messageContent) {
                this.uiController.showError('Could not find message content to copy.');
                return;
            }

            // Get the text content, preserving HTML tables if present
            const responseText = messageContent.innerHTML;
            
            // Use the existing copy functionality with HTML table support
            await this.copyResponseWithHtml(responseText);
            
            // Show success feedback
            this.uiController.showStatus('Latest response copied to clipboard!');
            
            // Mark that clipboard was used for telemetry
            this.hasUsedClipboard = true;
            
            // Log copy event
            this.logger.logEvent('latest_response_copied', {
                conversation_length: this.conversationHistory.length,
                response_length: responseText.length
            }, 'Information', this.getUserEmailForTelemetry());
            
        } catch (error) {
            console.error('[ERROR] - Failed to copy latest response:', error);
            this.uiController.showError('Failed to copy response to clipboard. Please try again.');
        }
    }

    getAIConfiguration() {
        // Prioritize saved settings over UI dropdown for both service and model selection
        // This ensures auto-analysis uses the user's actual saved preferences
        const settings = this.settingsManager.getSettings();
        const service = settings['model-service'] || (this.modelServiceSelect ? this.modelServiceSelect.value : '');
        
        // For model selection, prioritize saved settings first
        let model = settings['model-select'] || '';
        
        // Only use UI dropdown value if no saved setting exists
        if (!model && this.modelSelect && this.modelSelect.value) {
            model = this.modelSelect.value;
        }
        
        // Final fallback to getSelectedModel() for backward compatibility
        if (!model) {
            model = this.getSelectedModel();
        }
        
        window.debugLog(`[VERBOSE] - getAIConfiguration: service from settings: '${settings['model-service']}', from dropdown: '${this.modelServiceSelect ? this.modelServiceSelect.value : 'N/A'}', using: '${service}'`);
        window.debugLog(`[VERBOSE] - getAIConfiguration: model from settings: '${settings['model-select']}', from dropdown: '${this.modelSelect ? this.modelSelect.value : 'N/A'}', using: '${model}'`);
        
        // Get provider-specific configuration
        let apiKey = '';
        let endpointUrl = '';
        
        // Check if settings panel is open for context-aware behavior
        const settingsPanel = document.getElementById('settings-panel');
        const isSettingsOpen = settingsPanel && !settingsPanel.classList.contains('hidden');
        
        if (service) {
            const providerConfig = this.settingsManager.getProviderConfig(service);
            apiKey = providerConfig['api-key'] || '';
            endpointUrl = providerConfig['endpoint-url'] || '';
            
            window.debugLog(`[VERBOSE] - getAIConfiguration: provider config for '${service}': apiKey=${apiKey ? '[HIDDEN]' : 'EMPTY'}, endpoint=${endpointUrl}`);
            
            // Use UI field values ONLY when in settings context and they are visible/populated
            // In main taskpane context, always use saved provider configurations
            
            if (isSettingsOpen) {
                // We're in settings context - UI fields are visible and may have unsaved changes
                const apiKeyElement = document.getElementById('api-key');
                const endpointUrlElement = document.getElementById('endpoint-url');
                if (apiKeyElement && apiKeyElement.value.trim()) {
                    apiKey = apiKeyElement.value.trim();
                    window.debugLog(`[VERBOSE] - getAIConfiguration: using API key from settings UI field`);
                }
                if (endpointUrlElement && endpointUrlElement.value.trim()) {
                    const uiEndpoint = endpointUrlElement.value.trim();
                    
                    // Validate that UI endpoint matches the expected provider
                    if (this.defaultProvidersConfig && this.defaultProvidersConfig[service]) {
                        const expectedEndpoint = this.defaultProvidersConfig[service].baseUrl;
                        if (expectedEndpoint && uiEndpoint !== expectedEndpoint) {
                            // Check if UI endpoint belongs to a different provider
                            let contaminatingProvider = null;
                            for (const [providerName, providerConfig] of Object.entries(this.defaultProvidersConfig)) {
                                if (providerName !== service && providerConfig.baseUrl === uiEndpoint) {
                                    contaminatingProvider = providerName;
                                    break;
                                }
                            }
                            
                            if (contaminatingProvider) {
                                console.warn(`[WARN] - getAIConfiguration: UI endpoint contamination detected! Provider ${service} has endpoint from ${contaminatingProvider}. Using saved endpoint.`);
                                // Keep the saved endpoint, don't use contaminated UI value
                                window.debugLog(`[VERBOSE] - getAIConfiguration: used saved endpoint instead of contaminated UI field: ${endpointUrl}`);
                            } else {
                                // UI has a custom endpoint, use it
                                endpointUrl = uiEndpoint;
                                window.debugLog(`[VERBOSE] - getAIConfiguration: using endpoint from settings UI field: ${endpointUrl}`);
                            }
                        } else {
                            endpointUrl = uiEndpoint;
                            window.debugLog(`[VERBOSE] - getAIConfiguration: using endpoint from settings UI field: ${endpointUrl}`);
                        }
                    } else {
                        endpointUrl = uiEndpoint;
                        window.debugLog(`[VERBOSE] - getAIConfiguration: using endpoint from settings UI field: ${endpointUrl}`);
                    }
                }
            } else {
                // Main taskpane context - always use saved provider configurations
                // UI fields are hidden and not reliable in main context
                window.debugLog(`[VERBOSE] - getAIConfiguration: using saved provider config (main taskpane context)`);
            }
        }
        
        const config = {
            service,
            apiKey,
            endpointUrl,
            model,
            settingsManager: this.settingsManager
        };
        
        window.debugLog(`[VERBOSE] - getAIConfiguration returning:`, { 
            service: config.service, 
            apiKey: config.apiKey ? '[HIDDEN]' : 'EMPTY', 
            endpointUrl: config.endpointUrl, 
            model: config.model,
            savedModelService: settings['model-service'],
            uiDropdownValue: this.modelServiceSelect ? this.modelServiceSelect.value : 'N/A',
            settingsModalOpen: isSettingsOpen,
            usedUIFields: !isSettingsOpen
        });
        
        return config;
    }

    /**
     * Get AI configuration specifically for settings operations (uses settings provider dropdown)
     * @returns {Object} AI configuration object
     */
    getSettingsAIConfiguration() {
        // Get provider from settings dropdown (not main taskpane dropdown)
        const settingsProviderSelect = document.getElementById('settings-provider-select');
        const service = settingsProviderSelect?.value || '';
        
        if (!service) {
            window.debugLog('[VERBOSE] - getSettingsAIConfiguration: No provider selected in settings');
            return { service: '', apiKey: '', endpointUrl: '', model: '' };
        }
        
        // Get provider-specific configuration from saved settings
        const providerConfig = this.settingsManager.getProviderConfig(service);
        let apiKey = providerConfig['api-key'] || '';
        let endpointUrl = providerConfig['endpoint-url'] || '';
        
        // For settings context, use current UI values (for immediate testing before saving)
        // This ensures test connection uses what the user just entered
        const apiKeyElement = document.getElementById('api-key');
        const endpointUrlElement = document.getElementById('endpoint-url');
        if (apiKeyElement && apiKeyElement.value) {
            apiKey = apiKeyElement.value;
        }
        if (endpointUrlElement && endpointUrlElement.value) {
            endpointUrl = endpointUrlElement.value;
        }
        
        // For settings operations, we don't need a specific model - use default
        const model = this.getDefaultModelForProvider(service);
        
        const config = {
            service,
            apiKey,
            endpointUrl,
            model
        };
        
        window.debugLog(`[VERBOSE] - getSettingsAIConfiguration returning:`, { 
            service: config.service, 
            apiKey: config.apiKey ? '[HIDDEN]' : 'EMPTY', 
            endpointUrl: config.endpointUrl, 
            model: config.model
        });
        
        return config;
    }

    getResponseConfiguration() {
        // Use default values: Brief and Professional
        return {
            length: 2, // Brief
            tone: 3    // Professional
        };
    }

    getSelectedModel() {
        const service = this.modelServiceSelect ? this.modelServiceSelect.value : '';
        
        // Get default model from provider configuration instead of hardcoded map
        if (this.defaultProvidersConfig && this.defaultProvidersConfig[service]) {
            return this.defaultProvidersConfig[service].defaultModel || this.getFallbackModel();
        }
        
        // Final fallback from global config or ultimate hardcoded fallback
        return this.getFallbackModel();
    }

    getDefaultModelForProvider(provider) {
        if (!provider || !this.defaultProvidersConfig) {
            return this.getFallbackModel();
        }
        
        if (this.defaultProvidersConfig[provider] && this.defaultProvidersConfig[provider].defaultModel) {
            return this.defaultProvidersConfig[provider].defaultModel;
        }
        
        return this.getFallbackModel();
    }

    providerNeedsApiKey(provider) {
        // Ollama typically runs locally and doesn't need an API key
        if (provider === 'ollama') {
            return false;
        }
        
        // Most other providers (OpenAI, Claude, etc.) require API keys
        if (provider === 'openai' || provider === 'anthropic' || provider === 'claude') {
            return true;
        }
        
        // For onsite1/onsite2 or custom providers, assume they need API keys unless explicitly configured otherwise
        if (this.defaultProvidersConfig && this.defaultProvidersConfig[provider]) {
            // Check if the provider config indicates no API key needed
            return this.defaultProvidersConfig[provider].requiresApiKey !== false;
        }
        
        // Default to requiring API key for unknown providers
        return true;
    }

    getFallbackModel() {
        // Ultimate hardcoded fallback for internal deployments
        return 'llama3:latest';
    }

    async updateModelDropdown() {
        if (!this.modelServiceSelect || !this.modelSelectGroup || !this.modelSelect) return;
        
        const aiConfigPlaceholder = document.getElementById('ai-config-placeholder');
        this.modelSelectGroup.style.display = 'none';
        this.modelSelect.innerHTML = '';
        let models = [];
        let preferred = '';
        let errorMsg = '';
        if (this.modelServiceSelect.value === 'ollama') {
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            const endpointUrlElement = document.getElementById('endpoint-url');
            let baseUrl = (endpointUrlElement && endpointUrlElement.value) || 'http://localhost:11434';
            
            // Ensure we're using the correct Ollama endpoint
            const defaultOllamaUrl = this.defaultProvidersConfig?.ollama?.baseUrl || 'http://localhost:11434';
            if (baseUrl !== defaultOllamaUrl) {
                console.warn(`[WARN] - Ollama endpoint URL mismatch. Expected: ${defaultOllamaUrl}, Found: ${baseUrl}. Using correct URL.`);
                baseUrl = defaultOllamaUrl;
                // Update the UI to show the correct endpoint
                if (endpointUrlElement) {
                    endpointUrlElement.value = baseUrl;
                }
            }
            
            try {
                models = await AIService.fetchOllamaModels(baseUrl);
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
                preferred = this.defaultProvidersConfig?.ollama?.defaultModel || this.getFallbackModel();
                if (preferred && models.includes(preferred)) {
                    this.modelSelect.value = preferred;
                } else if (models.length) {
                    this.modelSelect.value = models[0];
                }
                
                // Save the model selection to settings if one was set
                if (this.modelSelect.value) {
                    await this.saveSettings();
                }
            } catch (err) {
                errorMsg = `Error fetching models: ${err.message || err}`;
                this.modelSelect.innerHTML = '<option value="">Error fetching models</option>';
            }
        } else if (this.modelServiceSelect.value !== 'ollama') {
            // Handle OpenAI-compatible services (openai, onsite1, onsite2, etc.)
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            
            const serviceKey = this.modelServiceSelect.value;
            
            // Get endpoint URL: user input -> provider config -> configured fallback
            let endpoint = '';
            const endpointUrlElement = document.getElementById('endpoint-url');
            if (endpointUrlElement && endpointUrlElement.value) {
                endpoint = endpointUrlElement.value;
            } else if (this.defaultProvidersConfig && this.defaultProvidersConfig[serviceKey] && this.defaultProvidersConfig[serviceKey].baseUrl) {
                endpoint = this.defaultProvidersConfig[serviceKey].baseUrl;
            } else {
                endpoint = 'http://localhost:11434/v1';
            }
            
            if (endpoint.endsWith('/')) endpoint = endpoint.slice(0, -1);
            const apiKey = document.getElementById('api-key').value;
            try {
                models = await AIService.fetchOpenAICompatibleModels(endpoint, apiKey);
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
            } catch (err) {
                errorMsg = err.message || `Error fetching models: ${err}`;
                
                // Show more helpful error display for authentication issues
                if (err.message && err.message.includes('Authentication failed')) {
                    // Highlight API key field or show settings reminder
                    const apiKeyField = document.getElementById('api-key');
                    if (apiKeyField) {
                        apiKeyField.style.borderColor = '#dc3545';
                        apiKeyField.style.borderWidth = '2px';
                        // Remove highlight after 5 seconds
                        setTimeout(() => {
                            apiKeyField.style.borderColor = '';
                            apiKeyField.style.borderWidth = '';
                        }, 5000);
                    }
                }
                
                this.modelSelect.innerHTML = '<option value="">Error fetching models</option>';
            }
            // Get preferred model from the specific service's config, not just openai
            preferred = this.defaultProvidersConfig?.[serviceKey]?.defaultModel || this.getFallbackModel();
            if (preferred && models.includes(preferred)) {
                this.modelSelect.value = preferred;
            } else if (models.length) {
                this.modelSelect.value = models[0];
            }
        } else {
            // Hide model dropdown for services that don't support model selection
            this.modelSelectGroup.style.display = 'none';
        }
        // Show error if needed
        let errorDiv = document.getElementById('model-select-error');
        if (errorMsg) {
            if (!errorDiv) {
                errorDiv = document.createElement('div');
                errorDiv.id = 'model-select-error';
                errorDiv.style.cssText = 'color: #dc3545; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; padding: 8px; margin-top: 5px; font-size: 14px; line-height: 1.4;';
                this.modelSelectGroup.appendChild(errorDiv);
            }
            
            // Create more helpful error messages with action suggestions
            let displayMessage = errorMsg;
            if (errorMsg.includes('Authentication failed')) {
                displayMessage = '🔐 ' + errorMsg + '\n\n💡 Tip: Check that your API key is entered correctly and has not expired.';
            } else if (errorMsg.includes('Access forbidden')) {
                displayMessage = '🚫 ' + errorMsg + '\n\n💡 Tip: Contact your administrator to verify API key permissions.';
            } else if (errorMsg.includes('Service not found')) {
                displayMessage = '🔗 ' + errorMsg + '\n\n💡 Tip: Verify your endpoint URL is correct and the service is running.';
            } else if (errorMsg.includes('Rate limit exceeded')) {
                displayMessage = '⏰ ' + errorMsg + '\n\n💡 Tip: Wait a few moments before trying again.';
            }
            
            errorDiv.innerHTML = displayMessage.replace(/\n/g, '<br>');
            errorDiv.style.display = 'block';
        } else if (errorDiv) {
            errorDiv.style.display = 'none';
        }
        // Hide AI config placeholder in main UI if model discovery succeeds
        if (aiConfigPlaceholder) {
            aiConfigPlaceholder.classList.add('hidden');
            aiConfigPlaceholder.innerHTML = '';
        }
        
        // Set up model change event listener if not already bound
        if (this.modelSelect && !this.modelSelect.hasModelChangeListener) {
            this.modelSelect.addEventListener('change', (e) => this.onModelChange(e));
            this.modelSelect.hasModelChangeListener = true;
        }
    }

    displayAnalysis(analysis) {
        const container = document.getElementById('email-analysis');
        
        // Make sure the analysis section is visible
        this.showAnalysisSection();
        
        // Build due dates section if present
        let dueDatesHtml = '';
        if (analysis.dueDates && analysis.dueDates.length > 0) {
            const dueDateItems = analysis.dueDates.map(dueDate => {
                const urgentClass = dueDate.isUrgent ? 'urgent-due-date' : '';
                const dateDisplay = dueDate.date !== 'unspecified' ? dueDate.date : 'Date not specified';
                const timeDisplay = dueDate.time !== 'unspecified' ? ` at ${dueDate.time}` : '';
                
                return `<li class="due-date-item ${urgentClass}">
                    <strong>${this.escapeHtml(dueDate.description)}</strong><br>
                    <span class="due-date-info">Due: ${dateDisplay}${timeDisplay}</span>
                    ${dueDate.isUrgent ? '<span class="urgent-badge">URGENT</span>' : ''}
                </li>`;
            }).join('');
            
            dueDatesHtml = `
                <h3 class="due-dates-header">⏰ Due Dates & Deadlines</h3>
                <ul class="due-dates-list">
                    ${dueDateItems}
                </ul>
            `;
        }
        
        container.innerHTML = `
            <div class="analysis-content">
                ${dueDatesHtml}
                
                <h3>Key Points</h3>
                <ul>
                    ${analysis.keyPoints.map(point => `<li>${this.escapeHtml(point)}</li>`).join('')}
                </ul>

                <h3>Intent & Sentiment</h3>
                <ul>
                    <li><strong>Purpose:</strong> ${this.escapeHtml(analysis.intent || 'Not specified')}</li>
                    <li><strong>Tone:</strong> ${this.escapeHtml(analysis.sentiment)}</li>
                    <li><strong>Urgency:</strong> ${analysis.urgencyLevel}/5 - ${this.escapeHtml(analysis.urgencyReason || 'No reason provided')}</li>
                </ul>

                <h3>Recommended Actions</h3>
                <ul>
                    ${analysis.actions.map(action => `<li>${this.escapeHtml(action)}</li>`).join('')}
                </ul>
                
                ${analysis.responseStrategy ? `
                <h3>Response Strategy</h3>
                <ul>
                    <li>${this.escapeHtml(analysis.responseStrategy)}</li>
                </ul>
                ` : ''}
            </div>
        `;
    }

    displayResponse(response) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Displaying response:', response);
        const container = document.getElementById('response-draft');
        
        if (!container) {
            console.error('[ERROR] - response-draft container not found');
            return;
        }
        
        if (!response || (!response.text && !response.suggestions)) {
            console.error('[ERROR] - Invalid response object:', response);
            container.innerHTML = '<div class="error">Error: Invalid response received</div>';
            return;
        }
        
        // Handle both regular responses (text) and follow-up suggestions (suggestions)
        const responseContent = response.text || response.suggestions;
        
        // Use separate formatting for display (less aggressive than clipboard)
        const cleanText = this.formatTextForDisplay(responseContent);
        
        // Render with HTML table support
        const formattedContent = this.renderWithHtmlTables(cleanText);
        
        container.innerHTML = `
            <div class="response-content">
                <div class="response-text" id="response-text-content">
                    ${formattedContent}
                </div>
            </div>
        `;
        
        console.info('[INFO] - Response displayed successfully');
    }

    /**
     * Renders content with HTML table support while keeping other content safe
     * @param {string} text - The text content that may contain HTML tables
     * @returns {string} Safely rendered HTML content
     */
    renderWithHtmlTables(text) {
        if (!text) return '';
        
        // Check if the text contains HTML table elements
        const hasHtmlTables = /<table[\s\S]*?<\/table>/gi.test(text);
        
        if (!hasHtmlTables) {
            // No tables - use standard escaping and line break conversion
            return this.escapeHtml(text).replace(/\n/g, '<br>');
        }
        
        // Clean up excessive spacing around tables first
        let cleanedText = text.replace(/\n+(<table)/gi, '\n\n$1');
        cleanedText = cleanedText.replace(/(<\/table>)\n+/gi, '$1\n\n');
        
        // Split content into table and non-table parts
        const parts = [];
        let lastIndex = 0;
        const tableRegex = /<table[\s\S]*?<\/table>/gi;
        let match;
        
        while ((match = tableRegex.exec(cleanedText)) !== null) {
            // Add text before table (escaped)
            if (match.index > lastIndex) {
                const beforeTable = cleanedText.substring(lastIndex, match.index);
                const escapedBefore = this.escapeHtml(beforeTable.trim());
                if (escapedBefore) {
                    parts.push(escapedBefore.replace(/\n/g, '<br>'));
                }
            }
            
            // Add table (sanitized but not escaped)
            const tableHtml = this.sanitizeHtmlTable(match[0]);
            parts.push(tableHtml);
            
            lastIndex = match.index + match[0].length;
        }
        
        // Add remaining text after last table (escaped)
        if (lastIndex < cleanedText.length) {
            const afterTable = cleanedText.substring(lastIndex);
            const escapedAfter = this.escapeHtml(afterTable.trim());
            if (escapedAfter) {
                parts.push(escapedAfter.replace(/\n/g, '<br>'));
            }
        }
        
        return parts.join('');
    }

    /**
     * Sanitizes HTML table content to ensure it's safe while preserving table structure
     * @param {string} tableHtml - HTML table string
     * @returns {string} Sanitized table HTML
     */
    sanitizeHtmlTable(tableHtml) {
        // Allow only table-related tags and basic formatting
        const allowedTags = ['table', 'thead', 'tbody', 'tfoot', 'tr', 'th', 'td', 'caption'];
        const allowedAttributes = ['style', 'class', 'colspan', 'rowspan'];
        
        // Basic sanitization - remove script tags and dangerous attributes
        let sanitized = tableHtml
            .replace(/<script[\s\S]*?<\/script>/gi, '')
            .replace(/<iframe[\s\S]*?<\/iframe>/gi, '')
            .replace(/on\w+\s*=\s*["'][^"']*["']/gi, '') // Remove event handlers
            .replace(/javascript:/gi, ''); // Remove javascript: URLs
        
        // Ensure proper table structure and add default styling if missing
        if (!sanitized.includes('style=') && !sanitized.includes('border')) {
            sanitized = sanitized.replace(
                /<table(?![^>]*style=)/gi, 
                '<table style="border-collapse: collapse; width: 100%; margin: 10px 0; font-family: inherit;"'
            );
        }
        
        return sanitized;
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    async copyResponse() {
        try {
            // Get the original response text from the currentResponse object for better formatting
            let responseText = '';
            
            if (this.currentResponse && (this.currentResponse.text || this.currentResponse.suggestions)) {
                // Handle both regular responses (text) and follow-up suggestions (suggestions)
                responseText = this.currentResponse.text || this.currentResponse.suggestions;
            } else {
                // Fallback to displayed content if currentResponse not available
                const responseElement = document.getElementById('response-text-content');
                responseText = responseElement ? responseElement.textContent : '';
            }
            
            if (!responseText) {
                this.uiController.showError('No response text to copy.');
                return;
            }
            
            // Check if the response contains HTML tables
            const hasHtmlTables = /<table[\s\S]*?<\/table>/gi.test(responseText);
            
            if (hasHtmlTables) {
                // Copy as both HTML and plain text for best compatibility
                await this.copyResponseWithHtml(responseText);
            } else {
                // Standard plain text copy
                const formattedText = this.formatTextForClipboard(responseText);
                await navigator.clipboard.writeText(formattedText);
            }
            
            this.uiController.showStatus('Response copied to clipboard.');
            
            // Track clipboard usage for telemetry
            this.hasUsedClipboard = true;
            
            // Log clipboard usage event
            this.logger.logEvent('response_copied', {
                content_type: this.currentResponse.suggestions ? 'followup_suggestions' : 'standard_response',
                email_context: this.currentEmail.context ? (this.currentEmail.context.isSentMail ? 'sent' : 'inbox') : 'unknown',
                refinement_count: this.refinementCount,
                response_length: responseText.length,
                contains_tables: hasHtmlTables
            }, 'Information', this.getUserEmailForTelemetry());
        } catch (error) {
            console.error('[ERROR] - Failed to copy response:', error);
            this.uiController.showError('Failed to copy response to clipboard.');
        }
    }

    /**
     * Copies response with HTML table support to clipboard
     * @param {string} responseText - The response text containing HTML tables
     */
    async copyResponseWithHtml(responseText) {
        // Create both HTML and plain text versions
        const htmlContent = this.formatResponseForHtmlClipboard(responseText);
        const plainTextContent = this.formatResponseForPlainTextClipboard(responseText);
        
        // Use the modern ClipboardItem API for rich content
        if (navigator.clipboard && typeof ClipboardItem !== 'undefined') {
            const clipboardItem = new ClipboardItem({
                'text/html': new Blob([htmlContent], { type: 'text/html' }),
                'text/plain': new Blob([plainTextContent], { type: 'text/plain' })
            });
            await navigator.clipboard.write([clipboardItem]);
        } else {
            // Fallback to plain text only
            await navigator.clipboard.writeText(plainTextContent);
        }
    }

    /**
     * Formats response for HTML clipboard (preserves tables)
     * @param {string} responseText - Raw response text
     * @returns {string} HTML formatted content
     */
    formatResponseForHtmlClipboard(responseText) {
        let formatted = responseText.trim();
        
        // Remove excessive line breaks around tables
        formatted = formatted.replace(/\n+(<table[\s\S]*?<\/table>)\n+/gi, '\n\n$1\n\n');
        formatted = formatted.replace(/^(<table[\s\S]*?<\/table>)\n+/gi, '$1\n\n');
        formatted = formatted.replace(/\n+(<table[\s\S]*?<\/table>)$/gi, '\n\n$1');
        
        // Convert line breaks to HTML, but be more conservative around tables
        const styledContent = formatted
            .replace(/\n{3,}/g, '\n\n') // Reduce multiple line breaks
            .replace(/\n/g, '<br>')
            .replace(/<br><br>(<table)/gi, '<br>$1') // Remove extra breaks before tables
            .replace(/(<\/table>)<br><br>/gi, '$1<br>') // Remove extra breaks after tables
            .replace(/<table(?![^>]*style=)/gi, '<table style="border-collapse: collapse; width: 100%; margin: 5px 0; font-family: Arial, sans-serif;"')
            .replace(/<th(?![^>]*style=)/gi, '<th style="border: 1px solid #ddd; padding: 8px; background-color: #f5f5f5; font-weight: bold;"')
            .replace(/<td(?![^>]*style=)/gi, '<td style="border: 1px solid #ddd; padding: 8px;"');
        
        return `<div style="font-family: Arial, sans-serif; line-height: 1.4;">${styledContent}</div>`;
    }

    /**
     * Formats response for plain text clipboard (converts tables to text)
     * @param {string} responseText - Raw response text
     * @returns {string} Plain text formatted content
     */
    formatResponseForPlainTextClipboard(responseText) {
        let formatted = responseText;
        
        // Convert HTML tables to plain text format
        formatted = formatted.replace(/<table[\s\S]*?<\/table>/gi, (tableMatch) => {
            return this.convertHtmlTableToPlainText(tableMatch);
        });
        
        // Remove any remaining HTML tags
        formatted = formatted.replace(/<[^>]*>/g, '');
        
        // Clean up excessive whitespace that might have been created
        formatted = formatted.replace(/\n{4,}/g, '\n\n'); // Reduce excessive line breaks
        formatted = formatted.replace(/[ \t]+/g, ' '); // Normalize spaces
        formatted = formatted.replace(/^\s+|\s+$/gm, ''); // Trim each line
        
        // Apply standard formatting but don't add extra paragraph breaks
        return this.formatTextForClipboardMinimal(formatted);
    }

    /**
     * Minimal formatting for clipboard - preserves structure without excessive spacing
     * @param {string} text - The text to format
     * @returns {string} Minimally formatted text
     */
    formatTextForClipboardMinimal(text) {
        let formatted = text.trim();
        
        // Normalize line endings
        formatted = formatted.replace(/\r\n?/g, '\n');
        
        // Only add breaks after greetings if they don't already exist
        if (!formatted.includes('\n\n')) {
            formatted = formatted.replace(/((?:Hi|Hello|Dear)\s+[^,]+,)\s*([A-Z])/gi, '$1\n\n$2');
            formatted = formatted.replace(/([.!?])\s*((?:Best\s+)?(?:regards?|sincerely|thanks?|cheers),?\s*\n?\s*[\w\s]+)$/gi, '$1\n\n$2');
        }
        
        // Clean up any excessive spacing
        formatted = formatted.replace(/\n{3,}/g, '\n\n');
        
        return formatted.trim();
    }

    /**
     * Initializes conversation history for a new email
     * @param {Object} emailData - The email being analyzed
     * @param {Object} analysis - The email analysis results
     */
    initializeConversationHistory(emailData, analysis) {
        this.originalEmailContext = {
            from: emailData.from,
            subject: emailData.subject,
            content: emailData.cleanBody || emailData.body,
            analysis: analysis,
            timestamp: new Date().toISOString()
        };
        
        this.conversationHistory = [];
        if (window.debugLog) {
            window.debugLog('[VERBOSE] - Conversation history initialized for new email');
        }
    }

    /**
     * Adds a refinement step to conversation history
     * @param {string} userInstruction - The refinement instruction from user
     * @param {string} previousResponse - The response being refined
     * @param {string} newResponse - The AI's refined response
     */
    addToConversationHistory(userInstruction, previousResponse, newResponse) {
        const conversationStep = {
            step: this.conversationHistory.length + 1,
            timestamp: new Date().toISOString(),
            userInstruction: userInstruction,
            previousResponse: previousResponse,
            newResponse: newResponse
        };
        
        this.conversationHistory.push(conversationStep);
        
        if (window.debugLog) {
            window.debugLog(`[VERBOSE] - Added refinement step ${conversationStep.step} to conversation history`);
        }
    }

    /**
     * Clears conversation history (called when starting new analysis)
     */
    clearConversationHistory() {
        this.conversationHistory = [];
        this.originalEmailContext = null;
        
        if (window.debugLog) {
            window.debugLog('[VERBOSE] - Conversation history cleared');
        }
    }

    /**
     * Converts an HTML table to plain text representation
     * @param {string} tableHtml - HTML table string
     * @returns {string} Plain text table representation
     */
    convertHtmlTableToPlainText(tableHtml) {
        try {
            // Create a temporary DOM element to parse the table
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = tableHtml;
            const table = tempDiv.querySelector('table');
            
            if (!table) return tableHtml; // Return original if parsing fails
            
            let textTable = '\n\n'; // Start with just two line breaks
            const rows = table.querySelectorAll('tr');
            
            rows.forEach((row, rowIndex) => {
                const cells = row.querySelectorAll('th, td');
                const cellTexts = Array.from(cells).map(cell => cell.textContent.trim());
                
                // Join cells with tabs for better formatting
                textTable += cellTexts.join('\t') + '\n';
                
                // Add separator line after header row
                if (rowIndex === 0 && row.querySelectorAll('th').length > 0) {
                    textTable += cellTexts.map(cell => '-'.repeat(Math.max(cell.length, 3))).join('\t') + '\n';
                }
            });
            
            return textTable + '\n'; // End with just one line break
        } catch (error) {
            console.warn('[WARN] - Failed to convert HTML table to plain text:', error);
            return tableHtml; // Return original if conversion fails
        }
    }

    /**
     * Format text for display in the TaskPane (more conservative than clipboard)
     * @param {string} text - The text to format
     * @returns {string} Formatted text for display
     */
    formatTextForDisplay(text) {
        // Start with the cleaned text
        let formatted = text.trim();
        
        // Remove ALL forms of tabs and tab-like characters aggressively
        formatted = formatted.replace(/\t/g, '');  // Regular tabs
        formatted = formatted.replace(/\u0009/g, ''); // Unicode tab
        formatted = formatted.replace(/\u00A0/g, ' '); // Non-breaking space to regular space
        formatted = formatted.replace(/\u2009/g, ' '); // Thin space to regular space
        formatted = formatted.replace(/\u200B/g, ''); // Zero-width space
        formatted = formatted.replace(/\u2000-\u200F/g, ''); // Various Unicode spaces
        formatted = formatted.replace(/\u2028/g, '\n'); // Line separator to newline
        formatted = formatted.replace(/\u2029/g, '\n\n'); // Paragraph separator to double newline
        
        // Remove excessive spaces
        formatted = formatted.replace(/[ ]{2,}/g, ' '); // Multiple spaces to single space
        
        // Normalize line endings
        formatted = formatted.replace(/\r\n?/g, '\n');
        
        // Remove leading/trailing whitespace from each line, including any hidden characters
        formatted = formatted.split('\n').map(line => {
            return line.replace(/^[\s\t\u00A0\u2000-\u200F\u2028\u2029]+|[\s\t\u00A0\u2000-\u200F\u2028\u2029]+$/g, '');
        }).join('\n');
        
        // Remove empty lines at the beginning and end
        formatted = formatted.replace(/^\n+/, '').replace(/\n+$/, '');
        
        // Only add minimal paragraph breaks - don't be as aggressive as clipboard version
        // Just ensure there's a break after greeting if it doesn't exist
        if (formatted.match(/(Hi|Hello|Dear)\s+[^,]+,\s*[A-Z]/)) {
            formatted = formatted.replace(/((?:Hi|Hello|Dear)\s+[^,]+,)\s*([A-Z])/gi, '$1\n\n$2');
        }
        
        if (window.debugLog) window.debugLog('[VERBOSE] - formatTextForDisplay - Original:', JSON.stringify(text));
        if (window.debugLog) window.debugLog('[VERBOSE] - formatTextForDisplay - Formatted:', JSON.stringify(formatted));
        
        return formatted;
    }

    /**
     * Format text for clipboard with proper line breaks and paragraph spacing
     * @param {string} text - The text to format
     * @returns {string} Formatted text with proper spacing
     */
    formatTextForClipboard(text) {
        // Start with the cleaned text
        let formatted = text.trim();
        
        // Remove any existing tabs and excessive spaces
        formatted = formatted.replace(/\t+/g, ' ');
        formatted = formatted.replace(/[ ]{2,}/g, ' ');
        
        // Normalize line endings to \n
        formatted = formatted.replace(/\r\n?/g, '\n');
        
        // If the text doesn't already have proper paragraph breaks, add them
        if (!formatted.includes('\n\n')) {
            // Add breaks after common greetings
            formatted = formatted.replace(/((?:Hi|Hello|Dear)\s+[^,]+,)\s*/gi, '$1\n\n');
            
            // Add breaks before common closings
            formatted = formatted.replace(/\s*((?:Best\s+)?(?:regards?|sincerely|thanks?|cheers),?\s*\n?\s*[\w\s]+)$/gi, '\n\n$1');
            
            // Add breaks after sentences that are likely to end paragraphs
            formatted = formatted.replace(/([.!?])\s+([A-Z][a-z])/g, '$1\n\n$2');
            
            // Clean up any triple+ line breaks
            formatted = formatted.replace(/\n{3,}/g, '\n\n');
        }
        
        // Ensure signature is on its own line
        formatted = formatted.replace(/([.!?])\s*((?:Best\s+)?(?:regards?|sincerely|thanks?|cheers),?)\s*([A-Z][\w\s]+)$/gi, '$1\n\n$2\n$3');
        
        // Final cleanup
        formatted = formatted.trim();
        
        if (window.debugLog) window.debugLog('[VERBOSE] - formatTextForClipboard Original:', JSON.stringify(text));
        if (window.debugLog) window.debugLog('[VERBOSE] - formatTextForClipboard Formatted:', JSON.stringify(formatted));
        
        return formatted;
    }



    showChatSection() {
        const chatSection = document.getElementById('refinement-section');
        
        if (chatSection) {
            chatSection.classList.remove('hidden');
            
            // Scroll to show the chat section
            setTimeout(() => {
                chatSection.scrollIntoView({ 
                    behavior: 'smooth', 
                    block: 'start' 
                });
            }, 300);
        }
    }

    initializeChat() {
        // Clear any existing chat messages
        const chatMessages = document.getElementById('chat-messages');
        if (chatMessages) {
            chatMessages.innerHTML = '';
        }
        
        // Add welcome message
        this.addChatMessage('system', 'Chat initialized! You can now refine your email response by asking questions or requesting changes.');
        
        // Focus the chat input
        const chatInput = document.getElementById('chat-input');
        if (chatInput) {
            setTimeout(() => chatInput.focus(), 500);
        }
    }

    initializeChatWithResponse() {
        // Clear any existing chat messages
        const chatMessages = document.getElementById('chat-messages');
        if (chatMessages) {
            chatMessages.innerHTML = '';
        }
        
        // Add the AI's response as the first message in chat
        this.addChatMessage('assistant', this.currentResponse.text);
        
        // Add helpful system message
        this.addChatMessage('system', 'Now you can chat with me to refine it! Try asking: "Make it shorter", "Add more details", "Create a table", "Change the tone", etc.');
        
        // Focus the chat input
        const chatInput = document.getElementById('chat-input');
        if (chatInput) {
            setTimeout(() => chatInput.focus(), 800);
        }
    }



    updateWorkflowStep(step) {
        // Show/hide different workflow sections based on step
        const step3SectionChat = document.getElementById('step3-section-chat');
        const step4Section = document.getElementById('step4-section');
        
        if (step === 3) {
            // Show Step 3 section after analysis is complete
            if (step3SectionChat) step3SectionChat.classList.remove('hidden');
            if (step4Section) step4Section.classList.add('hidden');
        } else if (step === 4) {
            // Chat active - hide Step 3, show Step 4
            if (step3SectionChat) step3SectionChat.classList.add('hidden');
            if (step4Section) step4Section.classList.remove('hidden');
        }
    }

    async onModelServiceChange(event) {
        if (window.debugLog) window.debugLog('[VERBOSE] onModelServiceChange triggered:', {
            value: event.target.value,
            oldValue: event.target.dataset.oldValue || 'undefined'
        });
        
        const customEndpoint = document.getElementById('custom-endpoint');
        if (customEndpoint) {
            if (event.target.value === 'custom') {
                customEndpoint.classList.remove('hidden');
            } else {
                customEndpoint.classList.add('hidden');
            }
        }
        
        // Record user's active choice to prevent domain filtering from overriding
        const newProvider = event.target.value;
        const oldProvider = event.target.dataset.oldValue;
        
        if (newProvider && newProvider !== 'undefined' && newProvider !== oldProvider) {
            // This is a real user change, not just initialization
            this.settingsManager.setSetting('user-active-provider-choice', newProvider);
            window.debugLog(`[VERBOSE] - Recorded user active choice: ${newProvider}`);
        }
        
        // Save current provider's settings before switching
        if (oldProvider && oldProvider !== 'undefined') {
            await this.saveCurrentProviderSettings(oldProvider);
        }
        
        // Load new provider's settings
        await this.loadProviderSettings(event.target.value);
        
        // Update provider labels in UI
        this.updateProviderLabels(event.target.value);
        
        // Reset Test Connection button state when switching providers
        this.resetTestConnectionButton();
        
        // Clear previous analysis and response since provider changed
        this.clearAnalysisAndResponse();
        
        // Store old value for next time
        event.target.dataset.oldValue = event.target.value;
        
        // Update model dropdown first (this will set default model and save settings)
        await this.updateModelDropdown();
        
        if (window.debugLog) window.debugLog('[VERBOSE] About to save settings after model service change');
        await this.saveSettings();
    }

    async onSettingsProviderChange(event) {
        const selectedProvider = event.target.value;
        if (window.debugLog) window.debugLog('[VERBOSE] Settings provider changed to:', selectedProvider);
        
        // Load the settings for the selected provider (settings-only version)
        await this.loadSettingsOnlyProviderConfig(selectedProvider);
        
        // Update provider labels in UI to reflect the selected provider
        this.updateProviderLabels(selectedProvider);
        
        // Reset Test Connection button state when switching providers
        this.resetTestConnectionButton();
    }

    async onModelChange(event) {
        if (window.debugLog) window.debugLog('[VERBOSE] onModelChange triggered:', {
            value: event.target.value,
            oldValue: event.target.dataset.oldValue || 'undefined',
            provider: this.modelServiceSelect?.value
        });
        
        // Only clear if this is actually a change (not just initialization)
        const oldModel = event.target.dataset.oldValue;
        if (oldModel && oldModel !== 'undefined' && oldModel !== event.target.value) {
            // Clear previous analysis and response since model changed
            this.clearAnalysisAndResponse();
            if (window.debugLog) window.debugLog('[VERBOSE] - Cleared analysis and response due to model change');
        }
        
        // Store old value for next time
        event.target.dataset.oldValue = event.target.value;
        
        // Save the model selection
        await this.saveSettings();
    }

    openSettings() {
        // Update current provider/model info in settings
        this.updateSettingsProviderInfo();
        document.getElementById('settings-panel').classList.remove('hidden');
    }

    updateSettingsProviderInfo() {
        const currentProviderElement = document.getElementById('settings-current-provider');
        const currentModelElement = document.getElementById('settings-current-model');
        
        if (currentProviderElement && currentModelElement) {
            const currentProvider = this.modelServiceSelect?.value || 'Not selected';
            const currentModel = this.modelSelect?.value || 'Not selected';
            
            // Get provider display name
            let providerDisplayName = currentProvider;
            if (this.defaultProvidersConfig && this.defaultProvidersConfig[currentProvider]) {
                providerDisplayName = this.defaultProvidersConfig[currentProvider].label || currentProvider;
            }
            
            currentProviderElement.textContent = providerDisplayName;
            currentModelElement.textContent = currentModel;
        }
    }

    closeSettings() {
        // Prevent Edge from detecting API key as password when closing settings
        const apiKeyField = document.getElementById('api-key');
        if (apiKeyField && apiKeyField.value) {
            // Temporarily store the value
            const apiKeyValue = apiKeyField.value;
            
            // Clear the field to prevent password manager detection
            apiKeyField.value = '';
            
            // Hide the settings panel
            document.getElementById('settings-panel').classList.add('hidden');
            
            // Restore the value after a brief delay
            setTimeout(() => {
                if (apiKeyField) {
                    apiKeyField.value = apiKeyValue;
                }
            }, 100);
        } else {
            // No API key value, just hide normally
            document.getElementById('settings-panel').classList.add('hidden');
        }
        
        // Force refresh the settings cache to ensure new API keys are immediately available
        // This ensures that getAIConfiguration() will use the newly saved settings
        setTimeout(() => {
            this.settingsManager.loadSettings().then(() => {
                if (window.debugLog) {
                    window.debugLog('[VERBOSE] - Settings refreshed after closing settings panel');
                }
            }).catch(error => {
                console.warn('[WARN] - Failed to refresh settings after closing settings panel:', error);
            });
        }, 150); // Wait a bit longer than the API key restoration
    }

    async resetSettings() {
        try {
            // Create a simple confirmation using the existing UI
            const confirmed = await this.showConfirmDialog(
                'Reset All Settings',
                'Are you sure you want to reset all settings to defaults? This will:\n\n' +
                '• Clear all API keys for all providers\n' +
                '• Reset all preferences to default values\n' +
                '• Clear all custom configurations\n' +
                '• Reset to default provider and model\n\n' +
                'This action cannot be undone.'
            );
            
            if (!confirmed) {
                return;
            }
            
            // Use SettingsManager to properly clear all settings
            const success = await this.settingsManager.clearAllSettings();
            
            if (success) {
                // Get default provider and model from S3 config
                const defaultProvider = this.defaultProvidersConfig?._config?.defaultProvider || 'ollama';
                const defaultModel = this.getDefaultModelForProvider(defaultProvider);
                
                // Set default provider and model
                if (this.modelServiceSelect) {
                    this.modelServiceSelect.value = defaultProvider;
                }
                
                // Save the default settings
                await this.settingsManager.saveSettings({
                    'model-service': defaultProvider,
                    'model-select': defaultModel
                });
                
                // Load the provider settings to populate UI with correct defaults
                await this.loadProviderSettings(defaultProvider);
                
                // Check if the default provider requires an API key
                const needsApiKey = this.providerNeedsApiKey(defaultProvider);
                
                if (needsApiKey) {
                    // Show success message with API key instruction
                    await this.showInfoDialog('Settings Reset - Action Required', 
                        `Settings have been reset to defaults.\n\nDefault provider: ${defaultProvider}\nDefault model: ${defaultModel}\n\n⚠️ IMPORTANT: This provider requires an API key.\n\nAfter the page reloads:\n1. Open Settings (⚙️)\n2. Enter your ${defaultProvider.toUpperCase()} API key\n3. Close Settings to save\n\nThe application will now reload.`);
                } else {
                    // Show success message for local providers
                    await this.showInfoDialog('Success', 
                        `Settings have been reset to defaults.\n\nDefault provider: ${defaultProvider}\nDefault model: ${defaultModel}\n\nThe application will now reload.`);
                }
                
                window.location.reload();
            } else {
                // Show error message
                await this.showInfoDialog('Error', 'Failed to reset settings. Please try again or contact support.');
            }
            
        } catch (error) {
            console.error('[ERROR] - Error during settings reset:', error);
            await this.showInfoDialog('Error', 'An error occurred while resetting settings. Please try again.');
        }
    }

    // Simple dialog replacement for Office Add-in environment
    showConfirmDialog(title, message) {
        return new Promise((resolve) => {
            // Create a simple overlay dialog since Office Add-ins don't support native dialogs
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: rgba(0,0,0,0.5); z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: white; padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: #d73502;">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0;">${message}</p>
                <div style="margin-top: 20px;">
                    <button id="confirm-yes" style="margin-right: 10px; padding: 8px 16px; background: #d73502; color: white; border: none; border-radius: 4px; cursor: pointer;">Reset Settings</button>
                    <button id="confirm-no" style="padding: 8px 16px; background: #ccc; color: black; border: none; border-radius: 4px; cursor: pointer;">Cancel</button>
                </div>
            `;
            
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            
            dialog.querySelector('#confirm-yes').onclick = () => {
                document.body.removeChild(overlay);
                resolve(true);
            };
            
            dialog.querySelector('#confirm-no').onclick = () => {
                document.body.removeChild(overlay);
                resolve(false);
            };
            
            // Close on overlay click
            overlay.onclick = (e) => {
                if (e.target === overlay) {
                    document.body.removeChild(overlay);
                    resolve(false);
                }
            };
        });
    }

    showInfoDialog(title, message) {
        return new Promise((resolve) => {
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: rgba(0,0,0,0.5); z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: white; padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: ${title === 'Error' ? '#d73502' : '#0078d4'};">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0;">${message}</p>
                <div style="margin-top: 20px;">
                    <button id="info-ok" style="padding: 8px 16px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer;">OK</button>
                </div>
            `;
            
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            
            dialog.querySelector('#info-ok').onclick = () => {
                document.body.removeChild(overlay);
                resolve();
            };
            
            // Close on overlay click
            overlay.onclick = (e) => {
                if (e.target === overlay) {
                    document.body.removeChild(overlay);
                    resolve();
                }
            };
        });
    }

    showHelpDialog(title, message, helpUrl) {
        return new Promise((resolve) => {
            // Create a custom dialog with "Open Help" and "Close" buttons
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: rgba(0,0,0,0.5); z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: white; padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: #0078d4;">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0; text-align: left;">${message}</p>
                <p style="margin: 16px 0; font-size: 14px; color: #666;">
                    <strong>Help URL:</strong><br>
                    <a href="${helpUrl}" target="_blank" style="color: #0078d4; word-break: break-all;">${helpUrl}</a>
                </p>
                <div style="margin-top: 20px; display: flex; gap: 10px; justify-content: center;">
                    <button id="help-open" style="padding: 8px 16px; background: #0078d4; color: white; border: none; border-radius: 4px; cursor: pointer;">Open Help</button>
                    <button id="help-close" style="padding: 8px 16px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer;">Close</button>
                </div>
            `;
            
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            
            dialog.querySelector('#help-open').onclick = () => {
                document.body.removeChild(overlay);
                resolve(true);
            };
            
            dialog.querySelector('#help-close').onclick = () => {
                document.body.removeChild(overlay);
                resolve(false);
            };
            
            // Close on overlay click
            overlay.onclick = (e) => {
                if (e.target === overlay) {
                    document.body.removeChild(overlay);
                    resolve(false);
                }
            };
        });
    }

    /**
     * Test connection with current provider settings
     */
    async testConnection() {
        const button = document.getElementById('test-connection');
        const buttonText = button.querySelector('.button-text');
        const buttonSpinner = button.querySelector('.button-spinner');
        
        if (!button || !buttonText || !buttonSpinner) {
            console.error('[ERROR] - Test connection button elements not found');
            return;
        }

        // Get current provider from settings dropdown (not main taskpane dropdown)
        const settingsProviderSelect = document.getElementById('settings-provider-select');
        const currentProvider = settingsProviderSelect?.value;
        if (!currentProvider) {
            this.uiController.showError('Please select a provider to configure first.');
            return;
        }

        // Set button to testing state
        button.disabled = true;
        button.classList.add('testing');
        buttonText.textContent = 'Testing...';
        buttonSpinner.classList.remove('hidden');

        // Declare config outside try block so it's accessible in catch block
        let config = null;

        try {
            // Get current configuration from settings context (not main taskpane)
            config = this.getSettingsAIConfiguration();
            
            // Validate that API key is provided for providers that need it
            if (this.providerNeedsApiKey(currentProvider) && !config.apiKey?.trim()) {
                throw new Error('API key is required for this provider.');
            }

            console.info(`[INFO] - Testing connection for provider: ${currentProvider}`);
            
            // Test the connection using AIService
            const success = await this.aiService.testConnection(config);
            
            if (success) {
                // Success state
                button.classList.remove('testing');
                button.classList.add('success');
                buttonText.textContent = '✓ Success';
                this.uiController.showSuccess(`Connection to ${currentProvider} successful!`);
                
                console.info(`[INFO] - Connection test passed for ${currentProvider}`);
                
                // Log success event
                this.logger.logEvent('connection_test_success', {
                    provider: currentProvider,
                    endpoint: config.endpointUrl,
                    has_api_key: !!config.apiKey
                }, 'Information', this.getUserEmailForTelemetry());
                
                // Reset button after 3 seconds
                setTimeout(() => {
                    button.classList.remove('success');
                    buttonText.textContent = 'Test';
                }, 3000);
                
            } else {
                throw new Error('Connection test failed - no response received');
            }
            
        } catch (error) {
            console.error(`[ERROR] - Connection test failed for ${currentProvider}:`, error);
            
            // Error state
            button.classList.remove('testing');
            button.classList.add('error');
            buttonText.textContent = '✗ Failed';
            
            // Show detailed error message
            let errorMessage = 'Connection test failed. ';
            
            if (error.message.includes('Authentication failed') || error.message.includes('401')) {
                errorMessage += 'Please check your API key.';
            } else if (error.message.includes('Access forbidden') || error.message.includes('403')) {
                errorMessage += 'API key permissions issue. Please verify your key has the correct permissions.';
            } else if (error.message.includes('Service not found') || error.message.includes('404')) {
                errorMessage += 'Service endpoint not found. Please verify your endpoint URL.';
            } else if (error.message.includes('Rate limit') || error.message.includes('429')) {
                errorMessage += 'Rate limit exceeded. Please wait a moment and try again.';
            } else if (error.message.includes('API key is required')) {
                errorMessage += 'Please enter your API key first.';
            } else if (error.message.includes('fetch') || error.message.includes('Network')) {
                errorMessage += 'Network error. Please check your internet connection and endpoint URL.';
            } else {
                errorMessage += error.message || 'Unknown error occurred.';
            }
            
            this.uiController.showError(errorMessage);
            
            // Log failure event
            this.logger.logEvent('connection_test_failed', {
                provider: currentProvider,
                endpoint: config.endpointUrl,
                has_api_key: !!config.apiKey,
                error_message: error.message
            }, 'Error', this.getUserEmailForTelemetry());
            
            // Reset button after 5 seconds
            setTimeout(() => {
                button.classList.remove('error');
                buttonText.textContent = 'Test';
            }, 5000);
            
        } finally {
            // Re-enable button and hide spinner
            button.disabled = false;
            buttonSpinner.classList.add('hidden');
        }
    }

    /**
     * Reset Test Connection button to its default state
     */
    resetTestConnectionButton() {
        const button = document.getElementById('test-connection');
        const buttonText = button?.querySelector('.button-text');
        const buttonSpinner = button?.querySelector('.button-spinner');
        
        if (button && buttonText && buttonSpinner) {
            // Remove all state classes
            button.classList.remove('testing', 'success', 'error');
            
            // Reset button text and state
            buttonText.textContent = 'Test';
            button.disabled = false;
            buttonSpinner.classList.add('hidden');
        }
    }

    async showProviderHelp() {
        const currentProvider = this.modelServiceSelect?.value || 'ollama';
        const providerConfig = this.defaultProvidersConfig?.[currentProvider];
        
        if (providerConfig) {
            const helpText = providerConfig.helpText || 'No help available for this provider.';
            const helpUrl = providerConfig.helpUrl;
            
            if (helpUrl) {
                // Show custom dialog with "Open Help" and "Close" buttons
                const openHelp = await this.showHelpDialog(
                    `Help: ${providerConfig.label || currentProvider}`,
                    helpText,
                    helpUrl
                );
                
                if (openHelp) {
                    // Open the help URL in a new window
                    try {
                        window.open(helpUrl, '_blank', 'noopener,noreferrer');
                    } catch (error) {
                        console.error('[ERROR] - Error opening help URL:', error);
                        await this.showInfoDialog('Error', 'Unable to open help page. Please visit the URL manually.');
                    }
                }
            } else {
                // Show info dialog for providers without URLs
                await this.showInfoDialog(
                    `Help: ${providerConfig.label || currentProvider}`,
                    helpText
                );
            }
        } else {
            await this.showInfoDialog('Help', 'No help available for the current provider.');
        }
    }

    toggleHelpDropdown() {
        const button = document.getElementById('help-dropdown-btn');
        const menu = document.getElementById('help-dropdown-menu');
        
        if (!button || !menu) return;
        
        const isExpanded = button.getAttribute('aria-expanded') === 'true';
        
        if (isExpanded) {
            // Close dropdown
            button.setAttribute('aria-expanded', 'false');
            menu.classList.add('hidden');
        } else {
            // Open dropdown
            button.setAttribute('aria-expanded', 'true');
            menu.classList.remove('hidden');
            
            // Close dropdown when clicking outside
            setTimeout(() => {
                const closeDropdown = (event) => {
                    if (!button.contains(event.target) && !menu.contains(event.target)) {
                        button.setAttribute('aria-expanded', 'false');
                        menu.classList.add('hidden');
                        document.removeEventListener('click', closeDropdown);
                    }
                };
                document.addEventListener('click', closeDropdown);
            }, 0);
        }
    }

    toggleHighContrast(enabled) {
        if (window.debugLog) window.debugLog('[VERBOSE] - toggleHighContrast called:', enabled);
        document.body.classList.toggle('high-contrast', enabled);
        if (window.debugLog) window.debugLog('[VERBOSE] - body classes after toggle:', document.body.classList.toString());
        this.saveSettings();
    }

    toggleScreenReaderMode(enabled) {
        this.accessibilityManager.setScreenReaderMode(enabled);
        this.saveSettings();
    }

    async loadSettingsIntoUI() {
    const settings = this.settingsManager.getSettings();
    // Debug logging
    const debugCheckbox = document.getElementById('debug-logging');
    if (debugCheckbox) debugCheckbox.checked = !!settings['debug-logging'];
    this.logger.setDebugEnabled(!!settings['debug-logging']);
    // Update the global debug function
    window.debugLog = (message, ...args) => {
        if (!!settings['debug-logging']) {
            console.debug(message, ...args);
        }
    };

        // Load form values (excluding provider-specific fields)
        Object.keys(settings).forEach(key => {
            // Skip provider-specific fields 
            if (key === 'api-key' || key === 'endpoint-url' || key === 'provider-configs') return;
            const element = document.getElementById(key);
            if (element) {
                if (element.type === 'checkbox') {
                    element.checked = settings[key];
                } else {
                    element.value = settings[key] || '';
                }
            }
        });

        // If no model-service is set, use configured default (domain filtering happens later when email context is available)
        const modelServiceSelect = document.getElementById('model-service');
        if (modelServiceSelect && (!settings['model-service'] || !modelServiceSelect.value)) {
            if (modelServiceSelect.options.length > 0) {
                // Use configured default provider for now - domain filtering will happen when email is loaded
                const defaultProviders = this.defaultProvidersConfig?._config?.defaultProviders || ['ollama'];
                const defaultProvider = defaultProviders[0];
                
                // Find the default provider in dropdown options
                let selectedOption = null;
                for (let option of modelServiceSelect.options) {
                    if (option.value === defaultProvider) {
                        selectedOption = option.value;
                        break;
                    }
                }
                
                // Fall back to first option if default not found
                const chosenOption = selectedOption || modelServiceSelect.options[0].value;
                modelServiceSelect.value = chosenOption;
                window.debugLog(`[VERBOSE] - TaskPane-${this.instanceId} Setting initial model-service to: ${chosenOption} (domain filtering deferred until email context)`);
                
                // Save this selection to settings
                const updatedSettings = this.settingsManager.getSettings();
                updatedSettings['model-service'] = chosenOption;
                await this.settingsManager.saveSettings(updatedSettings);
                window.debugLog(`[VERBOSE] - TaskPane-${this.instanceId} Saved initial model-service to settings: ${chosenOption}`);
            }
        }

        // Initialize settings provider dropdown selection
        const settingsProviderSelect = document.getElementById('settings-provider-select');
        if (settingsProviderSelect && settingsProviderSelect.options.length > 0) {
            // Get the current provider from main taskpane (after domain filtering and defaults are applied)
            const currentProvider = modelServiceSelect?.value || this.defaultProvidersConfig?._config?.defaultProviders?.[0] || 'ollama';
            
            // Set settings dropdown to match the current provider
            if (currentProvider && Array.from(settingsProviderSelect.options).some(opt => opt.value === currentProvider)) {
                settingsProviderSelect.value = currentProvider;
            } else {
                // Fall back to first option if current provider not found
                settingsProviderSelect.value = settingsProviderSelect.options[0].value;
            }
            
            window.debugLog(`[VERBOSE] - Initialized settings provider dropdown to: ${settingsProviderSelect.value}`);
        }

        // Load provider-specific settings for the current service
        const currentService = settings['model-service'] || this.defaultProvidersConfig?._config?.defaultProvider || 'openai';
        await this.loadProviderSettings(currentService);
        this.updateProviderLabels(currentService);

        // Trigger change events
        if (settings['model-service']) {
            document.getElementById('model-service').dispatchEvent(new Event('change'));
        }

        if (settings['high-contrast']) {
            if (window.debugLog) window.debugLog('[VERBOSE] - Applying high contrast setting on load:', settings['high-contrast']);
            this.toggleHighContrast(true);
        }

        if (settings['screen-reader-mode']) {
            this.toggleScreenReaderMode(true);
        }

        // Load writing style settings
        await this.loadWritingStyleSettings();
    }

    saveSettings() {
        const formSettings = {};
        
        // Collect all form values except provider-specific ones
        const inputs = document.querySelectorAll('input, select, textarea');
        if (window.debugLog) window.debugLog('[VERBOSE] saveSettings: Found', inputs.length, 'form elements');
        
        inputs.forEach((input, index) => {
            if (input.id) {
                // Skip provider-specific fields as they're handled separately
                if (input.id === 'api-key' || input.id === 'endpoint-url') {
                    return;
                }
                
                const value = input.type === 'checkbox' ? input.checked : input.value;
                formSettings[input.id] = value;
                
                if (input.id === 'model-service') {
                    if (window.debugLog) window.debugLog('[VERBOSE] model-service element details:', {
                        index: index,
                        id: input.id,
                        type: input.type,
                        value: input.value,
                        selectedIndex: input.selectedIndex,
                        options: input.options ? Array.from(input.options).map(opt => opt.value) : 'N/A',
                        settingsValue: value
                    });
                }
            }
        });
        
        if (window.debugLog) window.debugLog('[VERBOSE] saveSettings collected:', formSettings);
        
        // Merge form settings with existing settings to preserve provider-configs
        const currentSettings = this.settingsManager.getSettings();
        const mergedSettings = { ...currentSettings, ...formSettings };
        
        this.settingsManager.saveSettings(mergedSettings);
    }

    /**
     * Save current provider's API key and endpoint settings
     * @param {string} provider - The provider key to save settings for
     */
    /**
     * Context-aware provider settings save - determines if we're in settings panel or main taskpane
     */
    async saveProviderSettingsContextAware() {
        // Check if settings panel is currently open
        const settingsPanel = document.getElementById('settings-panel');
        const isSettingsOpen = settingsPanel && !settingsPanel.classList.contains('hidden');
        
        if (isSettingsOpen) {
            // We're in settings context - save to the settings provider
            const settingsProviderSelect = document.getElementById('settings-provider-select');
            const settingsProvider = settingsProviderSelect?.value;
            if (settingsProvider && settingsProvider !== 'undefined') {
                console.log(`[INFO] Saving settings context provider: ${settingsProvider}`);
                
                // IMPORTANT: Only save the fields without cross-provider validation/correction
                // This prevents the base URL from being incorrectly changed when switching between API key and endpoint fields
                await this.saveCurrentProviderSettingsSimple(settingsProvider);
            } else {
                console.warn(`[WARN] Settings panel open but no provider selected in settings dropdown`);
            }
        } else {
            // We're in main taskpane context - save to main provider and update models
            const mainProvider = this.modelServiceSelect?.value;
            if (mainProvider && mainProvider !== 'undefined') {
                console.log(`[INFO] Saving main taskpane provider: ${mainProvider}`);
                await this.saveCurrentProviderSettings(mainProvider);
                // Also trigger model lookup when endpoint or key changes in main context
                this.updateModelDropdown();
            }
        }
    }

    async saveCurrentProviderSettingsSimple(provider) {
        if (!provider || provider === 'undefined') return;
        
        const apiKeyElement = document.getElementById('api-key');
        const endpointUrlElement = document.getElementById('endpoint-url');
        
        const apiKey = apiKeyElement ? apiKeyElement.value.trim() : '';
        const endpointUrl = endpointUrlElement ? endpointUrlElement.value.trim() : '';
        
        // Simple save - no validation or correction, just save what the user entered
        window.debugLog(`[VERBOSE] Simple save for provider ${provider}:`, { 
            apiKey: apiKey.length ? '[HIDDEN]' : '[EMPTY]',
            endpointUrl: endpointUrl
        });
        
        await this.settingsManager.setProviderConfig(provider, apiKey, endpointUrl);
        console.debug(`Simple saved settings for provider ${provider}:`, { apiKey: apiKey ? '[HIDDEN]' : '[EMPTY]', endpointUrl });
    }

    async saveCurrentProviderSettings(provider) {
        if (!provider || provider === 'undefined') return;
        
        const apiKeyElement = document.getElementById('api-key');
        const endpointUrlElement = document.getElementById('endpoint-url');
        
        const apiKey = apiKeyElement ? apiKeyElement.value.trim() : '';
        const endpointUrl = endpointUrlElement ? endpointUrlElement.value.trim() : '';
        
        // Get the default configuration for this provider to validate against
        const defaultConfig = this.defaultProvidersConfig?.[provider];
        let finalApiKey = apiKey;
        let finalEndpointUrl = endpointUrl;
        
        // Validate API key - ensure it matches expected format for this provider
        if (defaultConfig && apiKey) {
            // For providers that should have their own name as API key (like ollama)
            const expectedApiKey = defaultConfig.apiFormat === 'ollama' ? provider : apiKey;
            if (defaultConfig.apiFormat === 'ollama' && apiKey !== provider) {
                console.warn(`[WARN] - Correcting API key for ${provider}: expected '${provider}', got '${apiKey}'`);
                finalApiKey = provider;
            }
        }
        
        // If the endpoint URL matches another provider's default, reset to this provider's default
        if (defaultConfig && endpointUrl) {
            for (const [otherProvider, otherConfig] of Object.entries(this.defaultProvidersConfig)) {
                if (otherProvider !== provider && otherConfig.baseUrl === endpointUrl) {
                    console.warn(`[WARN] - Detected cross-provider endpoint contamination: ${provider} has endpoint from ${otherProvider}. Resetting to correct default.`);
                    finalEndpointUrl = defaultConfig.baseUrl || '';
                    break;
                }
            }
        }
        
        window.debugLog(`[VERBOSE] Saving settings for provider ${provider}:`, { 
            originalApiKey: apiKey.length ? '[HIDDEN]' : '[EMPTY]',
            originalEndpointUrl: endpointUrl,
            finalApiKey: finalApiKey.length ? '[HIDDEN]' : '[EMPTY]',
            finalEndpointUrl,
            wasApiKeyCorrected: finalApiKey !== apiKey,
            wasEndpointCorrected: finalEndpointUrl !== endpointUrl
        });
        
        await this.settingsManager.setProviderConfig(provider, finalApiKey, finalEndpointUrl);
        console.debug(`Saved settings for provider ${provider}:`, { apiKey: finalApiKey ? '[HIDDEN]' : '[EMPTY]', endpointUrl: finalEndpointUrl });
    }

    /**
     * Load provider-specific settings into the UI
     * @param {string} provider - The provider key to load settings for
     */
    /**
     * Load provider settings for display in settings panel only - does not affect main taskpane state
     * @param {string} provider - The provider key to load settings for
     */
    async loadSettingsOnlyProviderConfig(provider) {
        if (!provider || provider === 'undefined') return;
        
        const providerConfig = this.settingsManager.getProviderConfig(provider);
        window.debugLog(`[VERBOSE] - Loading settings display for ${provider}:`, providerConfig);
        
        const apiKeyElement = document.getElementById('api-key');
        const endpointUrlElement = document.getElementById('endpoint-url');
        
        if (apiKeyElement) {
            let apiKeyToUse = providerConfig['api-key'] || '';
            
            // For display purposes, show the current stored value or default for Ollama
            const defaultConfig = this.defaultProvidersConfig?.[provider];
            if (defaultConfig && defaultConfig.apiFormat === 'ollama' && !apiKeyToUse) {
                apiKeyToUse = provider; // Show default for Ollama but don't save it
            }
            
            apiKeyElement.value = apiKeyToUse;
            window.debugLog(`[VERBOSE] - Displayed API key for ${provider}: ${apiKeyElement.value ? '[HIDDEN]' : 'EMPTY'}`);
        }
        
        if (endpointUrlElement) {
            let endpointToUse = providerConfig['endpoint-url'] || '';
            
            // For display, show stored value or default if none exists
            if (!endpointToUse && this.defaultProvidersConfig && this.defaultProvidersConfig[provider]) {
                endpointToUse = this.defaultProvidersConfig[provider].baseUrl || '';
            }
            
            endpointUrlElement.value = endpointToUse;
            window.debugLog(`[VERBOSE] - Displayed endpoint URL for ${provider}: ${endpointToUse}`);
        }
        
        console.debug(`Loaded settings display for provider ${provider} (no state changes):`, { 
            apiKey: providerConfig['api-key'] ? '[HIDDEN]' : '', 
            endpointUrl: endpointUrlElement ? endpointUrlElement.value : 'no element'
        });
    }

    async loadProviderSettings(provider) {
        if (!provider || provider === 'undefined') return;
        
        // Set flag to prevent blur events during loading
        this.isLoadingProviderSettings = true;
        
        try {
            const providerConfig = this.settingsManager.getProviderConfig(provider);
            window.debugLog(`[VERBOSE] - Loading provider settings for ${provider}:`, providerConfig);
            
            const apiKeyElement = document.getElementById('api-key');
            const endpointUrlElement = document.getElementById('endpoint-url');
            
            let settingsWereCorrected = false;
        
        if (apiKeyElement) {
            let apiKeyToUse = providerConfig['api-key'] || '';
            
            // Validate and correct API key if needed
            const defaultConfig = this.defaultProvidersConfig?.[provider];
            if (defaultConfig && defaultConfig.apiFormat === 'ollama') {
                // For Ollama, API key should be the provider name
                if (!apiKeyToUse || apiKeyToUse !== provider) {
                    apiKeyToUse = provider;
                    settingsWereCorrected = true;
                    window.debugLog(`[VERBOSE] - Corrected API key for ${provider} to: ${apiKeyToUse}`);
                }
            }
            
            apiKeyElement.value = apiKeyToUse;
            window.debugLog(`[VERBOSE] - Set API key field to: ${apiKeyElement.value ? '[HIDDEN]' : 'EMPTY'}`);
        }
        
        if (endpointUrlElement) {
            // Determine which endpoint to use
            let endpointToUse = providerConfig['endpoint-url'] || '';
            
            if (this.defaultProvidersConfig && this.defaultProvidersConfig[provider]) {
                const defaultEndpoint = this.defaultProvidersConfig[provider].baseUrl || '';
                
                // Validate endpoint URL and correct if contaminated
                if (endpointToUse && defaultEndpoint) {
                    // Check if current endpoint belongs to a different provider
                    for (const [otherProvider, otherConfig] of Object.entries(this.defaultProvidersConfig)) {
                        if (otherProvider !== provider && otherConfig.baseUrl === endpointToUse) {
                            console.warn(`[WARN] - Provider ${provider} has contaminated endpoint from ${otherProvider}. Correcting to default.`);
                            endpointToUse = defaultEndpoint;
                            settingsWereCorrected = true;
                            break;
                        }
                    }
                }
                
                // Additional check: if no endpoint set or contaminated, use default
                if (!endpointToUse && defaultEndpoint) {
                    endpointToUse = defaultEndpoint;
                    settingsWereCorrected = true;
                    window.debugLog(`[VERBOSE] - Set missing endpoint for ${provider} to default: ${defaultEndpoint}`);
                }
                
                // For onsite providers, check if stored endpoint is the old incorrect OpenAI URL
                if (provider.startsWith('onsite') && defaultEndpoint) {
                    // If stored endpoint is the old OpenAI URL, replace it with the correct baseUrl
                    if (endpointToUse === 'https://api.openai.com/v1') {
                        endpointToUse = defaultEndpoint;
                        settingsWereCorrected = true;
                    } else if (!endpointToUse) {
                        // If no stored endpoint, use the baseUrl from ai-providers.json
                        endpointToUse = defaultEndpoint;
                    }
                    // Otherwise keep the user's custom endpoint
                } else if (!endpointToUse && defaultEndpoint) {
                    // For other providers, use default only if no stored endpoint
                    endpointToUse = defaultEndpoint;
                }
            }
            
            endpointUrlElement.value = endpointToUse;
            window.debugLog(`[VERBOSE] - Set endpoint URL to: ${endpointToUse}`);
            
            // Additional verification: check if the UI field value matches what we just set
            setTimeout(() => {
                if (endpointUrlElement.value !== endpointToUse) {
                    console.warn(`[WARN] - Endpoint URL contamination detected! Expected: ${endpointToUse}, but UI shows: ${endpointUrlElement.value}. Force-correcting...`);
                    endpointUrlElement.value = endpointToUse;
                }
            }, 50);
        }
        
        // If we corrected any contaminated settings, save them immediately
        if (settingsWereCorrected) {
            console.info(`[INFO] - Corrected contaminated settings for ${provider}, saving to persist fixes`);
            await this.saveCurrentProviderSettings(provider);
        }
        
            console.debug(`Loaded settings for provider ${provider}:`, { 
                apiKey: providerConfig['api-key'] ? '[HIDDEN]' : '', 
                endpointUrl: endpointUrlElement ? endpointUrlElement.value : 'no element',
                settingsWereCorrected
            });
        } finally {
            // Clear the loading flag
            this.isLoadingProviderSettings = false;
        }
    }

    /**
     * Update provider labels in the UI to show which provider is currently selected
     * @param {string} provider - The current provider key
     */
    updateProviderLabels(provider) {
        const providerLabel = this.defaultProvidersConfig?.[provider]?.label || provider;
        
        const apiKeyLabel = document.getElementById('api-key-provider-label');
        const endpointUrlLabel = document.getElementById('endpoint-url-provider-label');
        
        if (apiKeyLabel) {
            apiKeyLabel.textContent = `(${providerLabel})`;
        }
        
        if (endpointUrlLabel) {
            endpointUrlLabel.textContent = `(${providerLabel})`;
        }
    }

    getUserId() {
        // Use the logger's consistent user context
        return this.logger?.getUserContext()?.userId || 'unknown';
    }

    /**
     * Get the current Outlook user's email address for telemetry context
     * @returns {string|null} Current user's email address or null
     */
    getUserEmailForTelemetry() {
        // For telemetry, we always want the actual user's email (the person using the add-in),
        // not the recipient's email, regardless of whether it's sent or inbox mail
        try {
            const userProfile = Office.context.mailbox.userProfile;
            return userProfile ? userProfile.emailAddress : null;
        } catch (error) {
            console.warn('[WARN] - Unable to get user profile for telemetry:', error);
            return null;
        }
    }

    getEmailIdentifiersForTelemetry() {
        if (!this.currentEmail) {
            return null;
        }

        // Create a hash of the subject for correlation without revealing content
        const subjectHash = this.currentEmail.subject ? 
            this.hashString(this.currentEmail.subject) : null;

        // Get available Office.js identifiers that don't reveal content
        const identifiers = {
            // Primary identifiers for email tracking
            conversationId: this.currentEmail.conversationId || null,
            itemId: this.currentEmail.itemId || null,
            itemClass: this.currentEmail.itemClass || null,
            
            // Content-safe metadata
            subjectHash: subjectHash,
            bodyLength: this.currentEmail.bodyLength || 0,
            hasAttachments: this.currentEmail.hasAttachments || false,
            hasInternetMessageId: this.currentEmail.hasInternetMessageId || false,
            
            // Email context
            itemType: this.currentEmail.itemType || null,
            isReply: this.currentEmail.isReply || false,
            date: this.currentEmail.date?.toISOString() || null
        };

        // Try to get additional identifiers if available from Office context
        try {
            if (Office.context.mailbox.item) {
                // Add any additional runtime identifiers
                if (Office.context.mailbox.item.itemId && !identifiers.itemId) {
                    identifiers.itemId = Office.context.mailbox.item.itemId;
                }
            }
        } catch (error) {
            console.debug('Could not access additional Office identifiers:', error);
        }

        return identifiers;
    }

    hashString(str) {
        // Simple hash function for subject correlation without revealing content
        let hash = 0;
        for (let i = 0; i < str.length; i++) {
            const char = str.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        return hash.toString(36); // Return as base-36 string
    }
}

// Initialize the application when Office.js is ready
Office.onReady(() => {
    const app = new TaskpaneApp();
    app.initialize().catch(error => {
        console.error('[ERROR] - Failed to initialize application:', error);
    });
});
