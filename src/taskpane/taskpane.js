// PromptEmail Taskpane JavaScript
// Main application logic for the email analysis interface

import '../assets/css/taskpane.css';
import { EmailAnalyzer } from '../services/EmailAnalyzer';
import { AIService } from '../services/AIService';
import { ClassificationDetector } from '../services/ClassificationDetector';
import { Logger } from '../services/Logger';
import { SettingsManager } from '../services/SettingsManager';
import { UIStateManager } from '../services/UIStateManager';
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
            console.warn('Could not load ai-providers.json:', e);
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
            console.warn('Could not load taskpane-resources.json:', e);
            return {};
        }
    }

    /**
     * Dynamically populate the resources dropdown from taskpane-resources.json
     */
    populateResourcesDropdown() {
        const dropdownMenu = document.getElementById('help-dropdown-menu');
        if (!dropdownMenu) {
            console.warn('Could not find help-dropdown-menu element');
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
            console.error(`Error opening ${resource.name}:`, error);
            
            // Fallback: copy URL to clipboard if available
            if (navigator.clipboard && navigator.clipboard.writeText) {
                try {
                    await navigator.clipboard.writeText(resource.url);
                    this.showInfoDialog(resource.name, 
                        `Could not open ${resource.name}. The URL has been copied to your clipboard:\n\n${resource.url}`);
                } catch (clipboardError) {
                    console.error('Could not copy to clipboard:', clipboardError);
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
            return { defaultProvider, allowedProviders };
        }

        // Use default providers for unmapped domains
        const allowedProviders = config.defaultProviders || ['ollama'];
        const defaultProvider = allowedProviders[0];
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
    }

    /**
     * Check if a provider is currently available
     * @param {string} providerId - The provider ID to check
     * @returns {boolean} True if the provider is available
     */
    async isProviderAvailable(providerId) {
        try {
            // For now, we'll do a simple check based on provider type
            // This could be expanded to do actual connectivity testing
            
            if (providerId === 'ollama') {
                // Check if ollama is running on localhost:11434
                try {
                    const response = await fetch('http://localhost:11434/api/tags', {
                        method: 'GET',
                        signal: AbortSignal.timeout(2000) // 2 second timeout
                    });
                    return response.ok;
                } catch (error) {
                    return false;
                }
            }
            
            // For other providers (bedrock, onsite, etc.), assume available if configured
            // Could add more specific checks here in the future
            return true;
            
        } catch (error) {
            return false;
        }
    }

    /**
     * Get the default provider from ai-providers.json configuration
     * @returns {string} The default provider ID
     */
    getDefaultProvider() {
        if (!this.defaultProvidersConfig || !this.defaultProvidersConfig._config) {
            return 'ollama'; // Fallback if config not loaded
        }
        return this.defaultProvidersConfig._config.defaultProvider || 'ollama';
    }

    /**
     * Apply domain-based provider filtering when user context is available
     */
    async applyDomainBasedProviderFiltering() {
        try {
            // Skip domain filtering if already loading provider settings to prevent race conditions
            if (this.isLoadingProviderSettings) {
                return;
            }
            
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
                console.warn('No user email available in email context, skipping domain filtering');
                return;
            }
            
            const providerConfig = this.getProvidersByDomain(userProfile);
            const { defaultProvider, allowedProviders } = providerConfig;
            

            
            // Filter dropdown to only show allowed providers
            this.filterProviderDropdown(allowedProviders);
            
            // Check if current selection needs to be changed
            const modelServiceSelect = document.getElementById('model-service');
            const currentSelection = modelServiceSelect?.value;
            

            
            // Check if user has explicitly chosen a different provider for this domain
            const settings = this.settingsManager.getSettings();
            const domainChoice = settings[`domain-choice-${userEmail.split('@')[1]}`];
            

            
            if (!currentSelection || !allowedProviders.includes(currentSelection)) {
                // Switch to domain default if:
                // 1. No current selection (first run/new profile), OR
                // 2. Current selection is not allowed for this domain
                const reason = !currentSelection ? 'no provider selected' : `current provider '${currentSelection}' not allowed for domain`;

                
                if (modelServiceSelect) {
                    modelServiceSelect.value = defaultProvider;
                    
                    // Save the new selection and mark domain choice
                    settings['model-service'] = defaultProvider;
                    settings[`domain-choice-${userEmail.split('@')[1]}`] = defaultProvider;
                    await this.settingsManager.saveSettings(settings);
                    
                    // Load provider settings for the new provider
                    await this.loadProviderSettings(defaultProvider);
                    
                    // Also update settings panel to prevent endpoint contamination
                    const settingsProviderSelect = document.getElementById('settings-provider-select');
                    if (settingsProviderSelect) {
                        settingsProviderSelect.value = defaultProvider;
                        await this.loadSettingsOnlyProviderConfig(defaultProvider);
                    }
                    
                    // Trigger change event to update related UI
                    modelServiceSelect.dispatchEvent(new Event('change'));
                }
            } else if (!domainChoice && currentSelection !== defaultProvider) {
                // First time seeing this domain and current selection is not the domain default
                // Switch to domain default for proper governance

                
                if (modelServiceSelect) {
                    modelServiceSelect.value = defaultProvider;
                    
                    // Save the new selection and mark domain choice
                    settings['model-service'] = defaultProvider;
                    settings[`domain-choice-${userEmail.split('@')[1]}`] = defaultProvider;
                    await this.settingsManager.saveSettings(settings);
                    
                    // Load provider settings for the new provider
                    await this.loadProviderSettings(defaultProvider);
                    
                    // Also update settings panel to prevent endpoint contamination
                    const settingsProviderSelect = document.getElementById('settings-provider-select');
                    if (settingsProviderSelect) {
                        settingsProviderSelect.value = defaultProvider;
                        await this.loadSettingsOnlyProviderConfig(defaultProvider);
                    }
                    
                    // Trigger change event to update related UI
                    modelServiceSelect.dispatchEvent(new Event('change'));
                }
            } else {
                // User has made a choice for this domain before
                if (domainChoice && allowedProviders.includes(domainChoice)) {
                    // Check if the domain choice provider is actually available
                    const isDomainChoiceAvailable = await this.isProviderAvailable(domainChoice);
                    
                    if (!isDomainChoiceAvailable) {
                        // Domain choice is configured but not available - fall back to domain default

                        if (modelServiceSelect) {
                            modelServiceSelect.value = defaultProvider;
                            settings['model-service'] = defaultProvider;
                            // Don't change domain-choice - keep user's preference for when it becomes available again
                            await this.settingsManager.saveSettings(settings);
                            await this.loadProviderSettings(defaultProvider);
                            
                            // Also update settings panel
                            const settingsProviderSelect = document.getElementById('settings-provider-select');
                            if (settingsProviderSelect) {
                                settingsProviderSelect.value = defaultProvider;
                                await this.loadSettingsOnlyProviderConfig(defaultProvider);
                            }
                            
                            modelServiceSelect.dispatchEvent(new Event('change'));
                        }
                        return;
                    }
                    
                    // Check if the current selection matches the domain choice
                    if (currentSelection !== domainChoice) {
                        // The current UI selection differs from stored domain choice
                        // If current selection is valid for this domain, respect it and update preferences
                        if (allowedProviders.includes(currentSelection)) {
                            settings['model-service'] = currentSelection;
                            settings[`domain-choice-${userEmail.split('@')[1]}`] = currentSelection;
                            settings['user-active-provider-choice'] = currentSelection;
                            await this.settingsManager.saveSettings(settings);
                            // Load provider settings for the current selection to ensure clean configuration
                            await this.loadProviderSettings(currentSelection);
                            
                            // Also update settings panel to prevent endpoint contamination
                            const settingsProviderSelect = document.getElementById('settings-provider-select');
                            if (settingsProviderSelect) {
                                settingsProviderSelect.value = currentSelection;
                                await this.loadSettingsOnlyProviderConfig(currentSelection);
                            }
                            
                            // Also update the UI dropdown to ensure consistency
                            if (modelServiceSelect && modelServiceSelect.value !== currentSelection) {
                                modelServiceSelect.value = currentSelection;
                            }

                        } else {
                            // Current selection is not allowed for this domain - use domain default

                            if (modelServiceSelect) {
                                modelServiceSelect.value = domainChoice;
                                settings['model-service'] = domainChoice;
                                await this.settingsManager.saveSettings(settings);
                                await this.loadProviderSettings(domainChoice);
                                
                                // Also update settings panel to prevent endpoint contamination
                                const settingsProviderSelect = document.getElementById('settings-provider-select');
                                if (settingsProviderSelect) {
                                    settingsProviderSelect.value = domainChoice;
                                    await this.loadSettingsOnlyProviderConfig(domainChoice);
                                }
                                
                                // Don't dispatch change event - we already loaded the provider settings
                                // This prevents double-loading and race conditions

                            }
                        }
                    } else {
                        // Current selection matches domain choice - all good

                    }
                } else {
                    // No valid domain choice stored, or stored choice is no longer allowed
                    if (allowedProviders.includes(currentSelection)) {
                        // Current selection is valid - make it the new domain choice

                        settings['model-service'] = currentSelection;
                        settings[`domain-choice-${userEmail.split('@')[1]}`] = currentSelection;
                        settings['user-active-provider-choice'] = currentSelection;
                        await this.settingsManager.saveSettings(settings);
                        await this.loadProviderSettings(currentSelection);
                        // Refresh models to ensure they match the current provider
                        await this.updateModelDropdown();
                    } else {
                        // Current selection is not valid - use domain default

                        if (modelServiceSelect) {
                            modelServiceSelect.value = defaultProvider;
                            settings['model-service'] = defaultProvider;
                            settings[`domain-choice-${userEmail.split('@')[1]}`] = defaultProvider;
                            await this.settingsManager.saveSettings(settings);
                            await this.loadProviderSettings(defaultProvider);
                            modelServiceSelect.dispatchEvent(new Event('change'));
                            // The change event will trigger updateModelDropdown, so no need to call it explicitly here
                        }
                    }
                }
            }
            
        } catch (error) {
            // Domain filtering not critical if it fails
        }
    }

    showAnalysisSection() {
        // Show the analysis section
        const analysisSection = document.getElementById('analysis-section');
        if (analysisSection) {
            analysisSection.classList.remove('hidden');

        }
    }
    
    /**
     * Shows a notification to the user about email truncation or other status updates
     * @param {string} message - The message to display
     * @param {string} type - Type of notification: 'info', 'warning', 'error', 'success'
     * @param {number} duration - How long to show notification in milliseconds (0 = persistent)
     */
    showNotification(message, type = 'info', duration = 5000) {
        const statusContainer = document.getElementById('status-messages');
        if (!statusContainer) {
            console.warn('Status messages container not found');
            return;
        }
        
        // Create notification element using existing CSS classes
        const notification = document.createElement('div');
        notification.className = `status-message status-${type}`;
        notification.setAttribute('role', 'alert');
        
        notification.innerHTML = `
            <div class="status-content">
                <span class="status-icon" aria-hidden="true"></span>
                <span class="status-text">${message}</span>
                <button class="status-close" type="button" aria-label="Close notification" title="Close">×</button>
            </div>
        `;
        
        // Add to container
        statusContainer.appendChild(notification);
        
        // Add close functionality
        const closeBtn = notification.querySelector('.status-close');
        closeBtn.addEventListener('click', () => {
            notification.remove();
        });
        
        // Auto-remove after duration (if specified)
        if (duration > 0) {
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.remove();
                }
            }, duration);
        }
        

    }
    
    /**
     * Shows a notification specifically for email truncation events
     * @param {Object} truncationInfo - Information about the truncation that occurred
     */
    showEmailTruncationNotification(truncationInfo) {
        if (!truncationInfo || !truncationInfo.wasTruncated) {
            return;
        }
        
        const originalKB = Math.round(truncationInfo.originalLength / 1024);
        const truncatedKB = Math.round(truncationInfo.truncatedLength / 1024);
        const removedKB = originalKB - truncatedKB;
        
        const title = 'Email Content Shortened';
        const message = `📧 Email content was shortened for AI processing (${originalKB}KB → ${truncatedKB}KB, ${removedKB}KB removed). Key portions at the beginning and end were preserved.`;
        
        // Show modal popup for immediate attention
        this.showModalNotification(title, message, 'info');
        
        // Add to permanent notification area
        this.addPermanentNotification('📧', title, message, 'truncation');
        
        // Also show temporary notification for those who prefer it
        this.showNotification(message, 'info', 8000); // Show for 8 seconds
        

    }
    
    /**
     * Shows a notification specifically for HTML-to-text conversion events
     * @param {Object} conversionInfo - Information about the HTML conversion that occurred
     */
    showHtmlConversionNotification(conversionInfo) {
        // Disabled - no longer show HTML conversion notifications to user
        return;
        

    }

    /**
     * Shows a modal popup notification that requires user interaction
     * @param {string} title - Title of the notification
     * @param {string} message - Message content
     * @param {string} type - Type of notification (info, success, warning, error)
     * @param {Object} customButtons - Optional custom button configuration
     */
    showModalNotification(title, message, type = 'info', customButtons = null) {
        const modal = document.getElementById('notification-modal');
        const modalTitle = document.getElementById('modal-title');
        const modalMessage = document.getElementById('modal-message');
        const modalFooter = modal.querySelector('.modal-footer');
        
        if (!modal || !modalTitle || !modalMessage || !modalFooter) {
            console.warn('Modal elements not found');
            return;
        }
        
        // Set content
        modalTitle.textContent = title;
        modalMessage.innerHTML = message;
        
        // Add type-specific styling
        modal.className = `modal-overlay modal-${type}`;
        modal.setAttribute('aria-hidden', 'false');
        
        // Handle custom buttons if provided
        if (customButtons) {
            modalFooter.innerHTML = customButtons.map(btn => 
                `<button id="${btn.id}" class="btn ${btn.class || 'btn-secondary'}">${btn.text}</button>`
            ).join('');
            
            // Add event listeners for custom buttons
            customButtons.forEach(btn => {
                const buttonElement = document.getElementById(btn.id);
                if (buttonElement && btn.action) {
                    buttonElement.addEventListener('click', btn.action);
                }
            });
        } else {
            // Reset to default buttons
            modalFooter.innerHTML = `
                <button id="modal-ok" class="btn btn-primary">OK</button>
                <button id="modal-show-permanent" class="btn btn-secondary">Keep Visible</button>
            `;
            
            // Re-attach default event listeners
            const modalOk = document.getElementById('modal-ok');
            const modalShowPermanent = document.getElementById('modal-show-permanent');
            
            if (modalOk) {
                modalOk.addEventListener('click', () => this.hideModalNotification());
            }
            
            if (modalShowPermanent) {
                modalShowPermanent.addEventListener('click', () => {
                    this.hideModalNotification();
                    this.showPermanentNotifications();
                });
            }
        }
        
        // Show modal
        modal.classList.remove('hidden');
        
        // Focus management - focus first button
        const firstButton = modalFooter.querySelector('button');
        if (firstButton) {
            firstButton.focus();
        }
        

    }

    /**
     * Shows the early access notice modal if user hasn't acknowledged it
     * @returns {boolean} True if the modal was shown, false otherwise
     */
    async checkAndShowEarlyAccessNotice() {
        const settings = this.settingsManager.getSettings();
        const hasAcknowledged = settings['early-access-acknowledged'];
        
        if (!hasAcknowledged) {
            this.showEarlyAccessModal();
            return true; // Modal was shown
        }
        
        return false; // Modal was not shown
    }

    /**
     * Shows the early access notice modal
     */
    showEarlyAccessModal() {
        const modal = document.getElementById('early-access-modal');
        const acknowledgeBtn = document.getElementById('early-access-acknowledge');
        const dontShowAgainCheckbox = document.getElementById('dont-show-again-checkbox');
        
        if (!modal || !acknowledgeBtn || !dontShowAgainCheckbox) {
            console.warn('Early access modal elements not found');
            return;
        }

        // Load content from configuration
        this.loadEarlyAccessContent();

        // Show modal
        modal.classList.remove('hidden');
        modal.setAttribute('aria-hidden', 'false');

        // Focus the acknowledge button
        acknowledgeBtn.focus();

        // Handle acknowledge button click
        const handleAcknowledge = async () => {
            const dontShowAgain = dontShowAgainCheckbox.checked;
            
            if (dontShowAgain) {
                // Save setting to not show again
                this.settingsManager.setSetting('early-access-acknowledged', true);
            }
            
            // Hide modal
            modal.classList.add('hidden');
            modal.setAttribute('aria-hidden', 'true');
            
            // Remove event listeners
            acknowledgeBtn.removeEventListener('click', handleAcknowledge);
            document.removeEventListener('keydown', handleEscapeKey);
            
            // After Early Access modal is dismissed, check for initial setup
            setTimeout(async () => {
                await this.checkForInitialSetupNeeded();
            }, 100); // Small delay to ensure modal is fully hidden
        };

        // Handle escape key
        const handleEscapeKey = (event) => {
            if (event.key === 'Escape') {
                handleAcknowledge();
            }
        };

        // Add event listeners
        acknowledgeBtn.addEventListener('click', handleAcknowledge);
        document.addEventListener('keydown', handleEscapeKey);
    }

    /**
     * Load early access notice content from configuration
     */
    loadEarlyAccessContent() {
        const earlyAccessConfig = this.taskpaneResourcesConfig?.earlyAccess;
        
        if (!earlyAccessConfig) {
            console.warn('Early access configuration not found');
            return;
        }

        // Update title
        const titleElement = document.getElementById('early-access-title');
        if (titleElement) {
            titleElement.textContent = earlyAccessConfig.title || '⚠️ Early Access Notice';
        }

        // Update content
        const contentElement = document.getElementById('early-access-content');
        if (contentElement) {
            contentElement.innerHTML = earlyAccessConfig.fullMessage || earlyAccessConfig.shortMessage || 'This is an MVP add-in for early evaluation.';
        }
    }

    /**
     * Adds a notification to the permanent notification area
     * @param {string} icon - Icon for the notification
     * @param {string} title - Title of the notification
     * @param {string} message - Message content
     * @param {string} category - Category for grouping (truncation, conversion, etc.)
     */
    addPermanentNotification(icon, title, message, category) {
        const permanentArea = document.getElementById('permanent-notifications');
        const notificationList = document.getElementById('notification-list');
        
        if (!permanentArea || !notificationList) {
            console.warn('Permanent notification elements not found');
            return;
        }
        
        // Show permanent area if hidden
        permanentArea.classList.remove('hidden');
        
        // Create notification element
        const notification = document.createElement('div');
        notification.className = `permanent-notification notification-${category}`;
        
        const timestamp = new Date().toLocaleTimeString();
        
        notification.innerHTML = `
            <div class="notification-icon">${icon}</div>
            <div class="notification-content">
                <div class="notification-title">${title}</div>
                <div class="notification-text">${message}</div>
                <div class="notification-time">${timestamp}</div>
            </div>
        `;
        
        // Add to top of list
        notificationList.insertBefore(notification, notificationList.firstChild);
        
        // Limit to 10 notifications
        const notifications = notificationList.children;
        if (notifications.length > 10) {
            notificationList.removeChild(notifications[notifications.length - 1]);
        }
        

    }

    /**
     * Hides the modal notification
     */
    hideModalNotification() {
        const modal = document.getElementById('notification-modal');
        if (modal) {
            modal.classList.add('hidden');
            modal.setAttribute('aria-hidden', 'true');
            
            if (window.debugLog) {
                window.debugLog('Modal notification hidden');
            }
        }
    }

    /**
     * Shows the permanent notifications area
     */
    showPermanentNotifications() {
        const permanentArea = document.getElementById('permanent-notifications');
        if (permanentArea) {
            permanentArea.classList.remove('hidden');
            permanentArea.scrollIntoView({ behavior: 'smooth' });
            
            if (window.debugLog) {
                window.debugLog('Permanent area shown');
            }
        }
    }

    /**
     * Clears all permanent notifications
     */
    clearPermanentNotifications() {
        const notificationList = document.getElementById('notification-list');
        const permanentArea = document.getElementById('permanent-notifications');
        
        if (notificationList) {
            notificationList.innerHTML = '';
        }
        
        if (permanentArea) {
            permanentArea.classList.add('hidden');
        }
        
        if (window.debugLog) {
            window.debugLog('Cleared all permanent notifications');
        }
    }

    /**
     * Clears all notifications from the status area
     */
    clearNotifications() {
        const statusContainer = document.getElementById('status-messages');
        if (statusContainer) {
            statusContainer.innerHTML = '';
        }
    }
    
    switchToAnalysisTab() {
        // Switch to the analysis tab in the UI
        const analysisTabButton = document.querySelector('.tab-button[aria-controls="panel-analysis"]');
        if (analysisTabButton) {

            analysisTabButton.click();
        } else {
            console.error('Analysis tab button not found');
        }
    }

    clearAnalysisAndResponse() {
        // Clear analysis results
        const analysisContainer = document.getElementById('email-analysis');
        if (analysisContainer) {
            analysisContainer.innerHTML = '';

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

        
    this.settingsManager = new SettingsManager();
    this.uiStateManager = new UIStateManager(this.settingsManager);
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
        // Track when user navigates away or closes the taskpane/Outlook
        window.addEventListener('beforeunload', () => {
            this.logSessionSummary();
        });
        
        // Note: Removed blur event to only capture actual closure events
        // Session summary will now only trigger when:
        // - User closes the taskpane
        // - User navigates away from the page
        // - Outlook is closed
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
            const highContrastValue = !!currentSettings['high-contrast'];
            if (window.debugLog) window.debugLog('Loading high-contrast setting:', currentSettings['high-contrast'], '→', highContrastValue);
            // Always apply the high-contrast setting, whether true or false (skip save during load)
            this.toggleHighContrast(highContrastValue, true);
            
            // Load provider config before UI setup
            this.defaultProvidersConfig = await this.fetchDefaultProvidersConfig();
            
            // Load taskpane resources config
            this.taskpaneResourcesConfig = await this.fetchTaskpaneResourcesConfig();
            
            // Update AIService with provider configuration
            this.aiService.updateProvidersConfig(this.defaultProvidersConfig);
            
            // Setup UI
            await this.setupUI();
            
            // Initialize UI state management
            await this.uiStateManager.initializeFromStorage();
            
            // Populate resources dropdown after UI is set up
            this.populateResourcesDropdown();
            
            // Setup accessibility
            this.accessibilityManager.initialize();
            
            // Initialize Splunk telemetry if enabled
            await this.initializeTelemetry();
            
            // Load current email
            await this.loadCurrentEmail();
            
            // Check and show early access notice first (before other modals)
            const earlyAccessShown = await this.checkAndShowEarlyAccessNotice();
            
            // Only check for initial setup if early access notice wasn't shown
            // If early access was shown, setup check will happen after it's dismissed
            if (!earlyAccessShown) {
                await this.checkForInitialSetupNeeded();
            }
            
            // Try automatic analysis if conditions are met
            await this.attemptAutoAnalysis();
            
            // Hide loading, show main content
            this.uiController.hideLoading();
            this.uiController.showMainContent();
            
            // Log session start
            this.logger.logEvent('session_start', {
            });
            
        } catch (error) {
            console.error('Failed to initialize TaskpaneApp:', error);
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

        
        try {
            // Logger already initialized telemetry config in constructor, just check if it's ready
            // If not initialized yet, wait for it
            if (!this.logger.telemetryConfig) {

                // Give it a moment to load
                await new Promise(resolve => setTimeout(resolve, 100));
            }
            
            // Start telemetry auto-flush if enabled
            if (this.logger.telemetryConfig?.telemetry?.enabled) {
                const provider = this.logger.telemetryConfig.telemetry.provider;
                if (provider === 'api_gateway') {
                    this.logger.startApiGatewayAutoFlush();

                }
            }
            
        } catch (error) {
            console.error('Failed to initialize telemetry:', error);
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

        // Modal notification event listeners
        const modalClose = document.getElementById('modal-close');
        const modalOk = document.getElementById('modal-ok');
        const modalShowPermanent = document.getElementById('modal-show-permanent');
        const modalOverlay = document.getElementById('notification-modal');
        
        if (modalClose) {
            modalClose.addEventListener('click', () => this.hideModalNotification());
        }
        
        if (modalOk) {
            modalOk.addEventListener('click', () => this.hideModalNotification());
        }
        
        if (modalShowPermanent) {
            modalShowPermanent.addEventListener('click', () => {
                this.hideModalNotification();
                this.showPermanentNotifications();
            });
        }
        
        // Close modal when clicking outside
        if (modalOverlay) {
            modalOverlay.addEventListener('click', (e) => {
                if (e.target === modalOverlay) {
                    this.hideModalNotification();
                }
            });
        }

        // Permanent notification event listeners
        const clearNotifications = document.getElementById('clear-notifications');
        if (clearNotifications) {
            clearNotifications.addEventListener('click', () => this.clearPermanentNotifications());
        }

        // Keyboard accessibility for modal
        document.addEventListener('keydown', (e) => {
            const modal = document.getElementById('notification-modal');
            if (modal && !modal.classList.contains('hidden')) {
                if (e.key === 'Escape') {
                    this.hideModalNotification();
                }
            }
        });
        
        // Model service change
        // Allow provider changes during session, but don't persist choice
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
                await this.settingsManager.saveSettings(settings, 'setting: debug-logging');
                this.logger.setDebugEnabled(enabled);
                // Update the global debug function
                window.debugLog = (message, ...args) => {
                    if (enabled) {
                        console.debug(message, ...args);
                    }
                };
            });
        }

        // Early access acknowledged checkbox
        const earlyAccessCheckbox = document.getElementById('early-access-acknowledged');
        if (earlyAccessCheckbox) {
            earlyAccessCheckbox.addEventListener('change', async (e) => {
                const acknowledged = e.target.checked;
                const settings = this.settingsManager.getSettings();
                settings['early-access-acknowledged'] = acknowledged;
                await this.settingsManager.saveSettings(settings, 'setting: early-access-acknowledged');
            });
        }

        // Show early access notice button
        const showEarlyAccessBtn = document.getElementById('show-early-access-notice');
        if (showEarlyAccessBtn) {
            showEarlyAccessBtn.addEventListener('click', () => {
                this.showEarlyAccessModal();
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
        
        // Special handling for provider-specific field (API key)
        ['api-key'].forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                // Track UI state changes
                element.addEventListener('input', () => {
                    this.uiStateManager.updateUIFormValue(id, element.value);
                });
                
                element.addEventListener('blur', () => {
                    // Update UI state
                    this.uiStateManager.updateUIFormValue(id, element.value);
                    
                    // Only save if we're not currently loading provider settings
                    if (!this.isLoadingProviderSettings) {
                        this.saveProviderSettingsContextAware();
                    }
                });
            }
        });
        
        // Hook UIStateManager into all other form elements for state tracking
        this.setupUIStateTracking();
    }
    
    /**
     * Setup UI state tracking for all form elements
     */
    setupUIStateTracking() {
        // Track all form inputs for state management
        const formElements = document.querySelectorAll('input, select, textarea');
        
        formElements.forEach(element => {
            if (element.id && element.id !== 'api-key') {
                // Skip provider-specific fields as they're handled separately
                
                element.addEventListener('input', () => {
                    this.uiStateManager.updateUIFormValue(element.id, this.getElementValue(element));
                });
                
                element.addEventListener('change', () => {
                    this.uiStateManager.updateUIFormValue(element.id, this.getElementValue(element));
                });
            }
        });
        
        // Setup periodic state saving
        this.setupPeriodicStateSaving();
        
        // Setup window unload handling to save state
        this.setupWindowUnloadHandling();
    }
    
    /**
     * Get form element value based on type
     */
    getElementValue(element) {
        if (element.type === 'checkbox') {
            return element.checked;
        } else if (element.type === 'radio') {
            return element.checked ? element.value : null;
        } else {
            return element.value;
        }
    }
    
    /**
     * Setup periodic saving of UI state to prevent data loss
     */
    setupPeriodicStateSaving() {
        // Save UI state every 30 seconds if there are unsaved changes
        setInterval(() => {
            if (this.uiStateManager.hasUnsavedChanges()) {
                this.uiStateManager.saveUIStateToStorage();
                if (window.debugLog) {
                    window.debugLog('Periodic UI state save completed');
                }
            }
        }, 30000); // 30 seconds
    }
    
    /**
     * Setup window unload handling to save state when user navigates away
     */
    setupWindowUnloadHandling() {
        // Save UI state when the window is about to unload
        window.addEventListener('beforeunload', () => {
            // Use synchronous version for unload events
            if (this.uiStateManager.hasUnsavedChanges()) {
                this.uiStateManager.saveUIStateToStorage();
                if (window.debugLog) {
                    window.debugLog('UI state saved during window unload');
                }
            }
        });
        
        // Also handle visibility change (when tab becomes hidden)
        document.addEventListener('visibilitychange', () => {
            if (document.hidden && this.uiStateManager.hasUnsavedChanges()) {
                this.uiStateManager.saveUIStateToStorage();
                if (window.debugLog) {
                    window.debugLog('UI state saved during visibility change');
                }
            }
        });
    }
    
    /**
     * Save current UI state to UIStateManager and trigger persistence
     */
    async saveCurrentUIState() {
        // Save current form state to UIStateManager
        await this.uiStateManager.saveUIStateToStorage();
        
        if (window.debugLog) {
            window.debugLog('UI state saved to storage');
        }
    }
    
    /**
     * Check if there are unsaved changes in the UI
     */
    hasUnsavedChanges() {
        return this.uiStateManager.hasUnsavedChanges();
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
            console.error('Failed to save writing sample:', error);
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

        const confirmed = await this.showConfirmDialog('Delete Sample', `Are you sure you want to delete the sample "${sample.title}"?`);
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
            console.error('Failed to delete writing sample:', error);
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
                        <button type="button" class="sample-btn edit-btn" data-sample-id="${sample.id}" data-action="edit" title="Edit sample">
                            ✏️
                        </button>
                        <button type="button" class="sample-btn delete-btn" data-sample-id="${sample.id}" data-action="delete" title="Delete sample">
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

                
                // Provider is now fixed to default from ai-providers.json
                // await this.applyDomainBasedProviderFiltering();
            }
            
            await this.displayEmailSummary(this.currentEmail);
        } catch (error) {
            console.error('Failed to load current email:', error);
            this.uiController.showError('Failed to load email. Please select an email and try again.');
        }
    }

    async checkForInitialSetupNeeded(showSettingsIfNeeded = true) {
        try {
            const currentSettings = await this.settingsManager.getSettings();
            
            // Always use the default provider from ai-providers.json
            const defaultProvider = this.getDefaultProvider();
            const providerConfig = this.settingsManager.getProviderConfig(defaultProvider);
            const apiKey = providerConfig['api-key'] || '';
            
            // Also check if this appears to be a first-time user (no last-updated timestamp)
            const isFirstTime = !currentSettings['last-updated'];
            
            // If no API key is set for the default provider, or if it's a first-time user
            if (!apiKey.trim() || isFirstTime) {
                if (showSettingsIfNeeded) {
                    console.info('Initial setup needed - showing setup modal');
                    
                    // Show modal notification for initial setup
                    const modalTitle = isFirstTime ? 'Welcome to Prompt Email!' : 'Setup Required';
                    const modalMessage = isFirstTime ? 
                        `🎉 Welcome to Prompt Email! To get started, you'll need to configure your AI provider settings.<br><br>📝 Choose your preferred AI service (OpenAI, Ollama, or on-premises)<br>🔑 Enter your API key<br>⚙️ Customize your preferences` :
                        `🔑 An API key is required to use Prompt Email.<br><br>Please configure your API key in the settings to start analyzing emails and generating responses.`;
                    
                    // Custom buttons for setup modal
                    const setupButtons = [
                        {
                            id: 'modal-setup-now',
                            text: 'Open Settings',
                            class: 'btn-primary',
                            action: () => {
                                this.hideModalNotification();
                                // Open settings panel using the correct method
                                this.openSettings();
                                // Highlight API key field after a brief delay
                                setTimeout(() => this.highlightApiKeyField(), 500);
                            }
                        },
                        {
                            id: 'modal-setup-later',
                            text: 'Maybe Later',
                            class: 'btn-secondary',
                            action: () => this.hideModalNotification()
                        }
                    ];
                    
                    // Show the modal using the same pattern as truncation notifications
                    this.showModalNotification(modalTitle, modalMessage, 'info', setupButtons);
                    
                    // Also show status message
                    if (isFirstTime) {
                        this.uiController.showStatus('Welcome! Please configure your AI provider settings to get started.');
                    } else {
                        this.uiController.showStatus('API key required. Please configure your API key in settings.');
                    }
                    
                    // Switch to settings tab
                    this.openSettings();
                    
                    // Highlight the API key field if it exists
                    setTimeout(() => this.highlightApiKeyField(), 500);
                    
                    // Log this event for analytics
                    this.logger.logEvent('initial_setup_prompted', {
                        selected_service: defaultProvider,
                        has_api_key: !!apiKey.trim(),
                        is_first_time: isFirstTime
                    }, 'Information', this.getUserEmailForTelemetry());
                }
                
                return true; // Indicates setup is needed
            }
            
            return false; // No setup needed
        } catch (error) {
            console.error('Error checking for initial setup:', error);
            return false; // Continue normally on error
        }
    }

    /**
     * Highlights the API key field with visual emphasis and helpful tooltip
     */
    highlightApiKeyField() {
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
    }

    async displayEmailSummary(email) {

        
        // Email overview section has been removed for cleaner UI
        
        // Detect classification for console logging only (no storage)
        if (email.body) {
            const classificationResult = this.classificationDetector.detectClassification(email.body);

        }

        // Context-aware UI adaptation (works behind the scenes)
        this.adaptUIForContext(email.context);
    }

    /**
     * Adapts the UI based on email context (sent vs inbox vs compose)
     * @param {Object} context - Context information from EmailAnalyzer
     */
    adaptUIForContext(context) {

        
        if (!context) {
            console.warn('No context provided for UI adaptation');
            return;
        }


        
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

                this.adaptUIForComposeMode();
            } else if (context.isSentMail) {

                this.adaptUIForSentMail();
            } else {

                this.adaptUIForInboxMail();
            }

        } catch (error) {
            console.error('Error adapting UI for context:', error);
        }
    }

    /**
     * Adapts UI for compose mode (writing new email)
     */
    adaptUIForComposeMode() {

        
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

        
        // Check if auto-analysis is enabled in settings
        const settings = this.settingsManager.getSettings();
        const autoAnalysisEnabled = settings['auto-analysis'] || false;
        
        if (!autoAnalysisEnabled) {

            return;
        }
        
        // Only auto-analyze if we have an email
        if (!this.currentEmail) {
            if (window.debugLog) window.debugLog('No email available for auto-analysis');
            return;
        }

        // Skip auto-analysis if user is in initial setup mode (no API key configured)
        const needsSetup = await this.checkForInitialSetupNeeded(false);
        if (needsSetup) {
            if (window.debugLog) window.debugLog('Initial setup needed, skipping auto-analysis');
            return;
        }

        try {
            // Get current AI provider settings
            const currentSettings = await this.settingsManager.getSettings();
            if (window.debugLog) window.debugLog('Auto-analysis settings check:', {
                fullSettings: currentSettings,
                modelService: currentSettings['model-service'],
                modelServiceType: typeof currentSettings['model-service'],
                modelServiceLength: currentSettings['model-service']?.length
            });
            
            const selectedService = currentSettings['model-service'];
            
            // Also check what the UI element shows
            const modelServiceElement = document.getElementById('model-service');
            if (window.debugLog) window.debugLog('UI element check:', {
                elementExists: !!modelServiceElement,
                elementValue: modelServiceElement?.value,
                elementType: typeof modelServiceElement?.value,
                optionsCount: modelServiceElement?.options?.length,
                selectedIndex: modelServiceElement?.selectedIndex,
                allOptions: modelServiceElement ? Array.from(modelServiceElement.options).map(opt => ({value: opt.value, text: opt.text, selected: opt.selected})) : 'N/A'
            });

            // Check for UI/settings mismatch and log it
            if (modelServiceElement && modelServiceElement.value !== selectedService) {
                console.warn(`UI/Settings mismatch detected: UI shows '${modelServiceElement.value}', settings show '${selectedService}'. This may cause auto-analysis to fail.`);
                window.debugLog(`Syncing UI dropdown to match saved settings: ${selectedService}`);
                modelServiceElement.value = selectedService;
                // Trigger change event to update related UI
                modelServiceElement.dispatchEvent(new Event('change'));
            }
            
            if (!selectedService) {
                console.warn('No AI service configured, skipping auto-analysis');
                return;
            }

            // Check for classification and blocking
            const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
            if (window.debugLog) window.debugLog('Email classification for auto-analysis:', classification);
            
            // Check if auto-analysis should be blocked due to classification
            const currentProvider = selectedService;
            const blockingCheck = this.checkClassificationBlocking(classification, currentProvider);
            if (blockingCheck.blocked) {
                console.warn('Auto-analysis blocked due to classification:', blockingCheck.reason);
                this.uiController.showWarning(`Auto-Analysis Blocked: ${blockingCheck.reason}`);
                return;
            }

            // Test AI service health
            const config = this.getAIConfiguration();
            const isHealthy = await this.aiService.testConnection(config);
            
            if (!isHealthy) {
                if (window.debugLog) window.debugLog('AI service not healthy, skipping auto-analysis');
                return;
            }

            console.info('Connection to AI service is healthy, performing automatic analysis...');
            await this.performAnalysisWithResponse();
            
        } catch (error) {
            console.error('Error during auto-analysis check:', error);
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
                    model_service: this.getProviderLabel(config.service),
                    model_name: config.model,
                    email_length: this.currentEmail.bodyLength,
                    auto_response_generated: false,
                    analysis_duration_ms: analysisEndTime - analysisStartTime
                }, 'Information', this.getUserEmailForTelemetry());
                
                this.uiController.showStatus('Email analyzed automatically. Click "Start Chat Assistant" to generate a response and begin refining.');
                return;
            }
            
            // Auto-generate response as well (if enabled in settings)
            console.info('Auto-generating response after analysis...');
            responseStartTime = Date.now();
            const responseConfig = this.getResponseConfiguration();
            
            // Check email context to determine response type
            const emailContext = this.currentEmail.context || { isSentMail: false };
            
            if (emailContext.isSentMail) {
                // Generate follow-up suggestions for sent mail
                console.info('Generating follow-up suggestions for sent mail...');
                this.currentResponse = await this.aiService.generateFollowupSuggestions(
                    this.currentEmail, 
                    this.currentAnalysis, 
                    { ...config, ...responseConfig }
                );
            } else {
                // Generate response for received mail
                console.info('Generating response for received mail...');
                this.currentResponse = await this.aiService.generateResponse(
                    this.currentEmail, 
                    this.currentAnalysis, 
                    { ...config, ...responseConfig }
                );
            }
            responseEndTime = Date.now();
            
            // Check for email truncation and notify user if it occurred
            const truncationInfo = this.aiService.getLastTruncationInfo();
            if (truncationInfo) {
                this.showEmailTruncationNotification(truncationInfo);
                this.aiService.clearTruncationInfo(); // Clear after showing notification
            }
            
            // Check for HTML conversion and notify user if it occurred
            const htmlConversionInfo = this.aiService.getLastHtmlConversionInfo();
            if (htmlConversionInfo) {
                this.showHtmlConversionNotification(htmlConversionInfo);
                this.aiService.clearHtmlConversionInfo(); // Clear after showing notification
            }
            
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
                model_service: this.getProviderLabel(config.service),
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
            console.error('Auto-analysis failed:', error);
            this.uiController.showStatus('Automatic analysis failed. You can still analyze manually.');
        }
    }

    /**
     * Check if AI analysis is blocked for the current provider due to classification keywords
     * @param {Object} classification - Classification detection result
     * @param {string} currentProvider - Current AI provider key
     * @returns {Object} Object with blocked status and reason
     */
    checkClassificationBlocking(classification, currentProvider) {
        // If no classification detected, allow analysis
        if (!classification.detected || !classification.text) {
            return { blocked: false, reason: null };
        }

        // Get provider config
        const providerConfig = this.defaultProvidersConfig?.[currentProvider];
        if (!providerConfig || !providerConfig.blockedClassifications) {
            return { blocked: false, reason: null };
        }

        // Check if classification matches any blocked keywords (case-insensitive)
        const classificationText = classification.text.toLowerCase();
        const blockedKeywords = providerConfig.blockedClassifications;
        
        for (const keyword of blockedKeywords) {
            if (classificationText.includes(keyword.toLowerCase())) {
                return {
                    blocked: true,
                    reason: `Email contains '${classification.text}' classification which matches blocked keyword '${keyword.toUpperCase()}' for provider '${providerConfig.label || currentProvider}'`
                };
            }
        }

        return { blocked: false, reason: null };
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

        // Check for classification
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        if (window.debugLog) window.debugLog('Email classification check:', classification);
        
        // Get current provider for blocking check and subsequent analysis
        const config = this.getAIConfiguration();
        const currentProvider = config.service;
        
        // Check if analysis should be blocked due to classification
        const blockingCheck = this.checkClassificationBlocking(classification, currentProvider);
        if (blockingCheck.blocked) {
            console.warn('AI analysis blocked due to classification:', blockingCheck.reason);
            this.uiController.showError(`AI Analysis Blocked: ${blockingCheck.reason}`);
            return;
        }
        
        // Proceed with analysis using cached config
        await this.performAnalysis(config);
    }

    async performAnalysis(config = null) {
        const analysisStartTime = Date.now();
        
        try {
            this.uiController.showStatus('Analyzing email...');
            this.uiController.setButtonLoading('analyze-email', true);
            
            // Use provided config or get AI configuration if not provided
            if (!config) {
                config = this.getAIConfiguration();
            }
            
            // Perform analysis
            this.currentAnalysis = await this.aiService.analyzeEmail(this.currentEmail, config);
            const analysisEndTime = Date.now();
            
            // Display results
            this.displayAnalysis(this.currentAnalysis);
            
            // Log successful analysis with flattened performance telemetry
            this.logger.logEvent('email_analyzed', {
                model_service: this.getProviderLabel(config.service),
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
            console.error('Analysis failed:', error);
            
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
                userMessage = 'Analysis failed: Service endpoint not found. Please check your provider configuration.';
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

        // Check for classification and blocking
        const classification = this.classificationDetector.detectClassification(this.currentEmail.body);
        if (window.debugLog) window.debugLog('Email classification detected for response generation:', classification);

        // Get current provider for blocking check
        const config = this.getAIConfiguration();
        const currentProvider = config.service;
        
        // Check if response generation should be blocked due to classification
        const blockingCheck = this.checkClassificationBlocking(classification, currentProvider);
        if (blockingCheck.blocked) {
            console.warn('Response generation blocked due to classification:', blockingCheck.reason);
            this.uiController.showError(`Response Generation Blocked: ${blockingCheck.reason}`);
            return;
        }

        try {
            this.uiController.showStatus('Starting chat assistant...');
            this.uiController.setButtonLoading('generate-response', true);
            
            // Get configuration
            const config = this.getAIConfiguration();
            const responseConfig = this.getResponseConfiguration();
            
            // Ensure we have analysis data - if not, run analysis first
            let analysisData = this.currentAnalysis;
            if (!analysisData) {
                console.warn('No current analysis available, running analysis first');
                this.uiController.showStatus('Analyzing email before generating response...');
                
                try {
                    // Run analysis first
                    await this.performAnalysis();
                    analysisData = this.currentAnalysis;
                    
                    if (!analysisData) {
                        // If analysis still failed, create minimal default
                        console.warn('Analysis failed, using default analysis');
                        analysisData = {
                            keyPoints: ['Email content needs response'],
                            sentiment: 'neutral',
                            responseStrategy: 'respond professionally and appropriately'
                        };
                    }
                } catch (analysisError) {
                    console.warn('Analysis failed, using default analysis:', analysisError);
                    analysisData = {
                        keyPoints: ['Email content needs response'],
                        sentiment: 'neutral',
                        responseStrategy: 'respond professionally and appropriately'
                    };
                }
                
                this.uiController.showStatus('Generating response...');
            }
            
            // Start timing for telemetry
            const responseStartTime = Date.now();
            
            // Generate response
            this.currentResponse = await this.aiService.generateResponse(
                this.currentEmail, 
                analysisData,
                { ...config, ...responseConfig }
            );
            
            // End timing for telemetry
            const responseEndTime = Date.now();
            
            console.info('Response generated:', this.currentResponse);
            
            // Log successful manual response generation
            this.logger.logEvent('response_generated', {
                model_service: this.getProviderLabel(config.service),
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                response_length: this.currentResponse.text ? this.currentResponse.text.length : 0,
                generation_type: 'manual_response',
                had_prior_analysis: !!this.currentAnalysis,
                email_context: this.currentEmail.context ? (this.currentEmail.context.isSentMail ? 'sent' : 'inbox') : 'unknown',
                refinement_count: this.refinementCount,
                clipboard_used: this.hasUsedClipboard,
                response_generation_duration_ms: responseEndTime - responseStartTime
            }, 'Information', this.getUserEmailForTelemetry());
            
            // Check for email truncation and notify user if it occurred
            const truncationInfo = this.aiService.getLastTruncationInfo();
            if (truncationInfo) {
                this.showEmailTruncationNotification(truncationInfo);
                this.aiService.clearTruncationInfo(); // Clear after showing notification
            }
            
            // Check for HTML conversion and notify user if it occurred
            const htmlConversionInfo = this.aiService.getLastHtmlConversionInfo();
            if (htmlConversionInfo) {
                this.showHtmlConversionNotification(htmlConversionInfo);
                this.aiService.clearHtmlConversionInfo(); // Clear after showing notification
            }
            
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
            console.error('Response generation failed:', error);
            
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
                userMessage = 'Response generation failed: Service endpoint not found. Please check your provider configuration.';
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

        // Get configuration outside try block so it's available in catch
        const config = this.getAIConfiguration();
        const responseConfig = this.getResponseConfiguration();

        try {
            this.uiController.showStatus('Generating follow-up suggestions...');
            this.uiController.setButtonLoading('generate-response', true);
            
            // Ensure we have analysis data - if not, run analysis first
            let analysisData = this.currentAnalysis;
            if (!analysisData) {
                console.warn('No current analysis available, running analysis first');
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
                    console.warn('Analysis failed, using default analysis:', analysisError);
                    analysisData = {
                        keyPoints: ['Sent email content analyzed'],
                        sentiment: 'neutral', 
                        responseStrategy: 'generate appropriate follow-up suggestions'
                    };
                }
                
                this.uiController.showStatus('Generating follow-up suggestions...');
            }
            
            // Start timing for telemetry
            const followupStartTime = Date.now();
            
            // Generate follow-up suggestions instead of response
            this.currentResponse = await this.aiService.generateFollowupSuggestions(
                this.currentEmail, 
                analysisData,
                { ...config, ...responseConfig }
            );
            
            // End timing for telemetry
            const followupEndTime = Date.now();
            
            console.info('Follow-up suggestions generated:', this.currentResponse);
            
            // Check for email truncation and notify user if it occurred
            const truncationInfo = this.aiService.getLastTruncationInfo();
            if (truncationInfo) {
                this.showEmailTruncationNotification(truncationInfo);
                this.aiService.clearTruncationInfo(); // Clear after showing notification
            }
            
            // Check for HTML conversion and notify user if it occurred
            const htmlConversionInfo = this.aiService.getLastHtmlConversionInfo();
            if (htmlConversionInfo) {
                this.showHtmlConversionNotification(htmlConversionInfo);
                this.aiService.clearHtmlConversionInfo(); // Clear after showing notification
            }
            
            // Log telemetry for follow-up suggestions generation
            this.logger.logEvent('followup_suggestions_generated', {
                model_service: this.getProviderLabel(config.service),
                model_name: config.model,
                email_length: this.currentEmail.bodyLength,
                recipients_count: this.currentEmail.recipients.split(',').length,
                suggestions_length: this.currentResponse.suggestions ? this.currentResponse.suggestions.length : 0,
                analysis_available: !!analysisData,
                generation_success: true,
                refinement_count: this.refinementCount,
                followup_generation_duration_ms: followupEndTime - followupStartTime
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
            console.error('Follow-up suggestion generation failed:', error);
            
            // Log telemetry for failed follow-up suggestions
            this.logger.logEvent('followup_suggestions_failed', {
                error_message: error.message,
                model_service: config ? this.getProviderLabel(config.service) : 'unknown',
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
                userMessage = 'Follow-up generation failed: Service endpoint not found. Please check your provider configuration.';
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
            
            // Store previous response for history tracking - with safety check
            if (!this.currentResponse || (!this.currentResponse.text && !this.currentResponse.suggestions)) {
                console.error('Current response is invalid:', this.currentResponse);
                this.uiController.showError('No valid response available. Please generate a response first.');
                return;
            }
            
            // Get the response text based on response type
            const previousResponse = this.currentResponse.text || this.currentResponse.suggestions;
            
            // Start timing for telemetry
            const chatStartTime = Date.now();
            
            // Use history-aware refinement with chat context
            this.currentResponse = await this.aiService.refineResponseWithHistory(
                this.currentResponse,
                message,
                config,
                responseConfig,
                this.originalEmailContext,
                this.conversationHistory
            );
            
            // End timing for telemetry
            const chatEndTime = Date.now();
            
            // Validate the refined response
            if (!this.currentResponse || !this.currentResponse.text) {
                console.error('Response refinement returned invalid response:', this.currentResponse);
                this.uiController.showError('Failed to refine response. Please try again.');
                return;
            }
            

            
            // Check for email truncation and notify user if it occurred
            const truncationInfo = this.aiService.getLastTruncationInfo();
            if (truncationInfo) {
                this.showEmailTruncationNotification(truncationInfo);
                this.aiService.clearTruncationInfo(); // Clear after showing notification
            }
            
            // Check for HTML conversion and notify user if it occurred
            const htmlConversionInfo = this.aiService.getLastHtmlConversionInfo();
            if (htmlConversionInfo) {
                this.showHtmlConversionNotification(htmlConversionInfo);
                this.aiService.clearHtmlConversionInfo(); // Clear after showing notification
            }
            
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
            
            // Log chat interaction with detailed metrics
            this.logger.logEvent('chat_message_sent', {
                model_service: this.getProviderLabel(config.service),
                model_name: config.model,
                refinement_count: this.refinementCount,
                message_length: message.length,
                conversation_length: this.conversationHistory.length,
                chat_duration_ms: chatEndTime - chatStartTime,
                response_length: this.currentResponse.text ? this.currentResponse.text.length : 0,
                previous_response_length: previousResponse ? previousResponse.length : 0,
                email_length: this.currentEmail ? this.currentEmail.bodyLength : 0,
                conversation_steps: this.conversationHistory.length
            }, 'Information', this.getUserEmailForTelemetry());
            
        } catch (error) {
            console.error('Chat message failed:', error);
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

    async clearChatHistory() {
        const chatMessages = document.getElementById('chat-messages');
        if (chatMessages) {
            chatMessages.innerHTML = '';
        }

        // Clear conversation history but keep original email context
        this.conversationHistory = [];
        
        // Add a system message
        this.addChatMessage('system', 'Generating fresh response based on analysis...');
        
        // Generate a fresh initial draft based on the analysis
        try {
            if (this.currentAnalysis && this.originalEmailContext) {
                // Get current configuration
                const config = this.getAIConfiguration();
                const responseConfig = this.getResponseConfiguration();
                
                // Show loading state
                this.uiController.setButtonLoading('clear-chat', true);
                
                // Generate fresh response based on current analysis
                if (this.originalEmailContext.isSentMail) {
                    // Generate fresh follow-up suggestions for sent mail
                    this.currentResponse = await this.aiService.generateFollowupSuggestions(
                        this.currentEmail, 
                        this.currentAnalysis, 
                        { ...config, ...responseConfig }
                    );
                } else {
                    // Generate fresh response for received mail
                    this.currentResponse = await this.aiService.generateResponse(
                        this.currentEmail, 
                        this.currentAnalysis, 
                        { ...config, ...responseConfig }
                    );
                }
                
                // Display the fresh response in main area AND chat
                this.displayResponse(this.currentResponse);
                
                // Clear the system message and add the fresh response to chat
                if (chatMessages) {
                    chatMessages.innerHTML = '';
                }
                
                // Add the fresh response to chat
                const responseContent = this.currentResponse.text || this.currentResponse.suggestions;
                this.addChatMessage('assistant', responseContent);
                this.addChatMessage('system', 'Fresh response generated based on original analysis. You can now refine it further.');
                
            } else {
                this.addChatMessage('system', 'Chat history cleared. No analysis available to generate fresh response.');
            }
        } catch (error) {
            console.error('Failed to generate fresh response:', error);
            if (chatMessages) {
                chatMessages.innerHTML = '';
            }
            this.addChatMessage('system', 'Chat history cleared. Failed to generate fresh response - please try generating a new response.');
        } finally {
            this.uiController.setButtonLoading('clear-chat', false);
        }
        
        // Log chat reset event
        this.logger.logEvent('chat_reset_with_fresh_response', {
            analysis_available: !!this.currentAnalysis,
            email_context_available: !!this.originalEmailContext
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
            console.error('Failed to copy latest response:', error);
            this.uiController.showError('Failed to copy response to clipboard. Please try again.');
        }
    }

    getAIConfiguration() {
        // Use current UI selection if available, otherwise default from ai-providers.json
        let service = '';
        if (this.modelServiceSelect && this.modelServiceSelect.value) {
            service = this.modelServiceSelect.value;
        } else {
            service = this.getDefaultProvider();
        }
        
        // Always use default model from ai-providers.json (no persistence of model selections)
        let model = '';
        if (service) {
            model = this.defaultProvidersConfig?.[service]?.defaultModel || '';
        }
        
        // Fall back to UI dropdown only if no provider config exists
        if (!model && this.modelSelect && this.modelSelect.value) {
            model = this.modelSelect.value;
        }
        
        // Get provider-specific configuration
        let apiKey = '';
        let endpointUrl = '';
        
        // Check if settings panel is open for context-aware behavior
        const settingsPanel = document.getElementById('settings-panel');
        const isSettingsOpen = settingsPanel && !settingsPanel.classList.contains('hidden');
        
        if (service) {
            const providerConfig = this.settingsManager.getProviderConfig(service);
            apiKey = providerConfig['api-key'] || '';
            
            // Use default endpoint from ai-providers.json
            endpointUrl = this.defaultProvidersConfig?.[service]?.baseUrl || '';
            
            // In settings context, use UI API key if it has unsaved changes
            if (isSettingsOpen) {
                const apiKeyElement = document.getElementById('api-key');
                
                // Use UI API key if populated
                if (apiKeyElement && apiKeyElement.value.trim()) {
                    apiKey = apiKeyElement.value.trim();
                }
            }
        }
        
        const config = {
            service,
            apiKey,
            endpointUrl,
            model,
            settingsManager: this.settingsManager
        };
        
        // Debug logging
        if (window.debugLog) {
            window.debugLog(`getAIConfiguration result:`, {
                service,
                'apiKey': apiKey ? '[HIDDEN]' : '[EMPTY]',
                endpointUrl: endpointUrl || '[EMPTY]',
                model: model || '[EMPTY]'
            });
        }
        
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
            window.debugLog('getSettingsAIConfiguration: No provider selected in settings');
            return { service: '', apiKey: '', endpointUrl: '', model: '' };
        }
        
        // Get provider-specific configuration from saved settings
        const providerConfig = this.settingsManager.getProviderConfig(service);
        let apiKey = providerConfig['api-key'] || '';
        let endpointUrl = providerConfig['endpoint-url'] || '';
        
        // For settings context, use current UI values for API key (for immediate testing before saving)
        // But for endpoint URL, always use defaults from ai-providers.json since field is read-only
        const apiKeyElement = document.getElementById('api-key');
        if (apiKeyElement && apiKeyElement.value) {
            apiKey = apiKeyElement.value;
        }
        
        // Always use default endpoint URL for this provider, ignoring any stored overrides
        const defaultEndpoint = this.defaultProvidersConfig?.[service]?.baseUrl || '';
        if (defaultEndpoint) {
            endpointUrl = defaultEndpoint;
        }
        
        // For settings operations, we don't need a specific model - use default
        const model = this.getDefaultModelForProvider(service);
        
        const config = {
            service,
            apiKey,
            endpointUrl,
            model
        };
        
        window.debugLog(`getSettingsAIConfiguration returning:`, { 
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

    /**
     * Gets the human-readable label for a provider service key
     * @param {string} serviceKey - The provider key (e.g., 'onsite1')
     * @returns {string} - The provider label (e.g., 'onsite1-label') or fallback to service key
     */
    getProviderLabel(serviceKey) {
        if (!serviceKey || !this.defaultProvidersConfig) {
            return serviceKey || 'unknown';
        }
        
        const providerConfig = this.defaultProvidersConfig[serviceKey];
        if (providerConfig && providerConfig.label) {
            return providerConfig.label;
        }
        
        // Fallback to service key if no label found
        return serviceKey;
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

    /**
     * Completely clears model dropdown state to prevent contamination
     */
    clearModelDropdownState() {
        if (this.modelSelect) {
            // Clear all options immediately
            this.modelSelect.innerHTML = '';
            this.modelSelect.value = '';
            
            // Remove any data attributes that might hold stale info
            this.modelSelect.removeAttribute('data-provider');
            this.modelSelect.removeAttribute('data-last-update');
        }
        
        // Clear any error displays
        const errorDiv = document.getElementById('model-fetch-error');
        if (errorDiv) {
            errorDiv.style.display = 'none';
            errorDiv.innerHTML = '';
        }
    }

    async updateModelDropdown() {
        if (!this.modelServiceSelect || !this.modelSelectGroup || !this.modelSelect) return;
        
        // Skip if settings haven't been loaded yet (prevents API key issues during startup)
        const settings = this.settingsManager.getSettings();
        if (!settings || !settings['settings-version']) {
            return;
        }
        
        // Also skip if provider configuration isn't available yet (prevents timing issues during provider switching)
        const currentProvider = this.modelServiceSelect.value;
        const providerConfig = this.settingsManager.getProviderConfig(currentProvider);
        if (!providerConfig || (!providerConfig['api-key'] && currentProvider !== 'ollama')) {
            return;
        }

        // Generate a unique operation ID to prevent race conditions
        const operationId = Date.now() + Math.random();
        this._currentModelDropdownOperation = operationId;
        
        // Immediately clear all dropdown state to prevent contamination
        this.clearModelDropdownState();
        
        const aiConfigPlaceholder = document.getElementById('ai-config-placeholder');
        this.modelSelectGroup.style.display = 'none';
        this.modelSelect.innerHTML = '<option value="">Loading...</option>';
        let models = [];
        let preferred = '';
        let errorMsg = '';
        

        
        // Store the current provider at the start to verify we're still on the same provider when async operations complete
        const providerAtStart = this.modelServiceSelect.value;
        
        if (this.modelServiceSelect.value === 'ollama') {
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            const endpointUrlElement = document.getElementById('endpoint-url');
            let baseUrl = (endpointUrlElement && endpointUrlElement.value) || 'http://localhost:11434';
            
            // Ensure we're using the correct Ollama endpoint
            const defaultOllamaUrl = this.defaultProvidersConfig?.ollama?.baseUrl || 'http://localhost:11434';
            if (baseUrl !== defaultOllamaUrl) {
                console.warn(`Ollama endpoint URL mismatch. Expected: ${defaultOllamaUrl}, Found: ${baseUrl}. Using correct URL.`);
                baseUrl = defaultOllamaUrl;
                // Update the UI to show the correct endpoint
                if (endpointUrlElement) {
                    endpointUrlElement.value = baseUrl;
                }
            }
            
            try {
                models = await AIService.fetchOllamaModels(baseUrl);
                
                // Check if this operation is still valid (prevents race conditions during rapid provider switching)
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                // Verify we're still on the same provider after async operation
                if (this.modelServiceSelect.value !== providerAtStart) {
                    return;
                }
                
                // Final check before populating dropdown
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                // Filter out models that start with "TEXT" or "IMAGE"
                models = models.filter(model => !model.toUpperCase().startsWith('TEXT') && !model.toUpperCase().startsWith('IMAGE'));
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
                
                // Always use default model from ai-providers.json
                preferred = this.defaultProvidersConfig?.ollama?.defaultModel || this.getFallbackModel();
                
                if (preferred && models.includes(preferred)) {
                    this.modelSelect.value = preferred;
                } else if (models.length) {
                    this.modelSelect.value = models[0];
                }
                
                // No longer saving model selections - always use ai-providers.json defaults
            } catch (err) {
                // Check if this operation is still valid (prevents race conditions)
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                // Verify we're still on the same provider after async operation
                if (this.modelServiceSelect.value !== providerAtStart) {
                    return;
                }
                
                errorMsg = `Error fetching models: ${err.message || err}`;
                
                // If we can't fetch models, show the default from ai-providers.json
                const defaultModel = this.defaultProvidersConfig?.ollama?.defaultModel || 'llama3:latest';
                
                this.modelSelect.innerHTML = `<option value="${defaultModel}">${defaultModel}</option>`;
                this.modelSelect.value = defaultModel;
            }
        } else if (this.defaultProvidersConfig?.[this.modelServiceSelect.value]?.apiFormat === 'bedrock') {
            // Handle AWS Bedrock services
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            
            const serviceKey = this.modelServiceSelect.value;
            const providerConfig = this.settingsManager.getProviderConfig(serviceKey);
            
            try {
                // Use the Bedrock-specific model fetching
                const aiService = new AIService(this.defaultProvidersConfig);
                const bedrockConfig = {
                    service: serviceKey,
                    apiKey: providerConfig['api-key'] || '',
                    region: this.defaultProvidersConfig[serviceKey].region || 'us-east-1'
                };
                
                models = await aiService.fetchBedrockModels(bedrockConfig);
                
                // Check if this operation is still valid (prevents race conditions)
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                // Verify we're still on the same provider after async operation
                if (this.modelServiceSelect.value !== providerAtStart) {
                    return;
                }
                
                // Final check before populating dropdown
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m.id}">${m.name} (${m.provider})</option>`).join('')
                    : '<option value="">No models found</option>';
                    
                // Always use default model from ai-providers.json
                preferred = this.defaultProvidersConfig?.[serviceKey]?.defaultModel || models[0]?.id || '';
                    
                if (preferred && models.find(m => m.id === preferred)) {
                    this.modelSelect.value = preferred;
                } else if (models.length) {
                    this.modelSelect.value = models[0].id;
                }
                
                // No longer saving model selections - always use ai-providers.json defaults
                
            } catch (err) {
                // Check if this operation is still valid (prevents race conditions)
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                // Verify we're still on the same provider after async operation
                if (this.modelServiceSelect.value !== providerAtStart) {
                    return;
                }
                
                errorMsg = `Error fetching Bedrock models: ${err.message || err}`;
                this.modelSelect.innerHTML = '<option value="">Error fetching models</option>';
            }
        } else if (this.modelServiceSelect.value !== 'ollama' && 
                   this.defaultProvidersConfig?.[this.modelServiceSelect.value]?.apiFormat !== 'bedrock') {
            // Handle OpenAI-compatible services (openai, onsite1, onsite2, etc.)
            this.modelSelectGroup.style.display = '';
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
            
            const serviceKey = this.modelServiceSelect.value;
            
            // Get endpoint URL from saved configuration (same logic as getAIConfiguration)
            const currentProvider = this.modelServiceSelect.value;
            const providerConfig = this.settingsManager.getProviderConfig(currentProvider);
            let endpoint = providerConfig['endpoint-url'] || '';
            
            // If no stored endpoint URL, use default from ai-providers.json
            if (!endpoint && this.defaultProvidersConfig && this.defaultProvidersConfig[serviceKey] && this.defaultProvidersConfig[serviceKey].baseUrl) {
                endpoint = this.defaultProvidersConfig[serviceKey].baseUrl;
            } else if (!endpoint) {
                // Final fallback
                endpoint = 'http://localhost:11434/v1';
            }
            
            if (endpoint.endsWith('/')) endpoint = endpoint.slice(0, -1);
            
            // Get API key from saved configuration for the current provider
            // (currentProvider and providerConfig already declared above)
            const apiKey = providerConfig['api-key'] || '';
            
            try {
                models = await AIService.fetchOpenAICompatibleModels(endpoint, apiKey);
                
                // Check if this operation is still valid (prevents race conditions)
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                // Verify we're still on the same provider after async operation
                if (this.modelServiceSelect.value !== providerAtStart) {
                    return;
                }
                
                // Final check before populating dropdown
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                // Filter out models that start with "TEXT" or "IMAGE"
                models = models.filter(model => !model.toUpperCase().startsWith('TEXT') && !model.toUpperCase().startsWith('IMAGE'));
                this.modelSelect.innerHTML = models.length
                    ? models.map(m => `<option value="${m}">${m}</option>`).join('')
                    : '<option value="">No models found</option>';
            } catch (err) {
                // Check if this operation is still valid (prevents race conditions)
                if (this._currentModelDropdownOperation !== operationId) {
                    return;
                }
                
                // Verify we're still on the same provider after async operation
                if (this.modelServiceSelect.value !== providerAtStart) {
                    return;
                }
                
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
            
            // Always use default model from ai-providers.json
            preferred = this.defaultProvidersConfig?.[serviceKey]?.defaultModel || this.getFallbackModel();
                
            if (preferred && models.includes(preferred)) {
                this.modelSelect.value = preferred;
            } else if (models.length) {
                this.modelSelect.value = models[0];
            }
            
            // No longer saving model selections - always use ai-providers.json defaults
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
        // Final check - only apply UI changes if this operation is still valid
        if (this._currentModelDropdownOperation !== operationId) {
            return;
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
        if (window.debugLog) window.debugLog('Displaying response:', response);
        const container = document.getElementById('response-draft');
        
        if (!container) {
            console.error('response-draft container not found');
            return;
        }
        
        if (!response || (!response.text && !response.suggestions)) {
            console.error('Invalid response object:', response);
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
        
        console.info('Response displayed successfully');
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
            console.error('Failed to copy response:', error);
            this.uiController.showError('Failed to copy response to clipboard.');
        }
    }

    /**
     * Copies response with HTML table support to clipboard
     * @param {string} responseText - The response text containing HTML tables
     */
    async copyResponseWithHtml(responseText) {
        // Create both HTML and plain text optimized for Outlook
        const outlookHtmlContent = this.formatResponseForOutlookHtml(responseText);
        const outlookPlainText = this.formatForOutlookClipboard(responseText);
        
        // Use the ClipboardItem API to provide both formats for Outlook
        if (navigator.clipboard && typeof ClipboardItem !== 'undefined') {
            try {
                const clipboardItem = new ClipboardItem({
                    'text/html': new Blob([outlookHtmlContent], { type: 'text/html' }),
                    'text/plain': new Blob([outlookPlainText], { type: 'text/plain' })
                });
                await navigator.clipboard.write([clipboardItem]);
            } catch (error) {
                console.error('Failed to copy rich content:', error);
                // Fallback to plain text
                await navigator.clipboard.writeText(outlookPlainText);
            }
        } else {
            // Fallback to plain text only
            await navigator.clipboard.writeText(outlookPlainText);
        }
    }

    /**
     * Formats response as HTML specifically optimized for Outlook
     * @param {string} responseText - Raw response text
     * @returns {string} Outlook-optimized HTML content
     */
    formatResponseForOutlookHtml(responseText) {
        let formatted = responseText.trim();
        
        // Aggressively clean up excessive newlines BEFORE converting to HTML
        formatted = formatted.replace(/\n{2,}/g, '\n\n'); // Max 2 newlines anywhere
        
        // Clean up spacing around tables specifically
        formatted = formatted.replace(/\n{2,}(<table[\s\S]*?<\/table>)/gi, '\n\n$1'); // Before tables
        formatted = formatted.replace(/(<\/table>)\n{2,}/gi, '$1\n\n'); // After tables
        
        // Convert remaining line breaks to HTML
        formatted = formatted.replace(/\n\n/g, '</p><p>'); // Double newlines become paragraph breaks
        formatted = formatted.replace(/\n/g, '<br>'); // Single newlines become breaks
        
        // Wrap in paragraphs
        formatted = `<p>${formatted}</p>`;
        
        // Clean up any empty paragraphs that might have been created
        formatted = formatted.replace(/<p>\s*<\/p>/g, '');
        formatted = formatted.replace(/<p>\s*<br>\s*<\/p>/g, '');
        
        // Optimize existing HTML tables for Outlook compatibility
        formatted = formatted.replace(/<table[^>]*>/gi, (match) => {
            return `<table cellpadding="4" cellspacing="0" border="1" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 14px; width: 100%; max-width: 600px; margin: 10px 0;">`;
        });
        
        // Ensure all th elements have proper Outlook-compatible styling
        formatted = formatted.replace(/<th[^>]*>/gi, 
            `<th style="border: 1px solid #000000; background-color: #f0f0f0; padding: 6px 8px; text-align: left; font-weight: bold; font-family: Arial, sans-serif;">`
        );
        
        // Ensure all td elements have proper Outlook-compatible styling
        formatted = formatted.replace(/<td[^>]*>/gi, 
            `<td style="border: 1px solid #000000; padding: 6px 8px; text-align: left; font-family: Arial, sans-serif; vertical-align: top;">`
        );
        
        // Clean up spacing around tables in the final HTML
        formatted = formatted.replace(/<\/p>\s*<table/gi, '</p><table');
        formatted = formatted.replace(/<\/table>\s*<p>/gi, '</table><p>');
        
        // Wrap the entire content in a container that Outlook likes
        const htmlContent = `<div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.4; color: #000000;">${formatted}</div>`;
        
        return htmlContent;
    }

    /**
     * Formats content specifically for Outlook clipboard compatibility
     * @param {string} responseText - Raw response text
     * @returns {string} Outlook-optimized content
     */
    formatForOutlookClipboard(responseText) {
        let formatted = responseText;
        
        // Convert HTML tables to a simpler format that Outlook handles better
        formatted = formatted.replace(/<table[\s\S]*?<\/table>/gi, (tableMatch) => {
            return this.convertTableForOutlook(tableMatch);
        });
        
        // Remove any remaining HTML tags
        formatted = formatted.replace(/<[^>]*>/g, '');
        
        // Outlook-specific formatting
        // Use Windows line endings (CRLF) which Outlook prefers
        formatted = formatted.replace(/\r\n?|\n/g, '\r\n');
        
        // Ensure proper paragraph spacing for Outlook
        formatted = formatted.replace(/\r\n{3,}/g, '\r\n\r\n'); // Max double line breaks
        
        // Clean up any excessive whitespace
        formatted = formatted.replace(/[ \t]+/g, ' '); // Normalize spaces
        formatted = formatted.replace(/^\s+|\s+$/gm, ''); // Trim lines
        
        return formatted.trim();
    }

    /**
     * Convert tables to a format that works better in Outlook
     * @param {string} tableHtml - HTML table string
     * @returns {string} Outlook-friendly table format
     */
    convertTableForOutlook(tableHtml) {
        try {
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = tableHtml;
            const table = tempDiv.querySelector('table');
            
            if (!table) return tableHtml;
            
            let outlookTable = '';
            const rows = table.querySelectorAll('tr');
            
            rows.forEach((row, rowIndex) => {
                const cells = row.querySelectorAll('th, td');
                const cellTexts = Array.from(cells).map(cell => cell.textContent.trim());
                
                if (rowIndex === 0 && row.querySelectorAll('th').length > 0) {
                    // Header - use a simple underlined format
                    outlookTable += cellTexts.join(' | ') + '\r\n';
                    outlookTable += '-'.repeat(Math.min(cellTexts.join(' | ').length, 60)) + '\r\n';
                } else {
                    // Data rows - use simple formatting that Outlook won't mess up
                    if (cellTexts.length <= 2) {
                        outlookTable += `${cellTexts.join(': ')}\r\n`;
                    } else {
                        // For multi-column, use a format that's clear in plain text
                        outlookTable += `${rowIndex}. ${cellTexts.join(' → ')}\r\n`;
                    }
                }
            });
            
            return outlookTable.replace(/\r\n$/, ''); // Remove trailing newline
        } catch (error) {
            console.warn('Failed to convert table for Outlook:', error);
            return tableHtml;
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
        if (window.debugLog) {
            window.debugLog('Original response text length:', responseText.length);
            window.debugLog('Original newline count:', (responseText.match(/\n/g) || []).length);
        }
        
        let formatted = responseText;
        
        // Convert HTML tables to plain text format
        formatted = formatted.replace(/<table[\s\S]*?<\/table>/gi, (tableMatch) => {
            return this.convertHtmlTableToPlainText(tableMatch);
        });
        
        // Remove any remaining HTML tags
        formatted = formatted.replace(/<[^>]*>/g, '');
        
        if (window.debugLog) {
            window.debugLog('After HTML removal, newline count:', (formatted.match(/\n/g) || []).length);
        }
        
        // NUCLEAR APPROACH: Completely rebuild the content with controlled spacing
        
        // Step 1: Split into lines and intelligently manage spacing
        let lines = formatted.split(/\r?\n/);
        let cleanedLines = [];
        let consecutiveEmptyLines = 0;
        
        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            
            if (line === '') {
                consecutiveEmptyLines++;
                // Allow maximum 1 empty line (which creates paragraph breaks)
                if (consecutiveEmptyLines <= 1) {
                    cleanedLines.push('');
                }
            } else {
                consecutiveEmptyLines = 0;
                cleanedLines.push(line);
            }
        }
        
        // Step 2: Rejoin with single newlines
        formatted = cleanedLines.join('\n');
        
        // Step 2a: Ensure paragraph breaks are preserved where needed
        // After sentences that should start new paragraphs
        formatted = formatted.replace(/([.!?])\n([A-Z][^.!?]*[.!?])/g, '$1\n\n$2');
        // After email signatures
        formatted = formatted.replace(/(\],?)\n(\*\*[^*]+\*\*)/g, '$1\n\n$2');
        
        // Step 3: Now carefully add paragraph breaks only where needed
        // After greeting
        formatted = formatted.replace(/((?:Hi|Hello|Dear)\s+[^,]+,)\n([A-Z])/gi, '$1\n\n$2');
        // Before signature  
        formatted = formatted.replace(/([.!?])\n((?:Best\s+)?(?:regards?|sincerely|thanks?|cheers),?\s*)/gi, '$1\n\n$2');
        // Before section headers (markdown style)
        formatted = formatted.replace(/([.!?])\n(\*\*[^*]+\*\*)/g, '$1\n\n$2');
        // After sentences ending with period before table headers
        formatted = formatted.replace(/(\w+\.)\n(\w+[\w\s]*\t)/g, '$1\n\n$2');
        
        // Step 4: Ensure there's proper spacing before tables but never more than double
        formatted = formatted.replace(/\n+(\w+[\w\s]*\t[\w\s]*)/g, '\n\n$1');
        
        // Step 5: Final safety - absolutely no triple newlines
        formatted = formatted.replace(/\n{3,}/g, '\n\n');
        
        if (window.debugLog) {
            window.debugLog('After complete rebuild, newline count:', (formatted.match(/\n/g) || []).length);
            window.debugLog('Final text preview:', formatted.substring(0, 300));
        }
        
        return formatted.trim();
    }

    /**
     * Minimal formatting for clipboard - preserves structure without excessive spacing
     * @param {string} text - The text to format
     * @returns {string} Minimally formatted text
     */
    formatTextForClipboardMinimal(text) {
        let formatted = text.trim();
        
        if (window.debugLog) {
            window.debugLog('formatTextForClipboardMinimal input newlines:', (formatted.match(/\n/g) || []).length);
        }
        
        // Normalize line endings
        formatted = formatted.replace(/\r\n?/g, '\n');
        
        // EXTREME cleanup - absolutely no more than 2 consecutive newlines ANYWHERE
        formatted = formatted.replace(/\n{2,}/g, '\n\n');
        
        // Multiple passes to catch any remaining patterns
        for (let i = 0; i < 3; i++) {
            formatted = formatted.replace(/\n{3,}/g, '\n\n'); // Keep reducing
            formatted = formatted.replace(/\n\s+\n/g, '\n\n'); // Remove whitespace between newlines
        }
        
        // Ensure proper spacing after greetings and before signatures only if needed
        if (!formatted.includes('\n\n')) {
            formatted = formatted.replace(/((?:Hi|Hello|Dear)\s+[^,]+,)\s*([A-Z])/gi, '$1\n\n$2');
            formatted = formatted.replace(/([.!?])\s*((?:Best\s+)?(?:regards?|sincerely|thanks?|cheers),?\s*\n?\s*[\w\s]+)$/gi, '$1\n\n$2');
        }
        
        // Final nuclear option - scan the entire text and replace any sequence of 3+ newlines
        formatted = formatted.replace(/\n{3,}/g, '\n\n');
        
        if (window.debugLog) {
            window.debugLog('formatTextForClipboardMinimal output newlines:', (formatted.match(/\n/g) || []).length);
        }
        
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
            window.debugLog('Conversation history initialized for new email');
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
            window.debugLog(`Added refinement step ${conversationStep.step} to conversation history`);
        }
    }

    /**
     * Clears conversation history (called when starting new analysis)
     */
    clearConversationHistory() {
        this.conversationHistory = [];
        this.originalEmailContext = null;
        
        if (window.debugLog) {

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
            
            let textTable = '';
            const rows = table.querySelectorAll('tr');
            
            // Convert to a more email-friendly bullet list format instead of tabs
            rows.forEach((row, rowIndex) => {
                const cells = row.querySelectorAll('th, td');
                const cellTexts = Array.from(cells).map(cell => cell.textContent.trim());
                
                if (rowIndex === 0 && row.querySelectorAll('th').length > 0) {
                    // Header row - make it a simple title
                    textTable += cellTexts.join(' | ') + '\n';
                    textTable += '='.repeat(Math.min(cellTexts.join(' | ').length, 50)) + '\n';
                } else {
                    // Data rows - convert to bullet points for better email compatibility
                    if (cellTexts.length === 1) {
                        textTable += `• ${cellTexts[0]}\n`;
                    } else if (cellTexts.length === 2) {
                        textTable += `• ${cellTexts[0]}: ${cellTexts[1]}\n`;
                    } else if (cellTexts.length === 3) {
                        // Three columns - likely step, direction, time/distance
                        textTable += `${cellTexts[0]}. ${cellTexts[1]} (${cellTexts[2]})\n`;
                    } else {
                        // Multiple columns - create structured list
                        textTable += `${rowIndex}. ${cellTexts.join(' - ')}\n`;
                    }
                }
            });
            
            // Remove trailing newline to prevent extra spacing
            textTable = textTable.replace(/\n$/, '');
            
            if (window.debugLog) {
                window.debugLog('Converted table to email-friendly format:', {
                    originalLength: tableHtml.length,
                    convertedLength: textTable.length,
                    preview: textTable.substring(0, 200)
                });
            }
            
            return textTable;
        } catch (error) {
            console.warn('Failed to convert HTML table to plain text:', error);
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
        
        if (window.debugLog) window.debugLog('formatTextForDisplay - Original:', JSON.stringify(text));
        if (window.debugLog) window.debugLog('formatTextForDisplay - Formatted:', JSON.stringify(formatted));
        
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
        
        if (window.debugLog) window.debugLog('formatTextForClipboard Original:', JSON.stringify(text));
        if (window.debugLog) window.debugLog('formatTextForClipboard Formatted:', JSON.stringify(formatted));
        
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
        
        // Determine the text to display based on response type
        let responseText = '';
        if (this.currentResponse) {
            if (this.currentResponse.type === 'followup') {
                // For follow-up suggestions, use the suggestions property
                responseText = this.currentResponse.suggestions || 'No follow-up suggestions available.';
            } else {
                // For regular responses, use the text property
                responseText = this.currentResponse.text || 'No response available.';
            }
        } else {
            responseText = 'No response available.';
        }
        
        // Add the AI's response as the first message in chat
        this.addChatMessage('assistant', responseText);
        
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
        if (window.debugLog) window.debugLog('onModelServiceChange triggered:', {
            value: event.target.value,
            oldValue: event.target.dataset.oldValue || 'undefined'
        });
        
        // If provider settings are already loading, skip this change to prevent race conditions
        if (this.isLoadingProviderSettings) {
            window.debugLog(`Skipping model service change - provider settings already loading`);
            return;
        }
        
        // Immediately clear model dropdown to prevent showing stale data from previous provider
        this.clearModelDropdownState();
        if (this.modelSelect) {
            this.modelSelect.innerHTML = '<option value="">Loading...</option>';
        }
        
        const customEndpoint = document.getElementById('custom-endpoint');
        if (customEndpoint) {
            if (event.target.value === 'custom') {
                customEndpoint.classList.remove('hidden');
            } else {
                customEndpoint.classList.add('hidden');
            }
        }
        
        // Allow provider changes during session, but don't persist the choice
        const newProvider = event.target.value;
        const oldProvider = event.target.dataset.oldValue;
        
        if (newProvider && newProvider !== 'undefined' && newProvider !== oldProvider) {
            // This is a real user change - allow it for this session only
            window.debugLog(`Provider changed for session: ${oldProvider} -> ${newProvider} (not persisted)`);
        }
        
        // Save current provider's settings before switching (use cached values, not form values)
        if (oldProvider && oldProvider !== 'undefined') {
            await this.saveCurrentProviderSettings(oldProvider, false);
        }
        
        // Load new provider's settings
        await this.loadProviderSettings(event.target.value);
        
        // Force refresh settings cache after loading provider settings
        await this.settingsManager.loadSettings();
        
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
    }

    async onSettingsProviderChange(event) {
        const selectedProvider = event.target.value;
        if (window.debugLog) window.debugLog('Settings provider changed to:', selectedProvider);
        
        // Load the settings for the selected provider (settings-only version)
        await this.loadSettingsOnlyProviderConfig(selectedProvider);
        
        // Update provider labels in UI to reflect the selected provider
        this.updateProviderLabels(selectedProvider);
        
        // Reset Test Connection button state when switching providers
        this.resetTestConnectionButton();
    }

    async onModelChange(event) {
        if (window.debugLog) window.debugLog('onModelChange triggered:', {
            value: event.target.value,
            oldValue: event.target.dataset.oldValue || 'undefined',
            provider: this.modelServiceSelect?.value
        });
        
        // Only clear if this is actually a change (not just initialization)
        const oldModel = event.target.dataset.oldValue;
        if (oldModel && oldModel !== 'undefined' && oldModel !== event.target.value) {
            // Clear previous analysis and response since model changed
            this.clearAnalysisAndResponse();
            if (window.debugLog) window.debugLog('Cleared analysis and response due to model change');
        }
        
        // Store old value for next time
        event.target.dataset.oldValue = event.target.value;
        
        // No longer saving model selections - always use ai-providers.json defaults
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
            this.settingsManager.loadSettings().then(async () => {
                if (window.debugLog) {
                    window.debugLog('Settings refreshed after closing settings panel');
                }
                
                // Refresh the model dropdown to clear any previous fetch errors
                // and populate with models from the now-properly-configured provider
                try {
                    const currentProvider = this.modelServiceSelect?.value;
                    if (currentProvider) {
                        await this.updateModelDropdown();
                        if (window.debugLog) {
                            window.debugLog('Model dropdown refreshed after closing settings');
                        }
                    }
                } catch (error) {
                    console.warn('Failed to refresh model dropdown after closing settings:', error);
                }
            }).catch(error => {
                console.warn('Failed to refresh settings after closing settings panel:', error);
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
                '• Delete all writing samples\n' +
                '• Reset all preferences to default values\n' +
                '• Clear all custom configurations\n' +
                '• Reset to default provider and model\n' +
                '• Remove data from both Office.js and browser storage\n\n' +
                'This action cannot be undone.'
            );
            
            if (!confirmed) {
                return;
            }
            
            // Use SettingsManager to properly clear all settings
            const success = await this.settingsManager.clearAllSettings();
            
            if (success) {
                // Get default provider based on user's domain, not hardcoded fallback
                let defaultProvider = 'ollama'; // Final fallback only
                let defaultModel = '';
                
                try {
                    // Get user's email for domain-based provider selection
                    const userProfile = Office.context.mailbox.userProfile;
                    if (userProfile && userProfile.emailAddress) {
                        const { defaultProvider: domainDefault } = this.getProvidersByDomain(userProfile);
                        defaultProvider = domainDefault;
                        console.info(`Reset settings using domain-based default provider: ${defaultProvider} for ${userProfile.emailAddress}`);
                    } else {
                        // Fallback to config default if user profile unavailable
                        defaultProvider = this.defaultProvidersConfig?._config?.defaultProvider || 'ollama';
                        console.warn(`User profile unavailable for reset, using config default: ${defaultProvider}`);
                    }
                } catch (error) {
                    console.warn(`Could not determine domain-based provider for reset, using fallback: ${error.message}`);
                    defaultProvider = this.defaultProvidersConfig?._config?.defaultProvider || 'ollama';
                }
                
                defaultModel = this.getDefaultModelForProvider(defaultProvider);
                
                // Set default provider and model
                if (this.modelServiceSelect) {
                    this.modelServiceSelect.value = defaultProvider;
                }
                
                // Save the reset settings with clean provider configs
                // Get current defaults (just restored by clearAllSettings) and only update provider settings
                const currentSettings = this.settingsManager.getSettings();
                await this.settingsManager.saveSettings({
                    ...currentSettings,  // Preserve all default settings including high-contrast: false
                    'model-service': defaultProvider,
                    'model-select': defaultModel,
                    'provider-configs': {
                        'openai': { 'api-key': '', 'endpoint-url': '' },    // Empty = use ai-providers.json defaults
                        'ollama': { 'api-key': '', 'endpoint-url': '' },    // Empty = use ai-providers.json defaults
                        'onsite1': { 'api-key': '', 'endpoint-url': '' },   // Empty = use ai-providers.json defaults
                        'onsite2': { 'api-key': '', 'endpoint-url': '' }    // Empty = use ai-providers.json defaults
                    }
                });
                
                // Wait a moment for settings to fully save
                await new Promise(resolve => setTimeout(resolve, 100));
                
                // Load the provider settings to populate UI with correct defaults
                await this.loadProviderSettings(defaultProvider);
                
                // Check if the default provider requires an API key
                const needsApiKey = this.providerNeedsApiKey(defaultProvider);
                
                // Get the provider label for display
                const providerLabel = this.defaultProvidersConfig?.[defaultProvider]?.label || defaultProvider;
                
                if (needsApiKey) {
                    // Show success message with API key instruction
                    await this.showInfoDialog('Settings Reset - Action Required', 
                        `Settings have been reset to defaults.\n\nDefault provider: ${providerLabel}\nDefault model: ${defaultModel}\n\n⚠️ IMPORTANT: This provider requires an API key.\n\nAfter the page reloads:\n1. Open Settings (⚙️)\n2. Enter your ${providerLabel} API key\n3. Close Settings to save\n\nThe application will now reload.`);
                } else {
                    // Show success message for local providers
                    await this.showInfoDialog('Success', 
                        `Settings have been reset to defaults.\n\nDefault provider: ${providerLabel}\nDefault model: ${defaultModel}\n\nThe application will now reload.`);
                }
                
                window.location.reload();
            } else {
                // Show error message
                await this.showInfoDialog('Error', 'Failed to reset settings. Please try again or contact support.');
            }
            
        } catch (error) {
            console.error('Error during settings reset:', error);
            await this.showInfoDialog('Error', 'An error occurred while resetting settings. Please try again.');
        }
    }

    // Simple dialog replacement for Office Add-in environment
    showConfirmDialog(title, message) {
        return new Promise((resolve) => {
            // Create a simple overlay dialog since Office Add-ins don't support native dialogs
            const isHighContrast = document.body.classList.contains('high-contrast');
            
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: ${isHighContrast ? 'rgba(0,0,0,0.9)' : 'rgba(0,0,0,0.5)'}; z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: ${isHighContrast ? '#000000' : 'white'}; 
                color: ${isHighContrast ? '#ffffff' : '#000000'};
                border: ${isHighContrast ? '2px solid #ffffff' : 'none'};
                padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            const titleColor = isHighContrast ? '#ffffff' : '#d73502';
            const yesButtonBg = isHighContrast ? '#ffff00' : '#d73502';
            const yesButtonColor = isHighContrast ? '#000000' : 'white';
            const yesButtonBorder = isHighContrast ? '2px solid #ffffff' : 'none';
            const noButtonBg = isHighContrast ? '#000000' : '#ccc';
            const noButtonColor = isHighContrast ? '#ffffff' : 'black';
            const noButtonBorder = isHighContrast ? '2px solid #ffffff' : 'none';
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: ${titleColor};">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0; color: ${isHighContrast ? '#ffffff' : '#000000'};">${message}</p>
                <div style="margin-top: 20px;">
                    <button id="confirm-yes" style="margin-right: 10px; padding: 8px 16px; background: ${yesButtonBg}; color: ${yesButtonColor}; border: ${yesButtonBorder}; border-radius: 4px; cursor: pointer;">Reset Settings</button>
                    <button id="confirm-no" style="padding: 8px 16px; background: ${noButtonBg}; color: ${noButtonColor}; border: ${noButtonBorder}; border-radius: 4px; cursor: pointer;">Cancel</button>
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
            const isHighContrast = document.body.classList.contains('high-contrast');
            
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: ${isHighContrast ? 'rgba(0,0,0,0.9)' : 'rgba(0,0,0,0.5)'}; z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: ${isHighContrast ? '#000000' : 'white'}; 
                color: ${isHighContrast ? '#ffffff' : '#000000'};
                border: ${isHighContrast ? '2px solid #ffffff' : 'none'};
                padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            const titleColor = isHighContrast 
                ? '#ffffff' 
                : (title === 'Error' ? '#d73502' : '#0078d4');
            const buttonBg = isHighContrast ? '#ffff00' : '#0078d4';
            const buttonColor = isHighContrast ? '#000000' : 'white';
            const buttonBorder = isHighContrast ? '2px solid #ffffff' : 'none';
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: ${titleColor};">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0; color: ${isHighContrast ? '#ffffff' : '#000000'};">${message}</p>
                <div style="margin-top: 20px;">
                    <button id="info-ok" style="padding: 8px 16px; background: ${buttonBg}; color: ${buttonColor}; border: ${buttonBorder}; border-radius: 4px; cursor: pointer;">OK</button>
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
            const isHighContrast = document.body.classList.contains('high-contrast');
            
            const overlay = document.createElement('div');
            overlay.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                background: ${isHighContrast ? 'rgba(0,0,0,0.9)' : 'rgba(0,0,0,0.5)'}; z-index: 10000; display: flex; 
                align-items: center; justify-content: center;
            `;
            
            const dialog = document.createElement('div');
            dialog.style.cssText = `
                background: ${isHighContrast ? '#000000' : 'white'}; 
                color: ${isHighContrast ? '#ffffff' : '#000000'};
                border: ${isHighContrast ? '2px solid #ffffff' : 'none'};
                padding: 20px; border-radius: 8px; max-width: 400px; 
                box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center;
            `;
            
            const titleColor = isHighContrast ? '#ffffff' : '#0078d4';
            const linkColor = isHighContrast ? '#ffff00' : '#0078d4';
            const descriptionColor = isHighContrast ? '#cccccc' : '#666';
            const openButtonBg = isHighContrast ? '#ffff00' : '#0078d4';
            const openButtonColor = isHighContrast ? '#000000' : 'white';
            const openButtonBorder = isHighContrast ? '2px solid #ffffff' : 'none';
            const closeButtonBg = isHighContrast ? '#000000' : '#6c757d';
            const closeButtonColor = isHighContrast ? '#ffffff' : 'white';
            const closeButtonBorder = isHighContrast ? '2px solid #ffffff' : 'none';
            
            dialog.innerHTML = `
                <h3 style="margin-top: 0; color: ${titleColor};">${title}</h3>
                <p style="white-space: pre-line; margin: 16px 0; text-align: left; color: ${isHighContrast ? '#ffffff' : '#000000'};">${message}</p>
                <p style="margin: 16px 0; font-size: 14px; color: ${descriptionColor};">
                    <strong>Help URL:</strong><br>
                    <a href="${helpUrl}" target="_blank" style="color: ${linkColor}; word-break: break-all;">${helpUrl}</a>
                </p>
                <div style="margin-top: 20px; display: flex; gap: 10px; justify-content: center;">
                    <button id="help-open" style="padding: 8px 16px; background: ${openButtonBg}; color: ${openButtonColor}; border: ${openButtonBorder}; border-radius: 4px; cursor: pointer;">Open Help</button>
                    <button id="help-close" style="padding: 8px 16px; background: ${closeButtonBg}; color: ${closeButtonColor}; border: ${closeButtonBorder}; border-radius: 4px; cursor: pointer;">Close</button>
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
            console.error('Test connection button elements not found');
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

            console.info(`Testing connection for provider: ${currentProvider}`);
            
            // Test the connection using AIService
            const success = await this.aiService.testConnection(config);
            
            if (success) {
                // First save the settings since the test was successful
                await this.saveCurrentProviderSettingsSimple(currentProvider);
                console.info(`Settings saved for provider: ${currentProvider}`);
                
                // Success state
                button.classList.remove('testing');
                button.classList.add('success');
                buttonText.textContent = '✓ Saved';
                // No popup needed - button visual feedback is sufficient
                
                console.info(`Connection test passed for ${currentProvider}`);
                
                // Log success event
                this.logger.logEvent('connection_test_success', {
                    provider: currentProvider,
                    endpoint: config.endpointUrl,
                    has_api_key: !!config.apiKey
                }, 'Information', this.getUserEmailForTelemetry());
                
                // Reset button after 3 seconds
                setTimeout(() => {
                    button.classList.remove('success');
                    buttonText.textContent = 'Save & Test';
                }, 3000);
                
            } else {
                throw new Error('Connection test failed - no response received');
            }
            
        } catch (error) {
            console.error(`Connection test failed for ${currentProvider}:`, error);
            
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
                buttonText.textContent = 'Save & Test';
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
            buttonText.textContent = 'Save & Test';
            button.disabled = false;
            buttonSpinner.classList.add('hidden');
        }
    }

    async showProviderHelp() {
        // Get current provider from settings dropdown (since help button is in settings panel)
        const settingsProviderSelect = document.getElementById('settings-provider-select');
        const currentProvider = settingsProviderSelect?.value || 'ollama';
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
                        console.error('Error opening help URL:', error);
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
            await this.showInfoDialog('Help', 'No provider configuration found.');
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

    toggleHighContrast(enabled, skipSave = false) {
        if (window.debugLog) window.debugLog('toggleHighContrast called:', enabled, 'skipSave:', skipSave);
        document.body.classList.toggle('high-contrast', enabled);
        
        // Sync the checkbox state with the body class
        const highContrastCheckbox = document.getElementById('high-contrast');
        if (highContrastCheckbox) {
            highContrastCheckbox.checked = enabled;
            if (window.debugLog) window.debugLog('Checkbox state synced to:', enabled);
        }
        
        if (window.debugLog) window.debugLog('body classes after toggle:', document.body.classList.toString());
        
        // Only save settings if not explicitly skipping
        if (!skipSave) {
            if (window.debugLog) window.debugLog('Saving high-contrast setting:', enabled);
            this.saveSettings();
        } else {
            if (window.debugLog) window.debugLog('Skipping save for high-contrast setting (loading mode)');
        }
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
    
    // Early access acknowledged
    const earlyAccessCheckbox = document.getElementById('early-access-acknowledged');
    if (earlyAccessCheckbox) earlyAccessCheckbox.checked = !!settings['early-access-acknowledged'];

        // Load form values (excluding provider-specific fields)
        Object.keys(settings).forEach(key => {
            // Skip provider-specific fields 
            if (key === 'api-key' || key === 'endpoint-url' || key === 'provider-configs' || key === 'providers') return;
            const element = document.getElementById(key);
            if (element) {
                let value;
                if (element.type === 'checkbox') {
                    element.checked = settings[key];
                    value = settings[key];
                } else {
                    element.value = settings[key] || '';
                    value = settings[key] || '';
                }
                
                // Update UI state manager with loaded value
                this.uiStateManager.updateUIFormValue(key, value);
            }
        });

        // Always use the default provider from ai-providers.json, but allow UI changes during session
        const modelServiceSelect = document.getElementById('model-service');
        if (modelServiceSelect) {
            const defaultProvider = this.getDefaultProvider();
            
            // Find the default provider in dropdown options
            let foundOption = false;
            for (let option of modelServiceSelect.options) {
                if (option.value === defaultProvider) {
                    modelServiceSelect.value = defaultProvider;
                    foundOption = true;
                    break;
                }
            }
            
            // Fall back to first option if default not found
            if (!foundOption && modelServiceSelect.options.length > 0) {
                modelServiceSelect.value = modelServiceSelect.options[0].value;
            }
            
            // Allow changes during session (don't disable dropdown)
            modelServiceSelect.disabled = false;
            
            window.debugLog(`Provider defaulted to: ${modelServiceSelect.value} (changeable during session)`);
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
            

            
            // Load the provider config for the settings panel to populate endpoint URL correctly
            await this.loadSettingsOnlyProviderConfig(settingsProviderSelect.value);
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
            if (window.debugLog) window.debugLog('Applying high contrast setting on load:', settings['high-contrast']);
            this.toggleHighContrast(true, true);
        }

        if (settings['screen-reader-mode']) {
            this.toggleScreenReaderMode(true);
        }
        
        // Refresh models for the current provider to ensure they match
        // This prevents stale models from previous provider configurations
        await this.updateModelDropdown();
    }

    async saveSettings() {
        const formSettings = {};
        
        // Collect all form values except provider-specific ones
        const inputs = document.querySelectorAll('input, select, textarea');
        if (window.debugLog) window.debugLog('saveSettings: Found', inputs.length, 'form elements');
        
        inputs.forEach((input, index) => {
            if (input.id) {
                // Skip provider-specific fields as they're handled separately
                if (input.id === 'api-key' || input.id === 'endpoint-url') {
                    return;
                }
                
                const value = input.type === 'checkbox' ? input.checked : input.value;
                formSettings[input.id] = value;
                
                if (input.id === 'model-service') {
                    if (window.debugLog) window.debugLog('model-service element details:', {
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
        
        if (window.debugLog) window.debugLog('saveSettings collected:', formSettings);
        
        // Merge form settings with existing settings to preserve provider-configs
        const currentSettings = this.settingsManager.getSettings();
        const mergedSettings = { ...currentSettings, ...formSettings };
        
        await this.settingsManager.saveSettings(mergedSettings, 'form settings');
        
        // Also save current UI state
        await this.saveCurrentUIState();
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

                
                // IMPORTANT: Only save the fields without cross-provider validation/correction
                // This prevents the base URL from being incorrectly changed when switching between API key and endpoint fields
                await this.saveCurrentProviderSettingsSimple(settingsProvider);
            } else {
                console.warn(`Settings panel open but no provider selected in settings dropdown`);
 gd            }
        } else {
            // We're in main taskpane context - save to main provider and update models
            const mainProvider = this.modelServiceSelect?.value;
            if (mainProvider && mainProvider !== 'undefined') {

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
        
        // Use the API key as provided by the user - no validation or contamination checking
        // This allows legitimate cases where API keys match provider names (e.g., local test environments)
        const finalApiKey = apiKey;
        
        // Since endpoint URL is now read-only, don't save user input - always use empty string
        // This ensures we always fall back to ai-providers.json defaults
        const endpointUrl = '';
        
        // Simple save - no validation or correction, just save what the user entered
        window.debugLog(`[VERBOSE] Simple save for provider ${provider}:`, { 
            apiKey: finalApiKey.length ? '[HIDDEN]' : '[EMPTY]',
            endpointUrl: 'ALWAYS_EMPTY (using ai-providers.json default)'
        });
        
        await this.settingsManager.setProviderConfig(provider, finalApiKey, endpointUrl);

    }

    async saveCurrentProviderSettings(provider, useFormValues = true) {
        if (!provider || provider === 'undefined') return;
        
        let apiKey, endpointUrl;
        
        if (useFormValues) {
            // Use current form values (for normal saves)
            const apiKeyElement = document.getElementById('api-key');
            const endpointUrlElement = document.getElementById('endpoint-url');
            
            apiKey = apiKeyElement ? apiKeyElement.value.trim() : '';
            endpointUrl = ''; // Always empty, forces use of defaults
        } else {
            // Use cached values from settings (for provider switches)
            const providerConfig = this.settingsManager.getProviderConfig(provider);
            apiKey = providerConfig['api-key'] || '';
            endpointUrl = ''; // Always empty, forces use of defaults
        }
        
        // Use the API key as provided - no validation or contamination checking
        // This allows legitimate cases where API keys match provider names (e.g., local test environments)
        const defaultConfig = this.defaultProvidersConfig?.[provider];
        const finalApiKey = apiKey;
        
        // Since endpoint URL is now read-only, don't save user input - always use empty string
        // This ensures we always fall back to ai-providers.json defaults
        const finalEndpointUrl = ''; // Always empty, forces use of defaults
        
        // Simple save - no validation or correction, just save what the user entered
        window.debugLog(`[VERBOSE] Simple save for provider ${provider}:`, { 
            apiKey: finalApiKey.length ? '[HIDDEN]' : '[EMPTY]',
            endpointUrl: 'ALWAYS_EMPTY (using ai-providers.json default)'
        });
        
        await this.settingsManager.setProviderConfig(provider, finalApiKey, finalEndpointUrl);
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

        
        const apiKeyElement = document.getElementById('api-key');
        const endpointUrlElement = document.getElementById('endpoint-url');
        
        if (apiKeyElement) {
            let apiKeyToUse = providerConfig['api-key'] || '';
            
            // Use the stored API key as-is, without clearing it for any provider
            // This allows legitimate cases where API keys match provider names
            const defaultConfig = this.defaultProvidersConfig?.[provider];
            if (defaultConfig && defaultConfig.apiFormat === 'ollama') {
                // Add placeholder to indicate Ollama may not need an API key (but allow user override)
                apiKeyElement.placeholder = 'API key (optional for Ollama)';
            } else {
                // Reset placeholder for non-Ollama providers
                apiKeyElement.placeholder = 'Enter your API key';
            }
            
            apiKeyElement.value = apiKeyToUse;

        }
        
        if (endpointUrlElement) {
            let endpointToUse = providerConfig['endpoint-url'] || '';
            
            // For display, show stored value or default if none exists
            if (!endpointToUse && this.defaultProvidersConfig && this.defaultProvidersConfig[provider]) {
                endpointToUse = this.defaultProvidersConfig[provider].baseUrl || '';
            }
            
            endpointUrlElement.value = endpointToUse;

        }
        

    }

    async loadProviderSettings(provider) {
        if (!provider || provider === 'undefined') return;
        
        // Prevent overlapping provider settings loading
        if (this.isLoadingProviderSettings) {
            window.debugLog(`Skipping overlapping provider settings load for ${provider}`);
            return;
        }
        
        // Set flag to prevent blur events and overlapping loads during loading
        this.isLoadingProviderSettings = true;
        
        try {
            const providerConfig = this.settingsManager.getProviderConfig(provider);
            
            const apiKeyElement = document.getElementById('api-key');
        
            if (apiKeyElement) {
                let apiKeyToUse = providerConfig['api-key'] || '';
                
                // Use the stored API key as-is, without any contamination checking
                // This allows legitimate cases where API keys match provider names
                const defaultConfig = this.defaultProvidersConfig?.[provider];
                if (defaultConfig && defaultConfig.apiFormat === 'ollama') {
                    // Add placeholder to indicate Ollama may not need an API key (but allow user override)
                    apiKeyElement.placeholder = 'API key (optional for Ollama)';
                } else {
                    // Reset placeholder for non-Ollama providers
                    apiKeyElement.placeholder = 'Enter your API key';
                }
                
                apiKeyElement.value = apiKeyToUse;
                
                // Update UI state manager with provider-specific values
                this.uiStateManager.updateUIFormValue('api-key', apiKeyToUse);
            }

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
        
        if (apiKeyLabel) {
            apiKeyLabel.textContent = `(${providerLabel})`;
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
            console.warn('Unable to get user profile for telemetry:', error);
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
        console.error('Failed to initialize application:', error);
    });
});
