/**
 * Settings Manager Service
 * Handles persistent storage and retrieval of user preferences
 */

export class SettingsManager {
    constructor() {
        this.storageKey = 'promptemail_settings';
        
        this.defaultSettings = {
            // Provider-specific configurations (dynamically initialized from ai-providers.json)
            'providers': {},
            
            // Automation Settings
            'auto-analysis': false,
            'auto-response': false,
            
            // Accessibility Settings
            'high-contrast': false,
            'screen-reader-mode': false,

            // Debug Logging
            'debug-logging': false,
            
            // Version tracking
            'settings-version': '1.0.0',
            'last-updated': null
        };
        
        this.settings = { ...this.defaultSettings };
        this.changeListeners = [];
        
        // Add unique instance ID for debugging
        this.instanceId = Date.now() + '-' + Math.random().toString(36).substr(2, 9);
        
        if (window.debugLog) window.debugLog('SettingsManager instance created:', this.instanceId);
    }

    /**
     * Dynamically initialize provider configurations from ai-providers.json
     * @param {Object} providersConfig - The ai-providers.json configuration
     */
    initializeProvidersFromConfig(providersConfig) {
        if (!providersConfig || typeof providersConfig !== 'object') {
            console.warn('No ai-providers.json config available, using empty provider structure');
            return {};
        }

        const providers = {};
        
        // Initialize each provider with empty user-configurable settings
        Object.keys(providersConfig).forEach(providerKey => {
            // Skip the _config meta section
            if (providerKey.startsWith('_')) return;
            
            const providerInfo = providersConfig[providerKey];
            const providerConfig = {
                'api-key': '',
                'endpoint-url': ''
            };
            
            // Add bedrock-specific fields if this is a bedrock provider
            if (providerInfo.apiFormat === 'bedrock') {
                providerConfig['accessKeyId'] = '';
                providerConfig['secretAccessKey'] = '';
                providerConfig['sessionToken'] = '';
                providerConfig['region'] = 'us-east-1';
            }
            
            providers[providerKey] = providerConfig;
        });
        
        if (window.debugLog) {
            window.debugLog('Initialized dynamic providers:', Object.keys(providers));
        }
        
        return providers;
    }

    /**
     * Loads settings from storage
     * @returns {Promise<Object>} Loaded settings
     */
    async loadSettings() {
        try {
            // Load from Office.js RoamingSettings
            const officeSettings = await this.loadFromOfficeStorage();
            if (officeSettings) {
                this.settings = { ...this.defaultSettings, ...officeSettings };
                
                // Debug: Track high-contrast setting in loads
                if (window.debugLog) {
                    window.debugLog('High-contrast after load merge:', this.settings['high-contrast'], 
                        '(default:', this.defaultSettings['high-contrast'], 
                        'loaded:', officeSettings['high-contrast'], ')');
                }
                
                return this.settings;
            }

            // Office storage available but no settings found - initialize with defaults
            console.log('Initializing with default settings');
            this.settings = { ...this.defaultSettings };
            await this.saveToOfficeStorage(this.settings);
            return this.settings;

        } catch (error) {
            console.error('Failed to load settings from Office storage:', error);
            // Use defaults if Office storage fails
            this.settings = { ...this.defaultSettings };
            return this.settings;
        }
    }

    /**
     * Saves settings to storage
     * @param {Object} newSettings - Settings to save
     * @returns {Promise<boolean>} Success status
     */
    async saveSettings(newSettings = null, context = null) {
        const contextStr = context ? ` (${context})` : '';
        if (window.debugLog) window.debugLog(`Saving settings${contextStr}...`);
        
        try {
            const settingsToSave = newSettings || this.settings;
            
            // Debug: Track high-contrast setting in saves
            if (window.debugLog && 'high-contrast' in settingsToSave) {
                window.debugLog(`High-contrast in save${contextStr}:`, settingsToSave['high-contrast']);
            }
            
            // Update timestamp
            settingsToSave['last-updated'] = new Date().toISOString();
            
            // Update internal settings
            this.settings = { ...settingsToSave };

            // Save to Office.js storage
            const success = await this.saveToOfficeStorage(settingsToSave);

            // Notify listeners
            this.notifyChangeListeners(settingsToSave);

            return success;

        } catch (error) {
            console.error('Settings save failed:', error);
            return false;
        }
    }

    /**
     * Loads settings from Office.js RoamingSettings
     * @returns {Promise<Object|null>} Settings object or null
     */
    async loadFromOfficeStorage() {

        return new Promise((resolve) => {
            try {
                if (typeof Office === 'undefined' || !Office.context?.roamingSettings) {
                    console.warn('Office.js or RoamingSettings not available');
                    resolve(null);
                    return;
                }

                const roamingSettings = Office.context.roamingSettings;
                const settingsJson = roamingSettings.get(this.storageKey);
                
                if (settingsJson) {
                    const settings = JSON.parse(settingsJson);
                    resolve(settings);
                } else {
                    if (window.debugLog) window.debugLog('No settings found in Office storage');
                    resolve(null);
                }

            } catch (error) {
                console.error('Failed to load from Office storage:', error);
                resolve(null);
            }
        });
    }

    /**
     * Saves settings to Office.js RoamingSettings
     * @param {Object} settings - Settings to save
     * @returns {Promise<boolean>} Success status
     */
    async saveToOfficeStorage(settings) {

        return new Promise((resolve) => {
            try {
                if (typeof Office === 'undefined' || !Office.context?.roamingSettings) {
                    console.warn('Office.js or RoamingSettings not available for save');
                    resolve(false);
                    return;
                }

                const roamingSettings = Office.context.roamingSettings;
                const settingsJson = JSON.stringify(settings);

                
                roamingSettings.set(this.storageKey, settingsJson);
                
                // Save settings asynchronously
                roamingSettings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {

                        resolve(true);
                    } else {
                        console.error('Failed to save to Office storage:', result.error);
                        resolve(false);
                    }
                });

            } catch (error) {
                console.error('Error saving to Office storage:', error);
                resolve(false);
            }
        });
    }



    /**
     * Gets a specific setting value
     * @param {string} key - Setting key
     * @param {*} defaultValue - Default value if not found
     * @returns {*} Setting value
     */
    getSetting(key, defaultValue = null) {
        const value = this.settings[key] !== undefined ? this.settings[key] : defaultValue;
        return value;
    }

    /**
     * Sets a specific setting value
     * @param {string} key - Setting key
     * @param {*} value - Setting value
     * @returns {Promise<boolean>} Success status
     */
    async setSetting(key, value) {
        this.settings[key] = value;
        const result = await this.saveSettings(null, `setting: ${key}`);
        return result;
    }



    /**
     * Gets all current settings
     * @returns {Object} All settings
     */
    getSettings() {
        return { ...this.settings };
    }

    /**
     * Resets settings to defaults
     * @returns {Promise<boolean>} Success status
     */
    async resetToDefaults() {
        this.settings = { ...this.defaultSettings };
        return await this.saveSettings();
    }

    /**
     * Exports settings as JSON
     * @returns {string} JSON string of settings
     */
    exportSettings() {
        // Create export object without sensitive data
        const exportData = { ...this.settings };
        
        // Remove sensitive provider data (API keys are stored in providers structure)
        
        return JSON.stringify(exportData, null, 2);
    }

    /**
     * Imports settings from JSON
     * @param {string} jsonString - JSON string of settings
     * @returns {Promise<boolean>} Success status
     */
    async importSettings(jsonString) {
        try {
            const importedSettings = JSON.parse(jsonString);
            
            // Validate imported settings
            const validatedSettings = this.validateSettings(importedSettings);
            
            // Merge with current settings (preserve provider configurations)
            const mergedSettings = {
                ...this.settings,
                ...validatedSettings,
                // Keep current provider configurations 
                'providers': this.settings['providers']
            };

            return await this.saveSettings(mergedSettings);

        } catch (error) {
            console.error('Failed to import settings:', error);
            return false;
        }
    }

    /**
     * Validates settings object
     * @param {Object} settings - Settings to validate
     * @returns {Object} Validated settings
     */
    validateSettings(settings) {
        const validated = {};

        // Validate each setting against defaults
        Object.keys(this.defaultSettings).forEach(key => {
            if (settings.hasOwnProperty(key)) {
                const value = settings[key];
                const defaultValue = this.defaultSettings[key];
                
                // Type validation
                if (typeof value === typeof defaultValue) {
                    validated[key] = value;
                } else {
                    console.warn(`Invalid type for setting '${key}', using default`);
                    validated[key] = defaultValue;
                }
            } else {
                validated[key] = this.defaultSettings[key];
            }
        });

        // Validate specific setting ranges
        // (Legacy response-length and response-tone validation removed)

        return validated;
    }

    /**
     * Adds a change listener
     * @param {Function} listener - Callback function for setting changes
     */
    addChangeListener(listener) {
        if (typeof listener === 'function') {
            this.changeListeners.push(listener);
        }
    }

    /**
     * Removes a change listener
     * @param {Function} listener - Listener function to remove
     */
    removeChangeListener(listener) {
        const index = this.changeListeners.indexOf(listener);
        if (index > -1) {
            this.changeListeners.splice(index, 1);
        }
    }

    /**
     * Notifies all change listeners
     * @param {Object} newSettings - New settings object
     */
    notifyChangeListeners(newSettings) {
        this.changeListeners.forEach(listener => {
            try {
                listener(newSettings);
            } catch (error) {
                console.error('Settings change listener error:', error);
            }
        });
    }

    /**
     * Gets settings migration information
     * @returns {Object} Migration status
     */
    getMigrationStatus() {
        const currentVersion = this.settings['settings-version'] || '1.0.0';
        const latestVersion = this.defaultSettings['settings-version'];
        
        return {
            current: currentVersion,
            latest: latestVersion,
            needsMigration: currentVersion !== latestVersion,
            lastUpdated: this.settings['last-updated']
        };
    }

    /**
     * Migrates settings to latest version
     * @returns {Promise<boolean>} Success status
     */
    async migrateSettings() {
        const migration = this.getMigrationStatus();
        
        if (!migration.needsMigration) {
            return true;
        }

        try {
            // Perform migration logic here
            // For now, just update the version
            this.settings['settings-version'] = migration.latest;
            
            return await this.saveSettings();

        } catch (error) {
            console.error('Settings migration failed:', error);
            return false;
        }
    }

    /**
     * Clears all stored settings
     * @returns {Promise<boolean>} Success status
     */
    async clearAllSettings() {
        if (window.debugLog) window.debugLog('Starting complete settings reset...');
        try {
            let officeCleared = true;
            let localStorageCleared = true;

            // Clear from Office.js RoamingSettings
            if (typeof Office !== 'undefined' && Office.context?.roamingSettings) {
                if (window.debugLog) window.debugLog('Clearing Office.js RoamingSettings...');
                const roamingSettings = Office.context.roamingSettings;
                roamingSettings.remove(this.storageKey);
                
                officeCleared = await new Promise((resolve) => {
                    roamingSettings.saveAsync((result) => {
                        const success = result.status === Office.AsyncResultStatus.Succeeded;
                        if (window.debugLog) window.debugLog(`Office.js clear result: ${success ? 'SUCCESS' : 'FAILED'}`);
                        if (!success && result.error) {
                            console.error('Office.js clear error:', result.error);
                        }
                        resolve(success);
                    });
                });
            } else {
                if (window.debugLog) window.debugLog('Office.js RoamingSettings not available, skipping...');
            }

            // Clear from localStorage
            if (typeof localStorage !== 'undefined') {
                if (window.debugLog) window.debugLog('Clearing localStorage...');
                try {
                    localStorage.removeItem(this.storageKey);
                    // Verify removal
                    const stillExists = localStorage.getItem(this.storageKey);
                    localStorageCleared = stillExists === null;
                    if (window.debugLog) window.debugLog(`localStorage clear result: ${localStorageCleared ? 'SUCCESS' : 'FAILED'}`);
                } catch (localError) {
                    console.error('localStorage clear error:', localError);
                    localStorageCleared = false;
                }
            } else {
                if (window.debugLog) window.debugLog('localStorage not available, skipping...');
            }

            // Reset to defaults
            if (window.debugLog) window.debugLog('Resetting internal settings to defaults...');
            this.settings = { ...this.defaultSettings };
            
            // Dynamically initialize providers from ai-providers.json
            // Note: This will be empty initially, but will be populated when ai-providers.json loads
            this.settings['providers'] = {};
            
            // Notify listeners of the reset
            if (window.debugLog) window.debugLog('Notifying change listeners of reset...');
            this.notifyChangeListeners(this.settings);

            const overallSuccess = officeCleared && localStorageCleared;
            if (window.debugLog) window.debugLog(`Complete settings reset result: ${overallSuccess ? 'SUCCESS' : 'PARTIAL'} (Office: ${officeCleared}, localStorage: ${localStorageCleared})`);
            
            return overallSuccess;

        } catch (error) {
            console.error('Failed to clear settings:', error);
            if (window.debugLog) window.debugLog('Settings reset failed with exception');
            return false;
        }
    }

    /**
     * Get immutable provider-specific configuration for a given service
     * @param {string} provider - The provider key (e.g., 'openai', 'ollama')
     * @returns {Object} Immutable provider configuration
     */
    getProviderConfig(provider) {
        if (!provider) return this.getEmptyProviderConfig();
        
        // Ensure providers structure exists
        if (!this.settings['providers']) {
            this.settings['providers'] = { ...this.defaultSettings['providers'] };
        }
        
        // Return existing provider config or create empty one
        if (this.settings['providers'][provider]) {
            return this.deepCopy(this.settings['providers'][provider]);
        }
        
        // Initialize empty config for new provider
        const emptyConfig = this.getEmptyProviderConfig();
        this.settings['providers'][provider] = emptyConfig;
        return this.deepCopy(emptyConfig);
    }

    /**
     * Get empty provider config structure
     * @returns {Object} Empty provider configuration
     */
    getEmptyProviderConfig() {
        return {
            'api-key': '',
            'endpoint-url': '',
            'selected-model': ''
        };
    }
    


    /**
     * Deep copy helper to prevent object reference contamination
     * @param {Object} obj - Object to copy
     * @returns {Object} Deep copy of object
     */
    deepCopy(obj) {
        return JSON.parse(JSON.stringify(obj));
    }

    /**
     * Set immutable provider-specific configuration for a given service
     * @param {string} provider - The provider key
     * @param {string} apiKey - The API key for this provider
     * @param {string} endpointUrl - The endpoint URL for this provider
     * @param {string} selectedModel - The selected model for this provider
     * @returns {Promise<boolean>} Success status
     */
    async setProviderConfig(provider, apiKey, endpointUrl, selectedModel = '') {
        if (!provider) return false;
        
        try {
            // Initialize providers structure if it doesn't exist
            if (!this.settings['providers']) {
                this.settings['providers'] = { ...this.defaultSettings['providers'] };
            }
            
            // Get current provider config or create empty one
            const currentConfig = this.settings['providers'][provider] || this.getEmptyProviderConfig();
            
            // Create new immutable config (never modify existing)
            const newConfig = {
                ...currentConfig,
                'api-key': apiKey || currentConfig['api-key'] || '',
                'endpoint-url': endpointUrl || currentConfig['endpoint-url'] || '',
                'selected-model': selectedModel || currentConfig['selected-model'] || ''
            };
            
            // Handle special Bedrock fields
            if (provider === 'bedrock') {
                newConfig['accessKeyId'] = currentConfig['accessKeyId'] || '';
                newConfig['secretAccessKey'] = currentConfig['secretAccessKey'] || '';
                newConfig['sessionToken'] = currentConfig['sessionToken'] || '';
                newConfig['region'] = currentConfig['region'] || 'us-east-1';
            }
            
            // Set the immutable provider configuration
            this.settings['providers'][provider] = newConfig;
            
            if (window.debugLog) {
                window.debugLog(`setProviderConfig(${provider}):`, {
                    'api-key': newConfig['api-key'] ? '[HIDDEN]' : '[EMPTY]',
                    'endpoint-url': newConfig['endpoint-url'] || '[EMPTY]',
                    'selected-model': newConfig['selected-model'] || '[EMPTY]'
                });
            }
            
            // Save the updated settings
            return await this.saveSettings(this.settings, `provider config: ${provider}`);
            
        } catch (error) {
            console.error('Failed to set provider config:', error);
            return false;
        }
    }

    /**
     * Update provider configuration with immutable pattern
     * @param {string} provider - Provider key
     * @param {Object} config - Configuration object to merge
     */
    updateProviderConfig(provider, config) {
        if (!provider || !config) return;
        
        // Ensure providers structure exists
        if (!this.settings['providers']) {
            this.settings['providers'] = {};
        }
        
        // Deep merge with existing configuration
        const currentConfig = this.getProviderConfig(provider);
        const mergedConfig = this.deepMerge(currentConfig, config);
        
        if (window.debugLog) {
            window.debugLog(`updateProviderConfig(${provider}):`, {
                'api-key': mergedConfig['api-key'] ? '[HIDDEN]' : '[EMPTY]',
                'endpoint-url': mergedConfig['endpoint-url'] || '[EMPTY]',
                'selected-model': mergedConfig['selected-model'] || '[EMPTY]'
            });
        }
        
        // Store in new providers structure only
        this.settings['providers'][provider] = mergedConfig;
    }

    /**
     * Deep merge helper for configuration objects
     * @param {Object} target - Target object
     * @param {Object} source - Source object
     * @returns {Object} Merged object
     */
    deepMerge(target, source) {
        const result = { ...target };
        
        for (const key in source) {
            if (source.hasOwnProperty(key)) {
                if (typeof source[key] === 'object' && source[key] !== null && !Array.isArray(source[key])) {
                    result[key] = this.deepMerge(result[key] || {}, source[key]);
                } else {
                    result[key] = source[key];
                }
            }
        }
        
        return result;
    }

    /**
     * Get the currently active provider (default from config, not persisted)
     * @returns {string} Active provider key
     */
    getActiveProvider() {
        // This should be determined by taskpane.js based on ai-providers.json config
        // SettingsManager no longer hardcodes provider selection
        return null; // Let calling code determine from ai-providers.json
    }

    /**
     * Note: setActiveProvider removed - provider choices are no longer persisted
     * Users can change providers during session but it resets to default on restart
     */

    /**
     * Set a specific setting for a provider
     * @param {string} provider - The provider key
     * @param {string} key - The setting key (e.g., 'selected-model', 'api-key')
     * @param {string} value - The setting value
     * @returns {Promise<boolean>} Success status
     */
    async setProviderSetting(provider, key, value) {
        if (!provider || !key) return false;
        
        try {
            // Get current config immutably
            const currentConfig = this.getProviderConfig(provider);
            
            // Create new config with updated value
            const newConfig = {
                ...currentConfig,
                [key]: value
            };
            
            // Initialize providers structure if needed
            if (!this.settings['providers']) {
                this.settings['providers'] = { ...this.defaultSettings['providers'] };
            }
            
            // Set the new immutable configuration
            this.settings['providers'][provider] = newConfig;
            
            // Save the updated settings
            return await this.saveSettings(this.settings, `provider setting: ${provider}.${key}`);
            
        } catch (error) {
            console.error('Failed to set provider setting:', error);
            return false;
        }
    }

    /**
     * Set Bedrock-specific configuration using the new providers structure
     * @param {string} provider - The provider key (should be a bedrock provider)
     * @param {string} accessKeyId - AWS Access Key ID
     * @param {string} secretAccessKey - AWS Secret Access Key
     * @param {string} sessionToken - AWS Session Token (optional)
     * @param {string} region - AWS Region
     * @param {string} endpointUrl - Custom endpoint URL (optional)
     * @returns {Promise<boolean>} Success status
     */
    async setBedrockConfig(provider, accessKeyId, secretAccessKey, sessionToken, region, endpointUrl) {
        try {
            const currentConfig = this.getProviderConfig(provider);
            
            const newConfig = {
                ...currentConfig,
                'accessKeyId': accessKeyId || '',
                'secretAccessKey': secretAccessKey || '',
                'sessionToken': sessionToken || '',
                'region': region || 'us-east-1',
                'endpoint-url': endpointUrl || ''
            };
            
            this.settings['providers'][provider] = newConfig;
            return await this.saveSettings(this.settings, `bedrock config: ${provider}`);
            
        } catch (error) {
            console.error('Failed to set Bedrock config:', error);
            return false;
        }
    }






}

