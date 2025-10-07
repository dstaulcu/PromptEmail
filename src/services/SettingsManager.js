/**
 * Settings Manager Service
 * Handles persistent storage and retrieval of user preferences
 */

export class SettingsManager {
    constructor() {
        this.storageKey = 'promptemail_settings';
        
        // Debug flag to disable Office.js storage for testing
        // TEMPORARY: Force localStorage-only mode for debugging
        this.forceLocalStorageOnly = false; // Set to false to re-enable Office.js storage
        
        // Alternative activation methods (if needed)
        this.forceLocalStorageOnly = this.forceLocalStorageOnly ||
                                     window.location.search.includes('debugStorage=localStorage') ||
                                     window.location.search.includes('debug-storage=local') || 
                                     localStorage.getItem('debug-force-local-storage') === 'true';
        
        if (this.forceLocalStorageOnly) {
            console.warn('[DEBUG] ========================================');
            console.warn('[DEBUG] OFFICE.JS STORAGE DISABLED FOR TESTING');
            console.warn('[DEBUG] Using localStorage ONLY - no sync with Office');
            console.warn('[DEBUG] ========================================');
        }
        
        this.defaultSettings = {
            // AI Configuration
            'model-service': 'openai',
            'api-key': '', // Legacy single API key for backwards compatibility
            'endpoint-url': '', // Legacy single endpoint for backwards compatibility
            
            // Provider-specific configurations (only user overrides, defaults come from ai-providers.json)
            'provider-configs': {
                'openai': { 'api-key': '', 'endpoint-url': '' },    // Empty = use ai-providers.json baseUrl
                'ollama': { 'api-key': '', 'endpoint-url': '' },    // Empty = use ai-providers.json baseUrl  
                'onsite1': { 'api-key': '', 'endpoint-url': '' },   // Empty = use ai-providers.json baseUrl
                'onsite2': { 'api-key': '', 'endpoint-url': '' }    // Empty = use ai-providers.json baseUrl
            },
            
            // Response Preferences
            'response-length': '1',
            'response-tone': '1',
            'custom-instructions': '',
            
            // Automation Settings
            'auto-analysis': false,
            'auto-response': false,
            
            // Accessibility Settings
            'high-contrast': false,
            'screen-reader-mode': false,

            // Debug Logging
            'debug-logging': false,
            
            // Writing Style Training
            'writing-samples': [],
            'style-analysis-enabled': false,
            'style-strength': 'medium', // light, medium, strong
            
            // UI Preferences
            'last-tab': 'analysis',
            'show-advanced-options': false,
            
            // Version tracking
            'settings-version': '1.0.0',
            'last-updated': null
        };
        
        this.settings = { ...this.defaultSettings };
        this.changeListeners = [];
        
        // Add unique instance ID for debugging
        this.instanceId = Date.now() + '-' + Math.random().toString(36).substr(2, 9);
        if (window.debugLog) window.debugLog('[VERBOSE] - SettingsManager instance created:', this.instanceId);
    }

    /**
     * Loads settings from storage
     * @returns {Promise<Object>} Loaded settings
     */
    async loadSettings() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Starting settings load process...');
        try {
            // Skip Office storage if debug flag is set
            if (this.forceLocalStorageOnly) {
                console.warn('[DEBUG] - SKIPPING Office storage - localStorage-only mode active');
            } else {
                if (window.debugLog) window.debugLog('[VERBOSE] - Attempting to load from Office storage...');
                // Try Office.js RoamingSettings first
                const officeSettings = await this.loadFromOfficeStorage();
                if (officeSettings) {
                    if (window.debugLog) window.debugLog('[VERBOSE] - Successfully loaded from Office storage:', officeSettings);
                    this.settings = { ...this.defaultSettings, ...officeSettings };
                    if (window.debugLog) window.debugLog('[VERBOSE] - Merged settings with defaults:', this.settings);
                    
                    // Migrate legacy settings if needed
                    const migrated = this.migrateLegacySettings();
                    if (migrated) {
                        await this.saveSettings(this.settings);
                        if (window.debugLog) window.debugLog('[VERBOSE] - Legacy settings migrated and saved');
                    }
                    
                    return this.settings;
                }
                console.warn('[WARN] - No Office storage settings found, trying localStorage...');
            }

            // Fallback to localStorage
            const localSettings = this.loadFromLocalStorage();
            if (localSettings) {
                if (window.debugLog) window.debugLog('[VERBOSE] - Successfully loaded from localStorage:', localSettings);
                this.settings = { ...this.defaultSettings, ...localSettings };
                if (window.debugLog) window.debugLog('[VERBOSE] - Merged settings with defaults:', this.settings);
                
                // Migrate legacy settings if needed
                const migrated = this.migrateLegacySettings();
                if (migrated) {
                    await this.saveSettings(this.settings);
                    if (window.debugLog) window.debugLog('[VERBOSE] - Legacy settings migrated and saved');
                }
                
                return this.settings;
            }
            if (window.debugLog) window.debugLog('[VERBOSE] - No localStorage settings found, using defaults...');

            // No stored settings found, use defaults
            this.settings = { ...this.defaultSettings };
            if (window.debugLog) window.debugLog('[VERBOSE] - Using default settings:', this.settings);
            await this.saveSettings(this.settings);
            if (window.debugLog) window.debugLog('[VERBOSE] - Default settings saved successfully');
            
            return this.settings;

        } catch (error) {
            console.error('[ERROR] - Failed to load settings:', error);
            if (window.debugLog) window.debugLog('[VERBOSE] - Falling back to default settings due to error');
            this.settings = { ...this.defaultSettings };
            if (window.debugLog) window.debugLog('[VERBOSE] - Using default settings after error:', this.settings);
            return this.settings;
        }
    }

    /**
     * Saves settings to storage
     * @param {Object} newSettings - Settings to save
     * @returns {Promise<boolean>} Success status
     */
    async saveSettings(newSettings = null) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Starting settings save process...');
        if (window.debugLog) window.debugLog('[VERBOSE] - Settings to save:', newSettings || this.settings);
        try {
            const settingsToSave = newSettings || this.settings;
            
            // Update timestamp
            settingsToSave['last-updated'] = new Date().toISOString();
            if (window.debugLog) window.debugLog('[VERBOSE] - Added timestamp:', settingsToSave['last-updated']);
            
            // Update internal settings
            this.settings = { ...settingsToSave };
            if (window.debugLog) window.debugLog('[VERBOSE] - Updated internal settings cache');

            // Save to Office.js RoamingSettings (unless debug flag is set)
            let officeSaved = true; // Default to success if we're skipping Office storage
            
            if (this.forceLocalStorageOnly) {
                console.warn('[DEBUG] - SKIPPING Office storage save - localStorage-only mode active');
            } else {
                if (window.debugLog) window.debugLog('[VERBOSE] - Attempting to save to Office storage...');
                officeSaved = await this.saveToOfficeStorage(settingsToSave);
                window.debugLog(`[VERBOSE] - Office storage save result: ${officeSaved ? 'SUCCESS' : 'FAILED'}`);
            }
            // Also save to localStorage as backup
            if (window.debugLog) window.debugLog('[VERBOSE] - Saving to localStorage as backup...');
            this.saveToLocalStorage(settingsToSave);
            if (window.debugLog) window.debugLog('[VERBOSE] - localStorage backup save completed');

            // Notify listeners
            if (window.debugLog) window.debugLog('[VERBOSE] - Notifying change listeners...');
            this.notifyChangeListeners(settingsToSave);
            window.debugLog(`[VERBOSE] - Notified ${this.changeListeners.length} listeners`);

            return officeSaved;

        } catch (error) {
            console.error('[ERROR] - Failed to save settings:', error);
            console.info('Save operation failed, returning false');
            return false;
        }
    }

    /**
     * Loads settings from Office.js RoamingSettings
     * @returns {Promise<Object|null>} Settings object or null
     */
    async loadFromOfficeStorage() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Loading from Office.js RoamingSettings...');
        return new Promise((resolve) => {
            try {
                if (typeof Office === 'undefined' || !Office.context?.roamingSettings) {
                    console.warn('[WARN] - Office.js or RoamingSettings not available');
                    resolve(null);
                    return;
                }

                const roamingSettings = Office.context.roamingSettings;
                if (window.debugLog) window.debugLog('[VERBOSE] - RoamingSettings object available');
                const settingsJson = roamingSettings.get(this.storageKey);
                if (window.debugLog) window.debugLog('[VERBOSE] - Raw settings from Office storage:', settingsJson);
                
                if (settingsJson) {
                    const settings = JSON.parse(settingsJson);
                    if (window.debugLog) window.debugLog('[VERBOSE] - Parsed Office settings:', settings);
                    resolve(settings);
                } else {
                    if (window.debugLog) window.debugLog('[VERBOSE] - No settings found in Office storage');
                    resolve(null);
                }

            } catch (error) {
                console.error('[ERROR] - Failed to load from Office storage:', error);
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
        if (window.debugLog) window.debugLog('[VERBOSE] - Saving to Office.js RoamingSettings...');
        if (window.debugLog) window.debugLog('[VERBOSE] - Settings to serialize:', settings);
        return new Promise((resolve) => {
            try {
                if (typeof Office === 'undefined' || !Office.context?.roamingSettings) {
                    console.warn('[WARN] - Office.js or RoamingSettings not available for save');
                    resolve(false);
                    return;
                }

                const roamingSettings = Office.context.roamingSettings;
                if (window.debugLog) window.debugLog('[VERBOSE] - RoamingSettings object available for save');
                const settingsJson = JSON.stringify(settings);
                if (window.debugLog) window.debugLog('[VERBOSE] - Serialized settings JSON:', `${settingsJson.length} characters`);
                
                roamingSettings.set(this.storageKey, settingsJson);
                if (window.debugLog) window.debugLog('[VERBOSE] - Settings data set in RoamingSettings');
                
                // Save settings asynchronously
                if (window.debugLog) window.debugLog('[VERBOSE] - Initiating async save to Office...');
                roamingSettings.saveAsync((result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        if (window.debugLog) window.debugLog('[VERBOSE] - Office storage save succeeded');
                        resolve(true);
                    } else {
                        console.error('[ERROR] - Failed to save to Office storage:', result.error);
                        resolve(false);
                    }
                });

            } catch (error) {
                console.error('[ERROR] - Error saving to Office storage:', error);
                resolve(false);
            }
        });
    }

    /**
     * Loads settings from localStorage
     * @returns {Object|null} Settings object or null
     */
    loadFromLocalStorage() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Loading from localStorage...');
        try {
            if (typeof localStorage === 'undefined') {
                console.warn('[WARN] - localStorage not available');
                return null;
            }

            const settingsJson = localStorage.getItem(this.storageKey);
            if (window.debugLog) window.debugLog('[VERBOSE] - Raw localStorage data:', settingsJson);
            if (settingsJson) {
                const settings = JSON.parse(settingsJson);
                if (window.debugLog) window.debugLog('[VERBOSE] - Parsed localStorage settings:', settings);
                return settings;
            }
            
            if (window.debugLog) window.debugLog('[VERBOSE] - No settings found in localStorage');
            return null;

        } catch (error) {
            console.error('[ERROR] - Failed to load from localStorage:', error);
            return null;
        }
    }

    /**
     * Saves settings to localStorage
     * @param {Object} settings - Settings to save
     */
    saveToLocalStorage(settings) {
        if (window.debugLog) window.debugLog('[VERBOSE] - Saving to localStorage...');
        if (window.debugLog) window.debugLog('[VERBOSE] - Settings to save to localStorage:', settings);
        try {
            if (typeof localStorage === 'undefined') {
                console.warn('[WARN] - localStorage not available for save');
                return;
            }

            const settingsJson = JSON.stringify(settings);
            if (window.debugLog) window.debugLog('[VERBOSE] - Serialized localStorage JSON:', `${settingsJson.length} characters`);
            localStorage.setItem(this.storageKey, settingsJson);
            if (window.debugLog) window.debugLog('[VERBOSE] - localStorage save completed');

        } catch (error) {
            console.error('[ERROR] - Failed to save to localStorage:', error);
        }
    }

    /**
     * Gets a specific setting value
     * @param {string} key - Setting key
     * @param {*} defaultValue - Default value if not found
     * @returns {*} Setting value
     */
    getSetting(key, defaultValue = null) {
        const value = this.settings[key] !== undefined ? this.settings[key] : defaultValue;
        console.info(`Getting setting '${key}':`, value);
        return value;
    }

    /**
     * Sets a specific setting value
     * @param {string} key - Setting key
     * @param {*} value - Setting value
     * @returns {Promise<boolean>} Success status
     */
    async setSetting(key, value) {
        console.info(`Setting '${key}' to:`, value);
        this.settings[key] = value;
        const result = await this.saveSettings();
        console.info(`Save result for '${key}':`, result);
        return result;
    }

    /**
     * Gets all current settings
     * @returns {Object} All settings
     */
    getSettings() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Getting all settings (instance:', this.instanceId, '):', this.settings);
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
        
        // Remove sensitive fields
        delete exportData['api-key'];
        delete exportData['endpoint-url'];
        
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
            
            // Merge with current settings (preserve sensitive data)
            const mergedSettings = {
                ...this.settings,
                ...validatedSettings,
                // Keep current API key and endpoint
                'api-key': this.settings['api-key'],
                'endpoint-url': this.settings['endpoint-url']
            };

            return await this.saveSettings(mergedSettings);

        } catch (error) {
            console.error('[ERROR] - Failed to import settings:', error);
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
                    console.warn(`[WARN] - Invalid type for setting '${key}', using default`);
                    validated[key] = defaultValue;
                }
            } else {
                validated[key] = this.defaultSettings[key];
            }
        });

        // Validate specific setting ranges
        if (validated['response-length']) {
            const length = parseInt(validated['response-length']);
            validated['response-length'] = (length >= 1 && length <= 5) ? length.toString() : '3';
        }

        if (validated['response-tone']) {
            const tone = parseInt(validated['response-tone']);
            validated['response-tone'] = (tone >= 1 && tone <= 5) ? tone.toString() : '3';
        }

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
                console.error('[ERROR] - Settings change listener error:', error);
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
            console.error('[ERROR] - Settings migration failed:', error);
            return false;
        }
    }

    /**
     * Clears all stored settings
     * @returns {Promise<boolean>} Success status
     */
    async clearAllSettings() {
        if (window.debugLog) window.debugLog('[VERBOSE] - Starting complete settings reset...');
        try {
            let officeCleared = true;
            let localStorageCleared = true;

            // Clear from Office.js RoamingSettings
            if (typeof Office !== 'undefined' && Office.context?.roamingSettings) {
                if (window.debugLog) window.debugLog('[VERBOSE] - Clearing Office.js RoamingSettings...');
                const roamingSettings = Office.context.roamingSettings;
                roamingSettings.remove(this.storageKey);
                
                officeCleared = await new Promise((resolve) => {
                    roamingSettings.saveAsync((result) => {
                        const success = result.status === Office.AsyncResultStatus.Succeeded;
                        if (window.debugLog) window.debugLog(`[VERBOSE] - Office.js clear result: ${success ? 'SUCCESS' : 'FAILED'}`);
                        if (!success && result.error) {
                            console.error('[ERROR] - Office.js clear error:', result.error);
                        }
                        resolve(success);
                    });
                });
            } else {
                if (window.debugLog) window.debugLog('[VERBOSE] - Office.js RoamingSettings not available, skipping...');
            }

            // Clear from localStorage
            if (typeof localStorage !== 'undefined') {
                if (window.debugLog) window.debugLog('[VERBOSE] - Clearing localStorage...');
                try {
                    localStorage.removeItem(this.storageKey);
                    // Verify removal
                    const stillExists = localStorage.getItem(this.storageKey);
                    localStorageCleared = stillExists === null;
                    if (window.debugLog) window.debugLog(`[VERBOSE] - localStorage clear result: ${localStorageCleared ? 'SUCCESS' : 'FAILED'}`);
                } catch (localError) {
                    console.error('[ERROR] - localStorage clear error:', localError);
                    localStorageCleared = false;
                }
            } else {
                if (window.debugLog) window.debugLog('[VERBOSE] - localStorage not available, skipping...');
            }

            // Reset to defaults
            if (window.debugLog) window.debugLog('[VERBOSE] - Resetting internal settings to defaults...');
            this.settings = { ...this.defaultSettings };
            
            // Ensure provider configs are completely reset to avoid contamination
            this.settings['provider-configs'] = {
                'openai': { 'api-key': '', 'endpoint-url': '' },    // Empty = use ai-providers.json defaults
                'ollama': { 'api-key': '', 'endpoint-url': '' },    // Empty = use ai-providers.json defaults
                'onsite1': { 'api-key': '', 'endpoint-url': '' },   // Empty = use ai-providers.json defaults
                'onsite2': { 'api-key': '', 'endpoint-url': '' }    // Empty = use ai-providers.json defaults
            };
            
            // Notify listeners of the reset
            if (window.debugLog) window.debugLog('[VERBOSE] - Notifying change listeners of reset...');
            this.notifyChangeListeners(this.settings);

            const overallSuccess = officeCleared && localStorageCleared;
            if (window.debugLog) window.debugLog(`[VERBOSE] - Complete settings reset result: ${overallSuccess ? 'SUCCESS' : 'PARTIAL'} (Office: ${officeCleared}, localStorage: ${localStorageCleared})`);
            
            return overallSuccess;

        } catch (error) {
            console.error('[ERROR] - Failed to clear settings:', error);
            if (window.debugLog) window.debugLog('[VERBOSE] - Settings reset failed with exception');
            return false;
        }
    }

    /**
     * Get provider-specific configuration for a given service
     * @param {string} provider - The provider key (e.g., 'openai', 'ollama')
     * @returns {Object} Provider configuration with api-key and endpoint-url
     */
    getProviderConfig(provider) {
        if (!provider) return { 'api-key': '', 'endpoint-url': '' };
        
        // Try provider-specific config first
        if (this.settings['provider-configs'] && this.settings['provider-configs'][provider]) {
            return { ...this.settings['provider-configs'][provider] };
        }
        
        // Fall back to legacy single settings for backwards compatibility
        return {
            'api-key': this.settings['api-key'] || '',
            'endpoint-url': this.settings['endpoint-url'] || ''
        };
    }

    /**
     * Set provider-specific configuration for a given service
     * @param {string} provider - The provider key
     * @param {string} apiKey - The API key for this provider
     * @param {string} endpointUrl - The endpoint URL for this provider
     * @returns {Promise<boolean>} Success status
     */
    async setProviderConfig(provider, apiKey, endpointUrl) {
        if (!provider) return false;
        
        try {
            // Initialize provider-configs if it doesn't exist
            if (!this.settings['provider-configs']) {
                this.settings['provider-configs'] = { ...this.defaultSettings['provider-configs'] };
            }
            
            // Initialize this provider's config if it doesn't exist
            if (!this.settings['provider-configs'][provider]) {
                this.settings['provider-configs'][provider] = { 'api-key': '', 'endpoint-url': '' };
            }
            
            // Update the provider's configuration
            this.settings['provider-configs'][provider]['api-key'] = apiKey || '';
            this.settings['provider-configs'][provider]['endpoint-url'] = endpointUrl || '';
            
            // Save the updated settings
            return await this.saveSettings(this.settings);
            
        } catch (error) {
            console.error('[ERROR] - Failed to set provider config:', error);
            return false;
        }
    }

    /**
     * Migrate legacy single API key/endpoint settings to provider-specific format
     * This ensures backwards compatibility when upgrading
     */
    migrateLegacySettings() {
        try {
            // Only migrate if we have legacy settings but no provider-configs
            if ((this.settings['api-key'] || this.settings['endpoint-url']) && 
                !this.settings['provider-configs']) {
                
                if (window.debugLog) window.debugLog('[VERBOSE] - Migrating legacy settings to provider-specific format');
                
                // Initialize provider configs
                this.settings['provider-configs'] = { ...this.defaultSettings['provider-configs'] };
                
                // Get the current model service or default to openai
                const currentService = this.settings['model-service'] || 'openai';
                
                // Migrate the legacy settings to the current provider
                if (this.settings['provider-configs'][currentService]) {
                    this.settings['provider-configs'][currentService]['api-key'] = this.settings['api-key'] || '';
                    this.settings['provider-configs'][currentService]['endpoint-url'] = this.settings['endpoint-url'] || '';
                }
                
                if (window.debugLog) window.debugLog('[VERBOSE] - Legacy settings migrated successfully');
                return true;
            }
            
            return false; // No migration needed
            
        } catch (error) {
            console.error('[ERROR] - Failed to migrate legacy settings:', error);
            return false;
        }
    }

    /**
     * Add a new writing sample
     * @param {string} title - Title/description for the sample
     * @param {string} content - The writing sample content
     * @returns {Promise<Object>} The added sample with ID
     */
    async addWritingSample(title, content) {
        const samples = this.settings['writing-samples'] || [];
        const newSample = {
            id: Date.now(), // Simple ID generation
            title: title.trim(),
            content: content.trim(),
            dateAdded: new Date().toISOString(),
            wordCount: content.trim().split(/\s+/).length
        };
        
        samples.push(newSample);
        await this.setSetting('writing-samples', samples);
        
        if (window.debugLog) {
            window.debugLog(`[VERBOSE] - Added writing sample: ${title} (${newSample.wordCount} words)`);
        }
        
        return newSample;
    }

    /**
     * Update an existing writing sample
     * @param {number} id - Sample ID to update
     * @param {string} title - Updated title
     * @param {string} content - Updated content
     * @returns {Promise<boolean>} Success status
     */
    async updateWritingSample(id, title, content) {
        const samples = this.settings['writing-samples'] || [];
        const index = samples.findIndex(sample => sample.id === id);
        
        if (index === -1) {
            console.warn(`[WARN] - Writing sample with ID ${id} not found`);
            return false;
        }
        
        samples[index] = {
            ...samples[index],
            title: title.trim(),
            content: content.trim(),
            wordCount: content.trim().split(/\s+/).length,
            lastModified: new Date().toISOString()
        };
        
        await this.setSetting('writing-samples', samples);
        
        if (window.debugLog) {
            window.debugLog(`[VERBOSE] - Updated writing sample: ${title} (${samples[index].wordCount} words)`);
        }
        
        return true;
    }

    /**
     * Delete a writing sample
     * @param {number} id - Sample ID to delete
     * @returns {Promise<boolean>} Success status
     */
    async deleteWritingSample(id) {
        const samples = this.settings['writing-samples'] || [];
        const initialLength = samples.length;
        const filteredSamples = samples.filter(sample => sample.id !== id);
        
        if (filteredSamples.length === initialLength) {
            console.warn(`[WARN] - Writing sample with ID ${id} not found`);
            return false;
        }
        
        await this.setSetting('writing-samples', filteredSamples);
        
        if (window.debugLog) {
            window.debugLog(`[VERBOSE] - Deleted writing sample with ID: ${id}`);
        }
        
        return true;
    }

    /**
     * Get all writing samples
     * @returns {Array} Array of writing samples
     */
    getWritingSamples() {
        const samples = this.settings['writing-samples'] || [];
        if (window.debugLog) {
            window.debugLog('[VERBOSE] - SettingsManager: getWritingSamples returning', samples.length, 'samples');
            if (samples.length > 0) {
                window.debugLog('[VERBOSE] - SettingsManager: Sample details:', samples.map(s => ({ id: s.id, title: s.title, wordCount: s.wordCount })));
            }
        }
        return samples;
    }

    /**
     * Get writing style settings
     * @returns {Object} Style settings
     */
    getStyleSettings() {
        return {
            enabled: this.settings['style-analysis-enabled'] !== undefined ? this.settings['style-analysis-enabled'] : false,
            strength: this.settings['style-strength'] || 'medium',
            samplesCount: (this.settings['writing-samples'] || []).length
        };
    }

    /**
     * Debug Methods for Storage Testing
     */
    
    /**
     * Enable localStorage-only mode for testing
     */
    enableLocalStorageOnlyMode() {
        this.forceLocalStorageOnly = true;
        localStorage.setItem('debug-force-local-storage', 'true');
        console.warn('[DEBUG] - Enabled localStorage-only mode. Office.js storage disabled.');
        console.info('[DEBUG] - To disable: settingsManager.disableLocalStorageOnlyMode()');
    }

    /**
     * Disable localStorage-only mode and return to normal dual storage
     */
    disableLocalStorageOnlyMode() {
        this.forceLocalStorageOnly = false;
        localStorage.removeItem('debug-force-local-storage');
        console.info('[DEBUG] - Disabled localStorage-only mode. Office.js storage re-enabled.');
    }

    /**
     * Get current storage mode for debugging
     */
    getStorageMode() {
        return {
            mode: this.forceLocalStorageOnly ? 'localStorage-only' : 'dual-storage',
            officeAvailable: typeof Office !== 'undefined' && Office.context?.roamingSettings,
            localStorageAvailable: typeof localStorage !== 'undefined'
        };
    }
}

