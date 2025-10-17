/**
 * UI State Manager
 * Manages transient UI state separately from persistent storage
 * Prevents UI values from contaminating saved settings
 */

export class UIStateManager {
    constructor(settingsManager) {
        this.settingsManager = settingsManager;
        
        // Transient UI state (never persisted)
        this.uiState = {
            // Current UI form values (may differ from saved settings)
            currentProvider: '',
            formValues: {
                'api-key': '',
                'endpoint-url': '',
                'selected-model': ''
            },
            
            // UI interaction state
            isSettingsPanelOpen: false,
            isDirty: false, // Has unsaved changes
            lastFormUpdate: null
        };
        
        // Track which UI elements are bound
        this.boundElements = new Set();
    }

    /**
     * Initialize UI state from saved settings (one-way: storage → UI)
     */
    async initializeFromStorage() {
        const activeProvider = this.settingsManager.getActiveProvider();
        const providerConfig = this.settingsManager.getProviderConfig(activeProvider);
        
        this.uiState.currentProvider = activeProvider;
        this.uiState.formValues = {
            'api-key': providerConfig['api-key'] || '',
            'endpoint-url': providerConfig['endpoint-url'] || '',
            'selected-model': providerConfig['selected-model'] || ''
        };
        this.uiState.isDirty = false;
        
        console.log('UI state initialized from storage:', activeProvider);
    }

    /**
     * Get current UI form values (transient)
     * @returns {Object} Current form values
     */
    getUIFormValues() {
        return { ...this.uiState.formValues };
    }

    /**
     * Update UI form value (transient, not saved)
     * @param {string} key - Form field key
     * @param {string} value - Form field value
     */
    updateUIFormValue(key, value) {
        if (this.uiState.formValues[key] !== value) {
            this.uiState.formValues[key] = value;
            this.uiState.isDirty = true;
            this.uiState.lastFormUpdate = new Date().toISOString();
            
            // Update UI element if bound
            this.updateUIElement(key, value);
        }
    }

    /**
     * Set current provider in UI (transient)
     * @param {string} provider - Provider key
     */
    setCurrentUIProvider(provider) {
        if (this.uiState.currentProvider !== provider) {
            // Save current form values before switching
            if (this.uiState.isDirty) {
                console.log('UI provider switch with unsaved changes, clearing form');
            }
            
            // Switch to new provider and load its config
            this.uiState.currentProvider = provider;
            const providerConfig = this.settingsManager.getProviderConfig(provider);
            
            this.uiState.formValues = {
                'api-key': providerConfig['api-key'] || '',
                'endpoint-url': providerConfig['endpoint-url'] || '',
                'selected-model': providerConfig['selected-model'] || ''
            };
            this.uiState.isDirty = false;
            
            // Update UI elements
            this.updateAllUIElements();
            
            console.log('UI switched to provider:', provider);
        }
    }

    /**
     * Get current UI provider
     * @returns {string} Current UI provider
     */
    getCurrentUIProvider() {
        return this.uiState.currentProvider;
    }

    /**
     * Check if UI has unsaved changes
     * @returns {boolean} Whether UI is dirty
     */
    hasUnsavedChanges() {
        return this.uiState.isDirty;
    }

    /**
     * Save current UI state to persistent storage
     * @returns {Promise<boolean>} Success status
     */
    async saveUIStateToStorage() {
        if (!this.uiState.isDirty) {
            console.log('No UI changes to save');
            return true;
        }

        try {
            const provider = this.uiState.currentProvider;
            const formValues = this.uiState.formValues;
            
            // Save provider configuration
            const success = await this.settingsManager.setProviderConfig(
                provider,
                formValues['api-key'],
                formValues['endpoint-url'],
                formValues['selected-model']
            );
            
            if (success) {
                this.uiState.isDirty = false;
                console.log('UI state saved to storage:', provider);
            }
            
            return success;
            
        } catch (error) {
            console.error('Failed to save UI state:', error);
            return false;
        }
    }

    /**
     * Discard unsaved UI changes and reload from storage
     */
    async discardUIChanges() {
        if (this.uiState.isDirty) {
            console.log('Discarding unsaved UI changes');
            await this.initializeFromStorage();
            this.updateAllUIElements();
        }
    }

    /**
     * Bind UI element to state management
     * @param {string} elementId - DOM element ID
     * @param {string} stateKey - State key to bind to
     */
    bindUIElement(elementId, stateKey) {
        const element = document.getElementById(elementId);
        if (!element) {
            console.warn('Cannot bind missing element:', elementId);
            return;
        }

        this.boundElements.add({ elementId, stateKey });

        // Add event listener for changes
        element.addEventListener('input', (event) => {
            this.updateUIFormValue(stateKey, event.target.value);
        });

        element.addEventListener('change', (event) => {
            this.updateUIFormValue(stateKey, event.target.value);
        });

        console.log('Bound UI element:', elementId, '→', stateKey);
    }

    /**
     * Update specific UI element from state
     * @param {string} stateKey - State key
     * @param {string} value - Value to set
     */
    updateUIElement(stateKey, value) {
        // Map state keys to element IDs
        const elementMappings = {
            'api-key': 'api-key',
            'endpoint-url': 'endpoint-url',
            'selected-model': 'selected-model'
        };

        const elementId = elementMappings[stateKey];
        if (elementId) {
            const element = document.getElementById(elementId);
            if (element) {
                element.value = value || '';
            }
        }
    }

    /**
     * Update all bound UI elements from current state
     */
    updateAllUIElements() {
        Object.keys(this.uiState.formValues).forEach(key => {
            this.updateUIElement(key, this.uiState.formValues[key]);
        });
    }

    /**
     * Set settings panel state
     * @param {boolean} isOpen - Whether settings panel is open
     */
    setSettingsPanelState(isOpen) {
        this.uiState.isSettingsPanelOpen = isOpen;
    }

    /**
     * Check if settings panel is open
     * @returns {boolean} Whether settings panel is open
     */
    isSettingsPanelOpen() {
        return this.uiState.isSettingsPanelOpen;
    }

    /**
     * Get UI state for debugging
     * @returns {Object} Current UI state
     */
    getDebugState() {
        return {
            currentProvider: this.uiState.currentProvider,
            formValues: { ...this.uiState.formValues },
            isDirty: this.uiState.isDirty,
            isSettingsPanelOpen: this.uiState.isSettingsPanelOpen,
            boundElements: Array.from(this.boundElements),
            lastFormUpdate: this.uiState.lastFormUpdate
        };
    }
}