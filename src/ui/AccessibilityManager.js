/**
 * Accessibility Manager
 * Handles accessibility features and ARIA support
 */

export class AccessibilityManager {
    constructor() {
        this.isScreenReaderMode = false;
        this.isHighContrastMode = false;
        this.announcements = [];
        this.announcementDelay = 100; // ms delay between announcements
    }

    /**
     * Initializes accessibility features
     */
    initialize() {
        this.setupAriaLive();
        this.setupKeyboardNavigation();
        this.setupFocusManagement();
        this.detectAccessibilityPreferences();
    }

    /**
     * Sets up ARIA live regions for announcements
     */
    setupAriaLive() {
        // Create or ensure screen reader announcement region exists
        let announceRegion = document.getElementById('sr-announcements');
        if (!announceRegion) {
            announceRegion = document.createElement('div');
            announceRegion.id = 'sr-announcements';
            announceRegion.className = 'sr-only';
            announceRegion.setAttribute('aria-live', 'assertive');
            announceRegion.setAttribute('aria-atomic', 'true');
            document.body.appendChild(announceRegion);
        }

        // Create polite announcement region
        let politeRegion = document.getElementById('sr-announcements-polite');
        if (!politeRegion) {
            politeRegion = document.createElement('div');
            politeRegion.id = 'sr-announcements-polite';
            politeRegion.className = 'sr-only';
            politeRegion.setAttribute('aria-live', 'polite');
            politeRegion.setAttribute('aria-atomic', 'true');
            document.body.appendChild(politeRegion);
        }
    }

    /**
     * Sets up keyboard navigation
     */
    setupKeyboardNavigation() {
        // Global keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            this.handleGlobalKeyboard(e);
        });

        // Focus trap for modal dialogs
        this.setupFocusTrapping();

        // Skip links for screen readers
        this.addSkipLinks();
    }

    /**
     * Handles global keyboard shortcuts
     * @param {KeyboardEvent} event - Keyboard event
     */
    handleGlobalKeyboard(event) {
        // Alt + A: Focus on analyze button
        if (event.altKey && event.key === 'a') {
            event.preventDefault();
            const analyzeBtn = document.getElementById('analyze-email');
            if (analyzeBtn && !analyzeBtn.disabled) {
                analyzeBtn.focus();
                this.announce('Analyze email button focused');
            }
        }

        // Alt + R: Focus on generate response button
        if (event.altKey && event.key === 'r') {
            event.preventDefault();
            const responseBtn = document.getElementById('generate-response');
            if (responseBtn && !responseBtn.disabled) {
                responseBtn.focus();
                this.announce('Generate response button focused');
            }
        }

        // Alt + S: Open settings
        if (event.altKey && event.key === 's') {
            event.preventDefault();
            const settingsBtn = document.getElementById('open-settings');
            if (settingsBtn) {
                settingsBtn.click();
                this.announce('Settings panel opened');
            }
        }

        // Escape: Close modals/panels
        if (event.key === 'Escape') {
            this.handleEscape();
        }

        // Tab navigation enhancements
        if (event.key === 'Tab') {
            this.enhanceTabNavigation(event);
        }
    }

    /**
     * Handles escape key to close panels
     */
    handleEscape() {
        // Close settings panel
        const settingsPanel = document.getElementById('settings-panel');
        if (settingsPanel && !settingsPanel.classList.contains('hidden')) {
            document.getElementById('close-settings').click();
            this.announce('Settings panel closed');
            return;
        }
    }

    /**
     * Enhances tab navigation
     * @param {KeyboardEvent} event - Tab key event
     */
    enhanceTabNavigation(event) {
        // Ensure tab order respects hidden elements
        const focusableElements = this.getFocusableElements();
        
        if (focusableElements.length === 0) return;

        const currentIndex = focusableElements.indexOf(document.activeElement);
        
        if (event.shiftKey) {
            // Shift+Tab - backward
            if (currentIndex <= 0) {
                event.preventDefault();
                focusableElements[focusableElements.length - 1].focus();
            }
        } else {
            // Tab - forward
            if (currentIndex >= focusableElements.length - 1) {
                event.preventDefault();
                focusableElements[0].focus();
            }
        }
    }

    /**
     * Gets all focusable elements that are currently visible
     * @returns {Array} Array of focusable elements
     */
    getFocusableElements() {
        const selector = 'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])';
        const elements = Array.from(document.querySelectorAll(selector));
        
        return elements.filter(element => {
            // Check if element is visible
            const style = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            
            return style.display !== 'none' && 
                   style.visibility !== 'hidden' && 
                   rect.width > 0 && 
                   rect.height > 0 &&
                   !element.closest('.hidden');
        });
    }

    /**
     * Sets up focus trapping for modal dialogs
     */
    setupFocusTrapping() {
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Tab') {
                const settingsPanel = document.getElementById('settings-panel');
                if (settingsPanel && !settingsPanel.classList.contains('hidden')) {
                    this.trapFocusInElement(e, settingsPanel);
                }
            }
        });
    }

    /**
     * Traps focus within a specific element
     * @param {KeyboardEvent} event - Tab event
     * @param {Element} container - Container element
     */
    trapFocusInElement(event, container) {
        const focusableElements = container.querySelectorAll(
            'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
        );
        
        if (focusableElements.length === 0) return;

        const firstElement = focusableElements[0];
        const lastElement = focusableElements[focusableElements.length - 1];

        if (event.shiftKey) {
            // Shift+Tab
            if (document.activeElement === firstElement) {
                event.preventDefault();
                lastElement.focus();
            }
        } else {
            // Tab
            if (document.activeElement === lastElement) {
                event.preventDefault();
                firstElement.focus();
            }
        }
    }

    /**
     * Adds skip links for screen readers
     */
    addSkipLinks() {
        const skipNav = document.createElement('div');
        skipNav.className = 'skip-links';
        skipNav.innerHTML = `
            <a href="#main-content" class="skip-link">Skip to main content</a>
            <a href="#analyze-email" class="skip-link">Skip to analyze button</a>
            <a href="#response-section" class="skip-link">Skip to results</a>
        `;
        
        document.body.insertBefore(skipNav, document.body.firstChild);
    }

    /**
     * Sets up focus management
     */
    setupFocusManagement() {
        // Focus management for dynamic content
        this.setupDynamicFocusManagement();
        
        // Restore focus after actions
        this.setupFocusRestoration();
    }

    /**
     * Sets up focus management for dynamic content
     */
    setupDynamicFocusManagement() {
        // Focus first interactive element when panels open
        const observer = new MutationObserver((mutations) => {
            mutations.forEach((mutation) => {
                if (mutation.type === 'attributes' && mutation.attributeName === 'class') {
                    const target = mutation.target;
                    
                    // Settings panel opened
                    if (target.id === 'settings-panel' && !target.classList.contains('hidden')) {
                        setTimeout(() => {
                            const firstInput = target.querySelector('input, button, select');
                            if (firstInput) {
                                firstInput.focus();
                                this.announce('Settings panel opened');
                            }
                        }, 100);
                    }
                    
                    // Response section shown
                    if (target.id === 'response-section' && !target.classList.contains('hidden')) {
                        setTimeout(() => {
                            const tabButton = document.getElementById('tab-analysis');
                            if (tabButton) {
                                tabButton.focus();
                                this.announce('Analysis results are ready');
                            }
                        }, 100);
                    }
                }
            });
        });

        observer.observe(document.body, {
            attributes: true,
            subtree: true,
            attributeFilter: ['class']
        });
    }

    /**
     * Sets up focus restoration after actions
     */
    setupFocusRestoration() {
        // Store focus before async operations
        this.focusBeforeOperation = null;

        // Button click handlers to store focus
        document.addEventListener('click', (e) => {
            if (e.target.matches('button')) {
                this.focusBeforeOperation = e.target;
            }
        });
    }

    /**
     * Announces text to screen readers
     * @param {string} message - Message to announce
     * @param {boolean} assertive - Whether to use assertive (true) or polite (false) announcement
     */
    announce(message, assertive = true) {
        if (!message) return;

        const regionId = assertive ? 'sr-announcements' : 'sr-announcements-polite';
        const region = document.getElementById(regionId);
        
        if (region) {
            // Add to announcement queue
            this.announcements.push({ message, assertive, timestamp: Date.now() });
            
            // Process queue
            this.processAnnouncementQueue();
        }
    }

    /**
     * Processes the announcement queue
     */
    processAnnouncementQueue() {
        if (this.announcements.length === 0) return;

        const announcement = this.announcements.shift();
        const regionId = announcement.assertive ? 'sr-announcements' : 'sr-announcements-polite';
        const region = document.getElementById(regionId);

        if (region) {
            // Clear previous announcement
            region.textContent = '';
            
            // Add new announcement after a brief delay
            setTimeout(() => {
                region.textContent = announcement.message;
                
                // Process next announcement
                if (this.announcements.length > 0) {
                    setTimeout(() => {
                        this.processAnnouncementQueue();
                    }, this.announcementDelay);
                }
            }, 50);
        }
    }

    /**
     * Detects user accessibility preferences
     */
    detectAccessibilityPreferences() {
        // Detect high contrast preference
        if (window.matchMedia) {
            const highContrastQuery = window.matchMedia('(prefers-contrast: high)');
            this.isHighContrastMode = highContrastQuery.matches;
            
            highContrastQuery.addEventListener('change', (e) => {
                this.isHighContrastMode = e.matches;
                this.updateAccessibilityStyles();
            });

            // Detect reduced motion preference
            const reducedMotionQuery = window.matchMedia('(prefers-reduced-motion: reduce)');
            if (reducedMotionQuery.matches) {
                document.body.classList.add('reduced-motion');
            }

            reducedMotionQuery.addEventListener('change', (e) => {
                document.body.classList.toggle('reduced-motion', e.matches);
            });
        }
    }

    /**
     * Sets screen reader mode
     * @param {boolean} enabled - Whether screen reader mode is enabled
     */
    setScreenReaderMode(enabled) {
        this.isScreenReaderMode = enabled;
        document.body.classList.toggle('screen-reader-mode', enabled);
        
        if (enabled) {
            this.enhanceForScreenReaders();
            this.announce('Screen reader mode enabled');
        } else {
            this.announce('Screen reader mode disabled');
        }
    }

    /**
     * Enhances interface for screen readers
     */
    enhanceForScreenReaders() {
        // Add more descriptive labels
        this.addDescriptiveLabels();
        
        // Add status descriptions
        this.addStatusDescriptions();
        
        // Enhance form controls
        this.enhanceFormControls();
    }

    /**
     * Adds descriptive labels for screen readers
     */
    addDescriptiveLabels() {
        // Enhance slider labels
        const sliders = document.querySelectorAll('.slider');
        sliders.forEach(slider => {
            const valueDisplay = document.getElementById(slider.id.replace('response-', '') + '-value');
            if (valueDisplay) {
                slider.setAttribute('aria-describedby', valueDisplay.id);
            }
        });

        // Add descriptions to buttons
        const analyzeBtn = document.getElementById('analyze-email');
        if (analyzeBtn) {
            analyzeBtn.setAttribute('aria-describedby', 'analyze-description');
            
            let description = document.getElementById('analyze-description');
            if (!description) {
                description = document.createElement('div');
                description.id = 'analyze-description';
                description.className = 'sr-only';
                description.textContent = 'Analyzes the current email using AI to identify key points, sentiment, and required actions';
                analyzeBtn.parentNode.appendChild(description);
            }
        }
    }

    /**
     * Adds status descriptions
     */
    addStatusDescriptions() {
        // Add loading status descriptions
        const loadingElement = document.getElementById('loading');
        if (loadingElement) {
            loadingElement.setAttribute('role', 'status');
            loadingElement.setAttribute('aria-label', 'Loading email analysis');
        }
    }

    /**
     * Enhances form controls for accessibility
     */
    enhanceFormControls() {
        // Add required indicators
        const requiredFields = ['api-key'];
        requiredFields.forEach(fieldId => {
            const field = document.getElementById(fieldId);
            if (field) {
                field.setAttribute('aria-required', 'true');
                
                const label = document.querySelector(`label[for="${fieldId}"]`);
                if (label && !label.querySelector('.required-indicator')) {
                    const indicator = document.createElement('span');
                    indicator.className = 'required-indicator';
                    indicator.textContent = ' (required)';
                    indicator.setAttribute('aria-label', 'required field');
                    label.appendChild(indicator);
                }
            }
        });

        // Add error descriptions
        this.setupErrorDescriptions();
    }

    /**
     * Sets up error descriptions for form validation
     */
    setupErrorDescriptions() {
        const inputs = document.querySelectorAll('input, select, textarea');
        inputs.forEach(input => {
            input.addEventListener('invalid', (e) => {
                this.addErrorDescription(e.target);
            });
            
            input.addEventListener('input', (e) => {
                this.removeErrorDescription(e.target);
            });
        });
    }

    /**
     * Adds error description to a form field
     * @param {Element} field - Form field element
     */
    addErrorDescription(field) {
        const errorId = field.id + '-error';
        let errorElement = document.getElementById(errorId);
        
        if (!errorElement) {
            errorElement = document.createElement('div');
            errorElement.id = errorId;
            errorElement.className = 'error-message';
            errorElement.setAttribute('role', 'alert');
            field.parentNode.appendChild(errorElement);
        }

        errorElement.textContent = field.validationMessage || 'Invalid input';
        field.setAttribute('aria-describedby', errorId);
        field.classList.add('error');
        
        this.announce(`Error in ${field.getAttribute('aria-label') || field.id}: ${errorElement.textContent}`);
    }

    /**
     * Removes error description from a form field
     * @param {Element} field - Form field element
     */
    removeErrorDescription(field) {
        const errorId = field.id + '-error';
        const errorElement = document.getElementById(errorId);
        
        if (errorElement) {
            errorElement.remove();
        }
        
        field.removeAttribute('aria-describedby');
        field.classList.remove('error');
    }

    /**
     * Updates accessibility styles based on preferences
     */
    updateAccessibilityStyles() {
        document.body.classList.toggle('high-contrast-auto', this.isHighContrastMode);
    }

    /**
     * Restores focus to previously focused element
     */
    restoreFocus() {
        if (this.focusBeforeOperation && document.contains(this.focusBeforeOperation)) {
            this.focusBeforeOperation.focus();
            this.focusBeforeOperation = null;
        }
    }

    /**
     * Gets accessibility status
     * @returns {Object} Current accessibility status
     */
    getStatus() {
        return {
            screenReaderMode: this.isScreenReaderMode,
            highContrastMode: this.isHighContrastMode,
            focusableElementsCount: this.getFocusableElements().length,
            announcementQueueLength: this.announcements.length
        };
    }
}
