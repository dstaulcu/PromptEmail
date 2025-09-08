/**
 * UI Controller
 * Manages UI state, loading states, and user feedback
 */

export class UIController {
    constructor() {
        this.loadingStates = new Map(); // Track loading states for different elements
        this.statusTimeout = null;
        this.errorTimeout = null;
    }

    /**
     * Shows the loading screen
     */
    showLoading() {
        const loading = document.getElementById('loading');
        const mainContent = document.getElementById('main-content');
        
        if (loading) loading.classList.remove('hidden');
        if (mainContent) mainContent.classList.add('hidden');
    }

    /**
     * Hides the loading screen
     */
    hideLoading() {
        const loading = document.getElementById('loading');
        if (loading) loading.classList.add('hidden');
    }

    /**
     * Shows the main content
     */
    showMainContent() {
        const mainContent = document.getElementById('main-content');
        if (mainContent) mainContent.classList.remove('hidden');
    }

    /**
     * Sets loading state for a specific button
     * @param {string} buttonId - ID of the button
     * @param {boolean} isLoading - Whether button is in loading state
     */
    setButtonLoading(buttonId, isLoading) {
        const button = document.getElementById(buttonId);
        if (!button) return;

        if (isLoading) {
            // Store original content
            this.loadingStates.set(buttonId, {
                originalText: button.innerHTML,
                originalDisabled: button.disabled
            });

            // Set loading state
            button.disabled = true;
            button.classList.add('loading');
            
            // Add spinner icon
            const spinner = '<span class="spinner-small" aria-hidden="true"></span>';
            const text = button.textContent.trim();
            button.innerHTML = `${spinner} ${text}`;
            
            // Update aria-label for screen readers
            button.setAttribute('aria-label', `${text} - loading`);
            
        } else {
            // Restore original state
            const originalState = this.loadingStates.get(buttonId);
            if (originalState) {
                button.innerHTML = originalState.originalText;
                button.disabled = originalState.originalDisabled;
                button.classList.remove('loading');
                button.removeAttribute('aria-label');
                
                this.loadingStates.delete(buttonId);
            }
        }
    }

    /**
     * Shows a status message to the user
     * @param {string} message - Status message
     * @param {string} type - Message type ('info', 'success', 'warning', 'error')
     * @param {number} timeout - Auto-hide timeout in ms (0 for no timeout)
     */
    showStatus(message, type = 'info', timeout = 5000) {
        const statusContainer = document.getElementById('status-messages');
        if (!statusContainer) return;

        // Clear previous timeout
        if (this.statusTimeout) {
            clearTimeout(this.statusTimeout);
        }

        // Create status element
        const statusElement = this.createStatusElement(message, type);
        
        // Clear previous messages and add new one
        statusContainer.innerHTML = '';
        statusContainer.appendChild(statusElement);
        
        // Show container
        statusContainer.classList.remove('hidden');

        // Auto-hide after timeout
        if (timeout > 0) {
            this.statusTimeout = setTimeout(() => {
                this.hideStatus();
            }, timeout);
        }

        // Announce to screen readers
        this.announceToScreenReader(message, type === 'error');
    }

    /**
     * Shows an error message
     * @param {string} message - Error message
     * @param {number} timeout - Auto-hide timeout
     */
    showError(message, timeout = 8000) {
        this.showStatus(message, 'error', timeout);
    }

    /**
     * Shows a success message
     * @param {string} message - Success message
     * @param {number} timeout - Auto-hide timeout
     */
    showSuccess(message, timeout = 4000) {
        this.showStatus(message, 'success', timeout);
    }

    /**
     * Shows a warning message
     * @param {string} message - Warning message
     * @param {number} timeout - Auto-hide timeout
     */
    showWarning(message, timeout = 6000) {
        this.showStatus(message, 'warning', timeout);
    }

    /**
     * Hides status messages
     */
    hideStatus() {
        const statusContainer = document.getElementById('status-messages');
        if (statusContainer) {
            statusContainer.classList.add('hidden');
            statusContainer.innerHTML = '';
        }

        if (this.statusTimeout) {
            clearTimeout(this.statusTimeout);
            this.statusTimeout = null;
        }
    }

    /**
     * Creates a status message element
     * @param {string} message - Status message
     * @param {string} type - Message type
     * @returns {Element} Status element
     */
    createStatusElement(message, type) {
        const statusElement = document.createElement('div');
        statusElement.className = `status-message status-${type}`;
        statusElement.setAttribute('role', type === 'error' ? 'alert' : 'status');
        statusElement.setAttribute('aria-live', type === 'error' ? 'assertive' : 'polite');

        // Add icon based on type
        const icon = this.getStatusIcon(type);
        
        statusElement.innerHTML = `
            <div class="status-content">
                <span class="status-icon" aria-hidden="true">${icon}</span>
                <span class="status-text">${this.escapeHtml(message)}</span>
                <button class="status-close" type="button" aria-label="Close message" onclick="this.parentElement.parentElement.remove()">
                    <span aria-hidden="true">×</span>
                </button>
            </div>
        `;

        return statusElement;
    }

    /**
     * Gets icon for status type
     * @param {string} type - Status type
     * @returns {string} Icon HTML
     */
    getStatusIcon(type) {
        const icons = {
            info: 'ℹ️',
            success: '✅',
            warning: '⚠️',
            error: '❌'
        };
        return icons[type] || icons.info;
    }

    /**
     * Escapes HTML in text
     * @param {string} text - Text to escape
     * @returns {string} Escaped text
     */
    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    /**
     * Announces message to screen readers
     * @param {string} message - Message to announce
     * @param {boolean} assertive - Whether to use assertive announcement
     */
    announceToScreenReader(message, assertive = false) {
        const regionId = assertive ? 'sr-announcements' : 'sr-announcements-polite';
        const region = document.getElementById(regionId);
        
        if (region) {
            // Clear and set new message
            region.textContent = '';
            setTimeout(() => {
                region.textContent = message;
            }, 100);
        }
    }

    /**
     * Shows/hides an element
     * @param {string} elementId - Element ID
     * @param {boolean} show - Whether to show the element
     */
    toggleElement(elementId, show) {
        const element = document.getElementById(elementId);
        if (element) {
            element.classList.toggle('hidden', !show);
        }
    }

    /**
     * Updates progress indicator
     * @param {string} progressId - Progress element ID
     * @param {number} percentage - Progress percentage (0-100)
     * @param {string} label - Progress label
     */
    updateProgress(progressId, percentage, label = '') {
        const progressElement = document.getElementById(progressId);
        if (!progressElement) return;

        const progressBar = progressElement.querySelector('.progress-bar');
        const progressLabel = progressElement.querySelector('.progress-label');

        if (progressBar) {
            progressBar.style.width = `${Math.max(0, Math.min(100, percentage))}%`;
            progressBar.setAttribute('aria-valuenow', percentage);
        }

        if (progressLabel && label) {
            progressLabel.textContent = label;
        }

        // Announce significant progress milestones
        if (percentage === 0) {
            this.announceToScreenReader('Process started');
        } else if (percentage === 100) {
            this.announceToScreenReader('Process completed');
        } else if (percentage % 25 === 0) {
            this.announceToScreenReader(`${percentage}% complete`);
        }
    }

    /**
     * Sets form field validation state
     * @param {string} fieldId - Field ID
     * @param {boolean} isValid - Whether field is valid
     * @param {string} errorMessage - Error message if invalid
     */
    setFieldValidation(fieldId, isValid, errorMessage = '') {
        const field = document.getElementById(fieldId);
        if (!field) return;

        field.classList.toggle('error', !isValid);
        field.classList.toggle('valid', isValid);

        // Handle error message
        const errorId = fieldId + '-error';
        let errorElement = document.getElementById(errorId);

        if (!isValid && errorMessage) {
            if (!errorElement) {
                errorElement = document.createElement('div');
                errorElement.id = errorId;
                errorElement.className = 'field-error';
                errorElement.setAttribute('role', 'alert');
                field.parentNode.appendChild(errorElement);
            }
            
            errorElement.textContent = errorMessage;
            field.setAttribute('aria-describedby', errorId);
            
        } else if (errorElement) {
            errorElement.remove();
            field.removeAttribute('aria-describedby');
        }

        // Update ARIA attributes
        field.setAttribute('aria-invalid', !isValid);
    }

    /**
     * Creates a modal dialog
     * @param {Object} options - Modal options
     * @returns {Element} Modal element
     */
    createModal(options = {}) {
        const {
            title = 'Dialog',
            content = '',
            buttons = [{ text: 'OK', action: 'close' }],
            size = 'medium',
            closable = true
        } = options;

        // Create modal structure
        const modal = document.createElement('div');
        modal.className = `modal modal-${size}`;
        modal.setAttribute('role', 'dialog');
        modal.setAttribute('aria-labelledby', 'modal-title');
        modal.setAttribute('aria-modal', 'true');

        const modalContent = `
            <div class="modal-overlay" onclick="this.parentElement.remove()"></div>
            <div class="modal-content">
                <header class="modal-header">
                    <h2 id="modal-title">${this.escapeHtml(title)}</h2>
                    ${closable ? '<button class="modal-close" onclick="this.closest(\'.modal\').remove()" aria-label="Close dialog">&times;</button>' : ''}
                </header>
                <div class="modal-body">
                    ${content}
                </div>
                <footer class="modal-footer">
                    ${buttons.map(btn => `
                        <button class="btn btn-${btn.type || 'secondary'}" 
                                onclick="${btn.action === 'close' ? 'this.closest(\'.modal\').remove()' : btn.action}">
                            ${this.escapeHtml(btn.text)}
                        </button>
                    `).join('')}
                </footer>
            </div>
        `;

        modal.innerHTML = modalContent;
        
        // Add to document
        document.body.appendChild(modal);

        // Focus management
        setTimeout(() => {
            const firstButton = modal.querySelector('button');
            if (firstButton) {
                firstButton.focus();
            }
        }, 100);

        return modal;
    }

    /**
     * Shows a confirmation dialog
     * @param {string} message - Confirmation message
     * @param {string} title - Dialog title
     * @returns {Promise<boolean>} User's choice
     */
    showConfirmation(message, title = 'Confirm') {
        return new Promise((resolve) => {
            const modal = this.createModal({
                title: title,
                content: `<p>${this.escapeHtml(message)}</p>`,
                buttons: [
                    {
                        text: 'Cancel',
                        type: 'secondary',
                        action: 'this.closest(\'.modal\').remove(); window.confirmResult = false;'
                    },
                    {
                        text: 'Confirm',
                        type: 'primary',
                        action: 'this.closest(\'.modal\').remove(); window.confirmResult = true;'
                    }
                ]
            });

            // Handle result
            const checkResult = () => {
                if (window.confirmResult !== undefined) {
                    const result = window.confirmResult;
                    delete window.confirmResult;
                    resolve(result);
                } else if (!document.contains(modal)) {
                    resolve(false);
                } else {
                    setTimeout(checkResult, 100);
                }
            };

            setTimeout(checkResult, 100);
        });
    }

    /**
     * Animates element changes
     * @param {string} elementId - Element ID
     * @param {string} animation - Animation type
     * @param {number} duration - Animation duration in ms
     */
    animateElement(elementId, animation = 'fadeIn', duration = 300) {
        const element = document.getElementById(elementId);
        if (!element) return;

        // Check for reduced motion preference
        const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
        if (prefersReducedMotion) {
            return; // Skip animations
        }

        element.classList.add(`animate-${animation}`);
        
        setTimeout(() => {
            element.classList.remove(`animate-${animation}`);
        }, duration);
    }

    /**
     * Updates tab state
     * @param {string} activeTabId - ID of the active tab
     */
    setActiveTab(activeTabId) {
        const allTabs = document.querySelectorAll('.tab-button');
        const allPanels = document.querySelectorAll('.tab-panel');

        // Update tab buttons
        allTabs.forEach(tab => {
            const isActive = tab.id === activeTabId;
            tab.classList.toggle('active', isActive);
            tab.setAttribute('aria-selected', isActive);
        });

        // Update tab panels
        allPanels.forEach(panel => {
            const isActive = panel.getAttribute('aria-labelledby') === activeTabId;
            panel.classList.toggle('active', isActive);
        });
    }

    /**
     * Scrolls element into view smoothly
     * @param {string} elementId - Element ID to scroll to
     * @param {string} behavior - Scroll behavior ('smooth' or 'auto')
     */
    scrollToElement(elementId, behavior = 'smooth') {
        const element = document.getElementById(elementId);
        if (element) {
            element.scrollIntoView({ 
                behavior: behavior, 
                block: 'start',
                inline: 'nearest'
            });
        }
    }

    /**
     * Updates slider display value
     * @param {string} sliderId - Slider ID
     * @param {Array} valueLabels - Array of value labels
     */
    updateSliderValue(sliderId, valueLabels = []) {
        const slider = document.getElementById(sliderId);
        const valueDisplay = document.getElementById(sliderId.replace('response-', '') + '-value');
        
        if (slider && valueDisplay) {
            const value = parseInt(slider.value) - 1;
            const label = valueLabels[value] || slider.value;
            valueDisplay.textContent = label;
            
            // Update ARIA label for screen readers
            slider.setAttribute('aria-valuetext', label);
        }
    }

    /**
     * Gets current UI state
     * @returns {Object} Current UI state
     */
    getUIState() {
        return {
            loadingStates: Object.fromEntries(this.loadingStates),
            visibleElements: Array.from(document.querySelectorAll(':not(.hidden)')).map(el => el.id).filter(Boolean),
            activeTab: document.querySelector('.tab-button.active')?.id,
            hasStatusMessage: !document.getElementById('status-messages')?.classList.contains('hidden')
        };
    }
}
