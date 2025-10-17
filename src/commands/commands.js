// Commands for Office Add-in ribbon integration
// This handles function commands that can be triggered from the ribbon

Office.onReady(() => {
    console.info('PromptEmail commands ready');
});

// Register command functions
if (typeof Office !== 'undefined' && Office.actions) {
    Office.actions.associate('showTaskpane', showTaskpane);
}

/**
 * Shows the task pane
 * @param {Office.AddinCommands.Event} event - The event object
 */
function showTaskpane(event) {
    // This function is called when the ribbon button is clicked
    // The taskpane will be shown automatically by Office
    
    // Complete the event
    if (event && event.completed) {
        event.completed();
    }
}

// Export for module systems
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        showTaskpane
    };
}
