// Office command functions
Office.onReady(() => {
    // Ready to register command functions
});

/**
 * Shows a notification in Word
 */
async function showNotification(message: string, type: 'info' | 'error' = 'info') {
    return Word.run(async (context) => {
        // Note: Notifications are not supported in all Word versions
        console.log(`[${type.toUpperCase()}] ${message}`);
    });
}

/**
 * Example command function that could be bound to a button
 */
async function insertSampleText(event: Office.AddinCommands.Event) {
    return Word.run(async (context) => {
        const paragraph = context.document.body.insertParagraph(
            "AIMTA Document Processor - Sample Text",
            Word.InsertLocation.end
        );
        paragraph.font.color = "blue";
        
        await context.sync();
        event.completed();
    }).catch((error) => {
        console.error('Error:', error);
        event.completed();
    });
}

// Register the command functions
Office.actions.associate("insertSampleText", insertSampleText);

// Export for use in other parts of the add-in
export { showNotification, insertSampleText };