/**
 * Classification Detector Service
 * Detects and logs classification text in email content
 */

export class ClassificationDetector {
    constructor() {
        // Classification patterns - flexible to handle real-world formatting
        this.patterns = [
            // Primary pattern: "Classification: LEVEL" format
            /^\s*Classification:\s*(.+?)\s*$/i,
            
            // Alternative patterns
            /^\s*CLASS:\s*(.+?)\s*$/i,
            /^\s*Security\s*Classification:\s*(.+?)\s*$/i,
            
            // Classification with additional markings 
            /^\s*Classification:\s*([^\/]+)\/\/([A-Z\s\/]+)/i
        ];
    }

    /**
     * Detects classification in email content and logs to console
     * @param {string} emailBody - The email body text
     * @returns {Object} Simple classification detection result
     */
    detectClassification(emailBody) {
        if (!emailBody || typeof emailBody !== 'string') {
            return {
                detected: false,
                text: null,
                details: 'No content to analyze'
            };
        }

        // Try to parse the structured classification block
        const result = this.parseClassificationText(emailBody);
        

        
        return result;
    }

    /**
     * Parses classification text from email content
     * Classification must be on the first line only
     * @param {string} emailBody - The email body text
     * @returns {Object} Simple classification detection result
     */
    parseClassificationText(emailBody) {
        // Get only the first line
        const lines = emailBody.split('\n');
        const firstLine = lines[0] ? lines[0].trim() : '';
        
        if (!firstLine) {
            return {
                detected: false,
                text: null,
                details: 'No content on first line'
            };
        }

        // Check each pattern against the first line only
        for (const pattern of this.patterns) {
            const match = firstLine.match(pattern);
            if (match) {
                const classificationText = match[1] ? match[1].trim() : match[0].trim();
                
                if (classificationText) {
                    return {
                        detected: true,
                        text: classificationText,
                        line: 1,
                        details: `Classification found on first line: ${classificationText}`
                    };
                }
            }
        }

        return {
            detected: false,
            text: null,
            details: 'No classification pattern found on first line'
        };
    }

    /**
     * Generates a simple classification report
     * @param {string} emailBody - Email content
     * @returns {Object} Simple report
     */
    generateReport(emailBody) {
        const detection = this.detectClassification(emailBody);

        return {
            ...detection,
            timestamp: new Date().toISOString(),
            analyzer: 'PromptEmail Classification Detector v2.0'
        };
    }
}

// For Node.js testing compatibility
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ClassificationDetector;
}
