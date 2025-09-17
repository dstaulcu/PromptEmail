# PromptEmail User Guide

## Getting Started

PromptEmail is an AI-powered Outlook add-in that helps you analyze emails, detect sentiment, and generate intelligent responses using advanced language models.

## Installation

### Prerequisites
- Microsoft Outlook (Office 365 or Outlook 2019+)
- Windows 10/11 or macOS
- Internet connection for AI services

### Installing the Add-in

#### Method 1: Sideloading (Development/Testing)
1. Download the `manifest.xml` file from your administrator
2. Open Outlook
3. Go to **File** → **Manage Add-ins** → **Get Add-ins**
4. Click **My add-ins** → **Add a custom add-in** → **Add from file**
5. Select the `manifest.xml` file
6. Click **Install**

#### Method 2: Microsoft AppSource (Production)
1. Open Outlook
2. Go to **Home** → **Get Add-ins**
3. Search for "PromptEmail"
4. Click **Add** to install

## Quick Start Guide

### First Launch
1. Open an email in Outlook
2. Look for the **PromptEmail** button in the ribbon
3. Click the button to open the add-in panel
4. Follow the initial setup wizard to configure your AI provider

### Basic Email Analysis
1. **Select an email** you want to analyze
2. **Click "Analyze Email"** in the PromptEmail panel
3. **Review the results** including:
   - Email summary
   - Tone and sentiment analysis
   - Key topics identified
   - Suggested actions

### Generating Responses
1. After analyzing an email, click **"Generate Response"**
2. Choose your response type:
   - **Professional** - Formal business tone
   - **Friendly** - Casual but polite
   - **Custom** - Specify your requirements
3. Review and edit the generated response
4. Copy to clipboard or insert directly into your reply

### Writing Samples for Personalized Responses
PromptEmail can learn your writing style to generate more authentic responses that sound like you wrote them.

#### Setting Up Writing Samples
1. **Open Settings** in the PromptEmail panel
2. **Navigate to Writing Style** section
3. **Enable Style Analysis** toggle
4. **Add your writing samples**:
   - Click **"Add Writing Sample"**
   - Paste examples of your professional emails or messages
   - Give each sample a descriptive title
   - Click **"Save Sample"**
5. **Configure Style Strength**:
   - **Light**: Subtle influence (uses 2 samples)
   - **Medium**: Balanced personalization (uses 3 samples)
   - **Strong**: Maximum style matching (uses 5 samples)

#### Best Practices for Writing Samples
- **Quality over Quantity**: Add 3-5 high-quality samples that represent your communication style
- **Variety**: Include different types of responses (formal, informal, brief, detailed)
- **Recent Examples**: Use recent writing that reflects your current communication style
- **Professional Content**: Focus on work-appropriate content for business email responses
- **Length**: Aim for 50-200 words per sample for optimal AI training

## Configuration

### AI Provider Setup

#### Option 1: Local AI (Ollama) - Recommended for Privacy
1. Install [Ollama](https://ollama.ai/) on your computer
2. Download a model: `ollama pull llama3`
3. In PromptEmail settings, select **"Ollama (Local)"**
4. Verify connection - should show green status

#### Option 2: Enterprise AI Services
1. Contact your IT administrator for API credentials
2. In PromptEmail settings, select your organization's AI provider
3. Enter the provided API key or endpoint details
4. Test connection

#### Option 3: External AI Services (OpenAI Compatible)
1. Obtain API key from your chosen provider
2. Select **"Custom OpenAI Compatible"** in settings
3. Configure endpoint URL and API key
4. Test connection

### Privacy Settings

#### Debug Logging
- **Enable**: Shows detailed operation logs in browser developer console for troubleshooting
- **Disable**: Minimal logging for normal operation
- **Recommendation**: Keep disabled unless troubleshooting

#### Telemetry
- **Usage Analytics**: Tracks add-in deployment, user engagement, AI provider attribution, and model performance
- **Performance Metrics**: Measures response times

### Writing Style Settings

#### Style Analysis
- **Enable**: Allows AI to use your writing samples for personalized responses
- **Disable**: AI uses standard response generation without personalization
- **Recommendation**: Enable for more authentic, personalized responses

#### Style Strength
- **Light**: Uses 2 writing samples for subtle style influence
- **Medium**: Uses 3 writing samples for balanced personalization (recommended)
- **Strong**: Uses 5 writing samples for maximum style matching

#### Writing Samples Management
- **Add Sample**: Input examples of your professional writing
- **Edit Sample**: Modify existing samples to better represent your style  
- **Delete Sample**: Remove samples that no longer represent your preferred style
- **Sample Limit**: No hard limit, but 3-7 quality samples typically provide optimal results

## Features Guide

### Email Analysis Features

#### Sentiment Analysis
- **Positive**: Friendly/enthusiastic tone detected
- **Neutral**: Professional/informational tone
- **Negative**: Frustrated/concerned tone detected
- **Mixed**: Multiple sentiments detected in the email

#### Content Summary
- Extracts key points from long emails
- Identifies action items and deadlines
- Highlights important information

#### Topic Detection
- Automatically categorizes email content
- Identifies project names, people, and organizations
- Flags urgent or time-sensitive items

#### Customization Options
- **Tone**: Professional, Friendly, Formal, Casual
- **Length**: Brief, Standard, Detailed
- **Style**: Direct, Diplomatic, Enthusiastic, Cautious
- **Include**: Specific points, questions, action items

## User Interface Guide

### Main Panel Layout

```
┌─────────────────────────────┐
│ PromptEmail                 │
├─────────────────────────────┤
│ Email Analysis              │
│ ┌─────────────────────────┐ │
│ │ [Analyze Email]         │ │
│ └─────────────────────────┘ │
│                             │
│ Results                     │
│ ┌─────────────────────────┐ │
│ │ Summary: ...            │ │
│ │ Tone: Professional      │ │
│ │ Topics: Project Update  │ │
│ └─────────────────────────┘ │
│                             │
│ Response                    │
│ ┌─────────────────────────┐ │
│ │ [Generate Response]     │ │
│ │ [Copy to Clipboard]     │ │
│ └─────────────────────────┘ │
│                             │
│ Settings                    │
│ ┌─────────────────────────┐ │
│ │ Writing Style           │ │
│ │ ☑ Style Analysis        │ │
│ │ Strength: [Medium ▼]    │ │
│ │ Writing Samples (4)     │ │
│ │ [+ Add Sample]          │ │
│ └─────────────────────────┘ │
└─────────────────────────────┘
```

## Troubleshooting

### Common Issues

#### "AI Service Not Connected"
**Symptoms**: Cannot analyze emails
**Solutions**:
1. Check internet connection
2. Verify AI provider settings
3. Restart Outlook
4. Contact IT support if using enterprise AI

#### "Email Analysis Failed"
**Symptoms**: Error message when clicking "Analyze Email"
**Solutions**:
1. Try selecting the email again
2. Check if email content is accessible
3. Try with a different email
4. Enable debug logging and retry

#### "Slow Response Times"
**Symptoms**: Long wait times for analysis or response generation
**Solutions**:
1. Check internet connection speed
2. Try using a local AI model (Ollama)
3. Reduce email length by selecting specific text
4. Contact admin about AI service performance

#### "Settings Won't Save"
**Symptoms**: Configuration changes don't persist
**Solutions**:
1. Ensure Outlook has proper permissions
2. Check if running in restricted mode
3. Try closing and reopening Outlook
4. Contact IT about Office 365 roaming settings

#### "Writing Samples Not Working"
**Symptoms**: Generated responses don't reflect your writing style
**Solutions**:
1. **Check Style Analysis**: Ensure "Enable Style Analysis" is turned on
2. **Add More Samples**: Include at least 3 quality writing examples
3. **Adjust Style Strength**: Try increasing from Light to Medium or Strong
4. **Sample Quality**: Ensure samples are 50-200 words and represent your style
5. **Clear Cache**: Close and reopen Outlook, then try again

#### "Writing Samples Won't Save"
**Symptoms**: Added samples disappear after closing Outlook
**Solutions**:
1. **Wait for Save**: Allow a few seconds after adding samples before closing
2. **Check Permissions**: Ensure Outlook can access Office 365 settings
3. **Network Connection**: Verify internet connectivity for settings sync
4. **Try Again**: Close Outlook completely and reopen, then re-add samples
5. **Contact IT**: May need Office 365 roaming settings enabled

#### "Responses Too Similar/Different"
**Symptoms**: Generated responses don't match desired style level
**Solutions**:
1. **Adjust Style Strength**: 
   - Use "Light" for subtle influence
   - Use "Medium" for balanced personalization
   - Use "Strong" for maximum style matching
2. **Review Samples**: Ensure samples represent the style you want
3. **Add Variety**: Include different types of communication in your samples
4. **Sample Length**: Aim for 50-200 words per sample

### Error Messages

#### "Invalid API Key"
- **Cause**: Incorrect or expired API credentials
- **Solution**: Update API key in settings or contact administrator

#### "Service Temporarily Unavailable"
- **Cause**: AI service maintenance or outage
- **Solution**: Try again later or switch to alternative provider

#### "Email Content Not Accessible"
- **Cause**: Protected email or insufficient permissions
- **Solution**: Check email security settings or try with different email

#### "Rate Limit Exceeded"
- **Cause**: Too many requests to AI service
- **Solution**: Wait a few minutes before trying again

### Getting Help

#### Self-Service Resources
1. **Settings Panel**: Check connection status and configuration
2. **Debug Logs**: Enable verbose logging to browser developer console
3. **Help Links**: Click the wiki icon for online documentation

#### Contacting Support
1. **Help Desk**: Use your organization's standard IT support process
2. **Email**: Include error messages and steps to reproduce
3. **Screenshots**: Capture any error dialogs or unexpected behavior
4. **Debug Logs**: Enable debug logging before reproducing the issue

## Privacy & Security

### Data Handling
- **AI Processing**: Email content and writing samples sent to configured AI provider only
- **Settings Storage**: Dual-layer storage using Office.js RoamingSettings (primary) with browser localStorage (fallback)
- **Writing Samples**: Stored locally with cross-device synchronization when Office 365 roaming is available
- **No External Sharing**: Writing samples are never shared with third parties except your chosen AI provider
- **User Control**: You can delete writing samples at any time through the settings panel
- **Telemetry**: Usage and performance statistics only (no writing sample content)

## Tips & Best Practices

### Maximizing Accuracy
1. **Clear Context**: Provide additional context for ambiguous content
2. **Iterative Refinement**: Use "Regenerate" to get alternative analyses

### Professional Usage
1. **Review Generated Content**: Always review AI suggestions before sending
2. **Customize Tone**: Adjust response style to match your communication style
3. **Context Awareness**: Ensure responses match the email thread context
4. **Confidentiality**: Be mindful of sensitive information in AI requests

### Writing Samples Best Practices
1. **Curate Your Samples**: Choose writing that represents how you want to sound professionally
2. **Regular Updates**: Refresh samples as your communication style evolves
3. **Sample Variety**: Include different scenarios (responses to requests, updates, confirmations, etc.)
4. **Length Balance**: Mix brief and detailed samples to handle various response types
5. **Professional Focus**: Use workplace-appropriate content that matches your role and industry
6. **Quality Check**: Re-read your samples to ensure they represent your best professional communication
7. **Privacy Awareness**: Only include content you're comfortable sharing with your AI provider

### Writing Style Strength Guide
- **Light (2 samples)**: 
  - Use when you want subtle style hints
  - Good for maintaining professional consistency
  - Minimal impact on AI-generated content
  
- **Medium (3 samples)**: 
  - Balanced approach for most users
  - Noticeable style influence without being overwhelming
  - Recommended for daily professional use
  
- **Strong (5 samples)**: 
  - Maximum personalization
  - AI closely mimics your writing patterns
  - Use when authenticity is crucial

---

## Support Information

**Documentation**: Check latest guides at project repository
**Enterprise Support**: Contact your IT administrator for business-specific issues

---

*Version 1.2.7 - Last updated: September 2025*
