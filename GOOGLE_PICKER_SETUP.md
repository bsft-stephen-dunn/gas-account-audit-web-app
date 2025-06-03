# Google Picker API Setup Instructions

This document explains how to set up Google Picker API for the folder selection feature.

## Current Status

The app currently uses a **fallback folder selector** that works without any additional configuration. The Google Picker API is optional and provides a more native Google Drive experience.

## Option 1: Use the Default Folder Selector (Recommended)

No configuration needed! The app will automatically use the dropdown folder selector which:
- Lists your Google Drive folders
- Allows you to select where to save documents
- Works immediately without any setup

## Option 2: Enable Google Picker API (Advanced)

If you want to use the native Google Picker interface, follow these steps:

### Step 1: Set Up Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com)
2. Create a new project or select an existing one
3. Note your **Project Number** (this is your App ID)

### Step 2: Enable Required APIs

In your Google Cloud Project:
1. Go to "APIs & Services" > "Library"
2. Search for and enable:
   - Google Picker API
   - Google Drive API

### Step 3: Create API Credentials

1. Go to "APIs & Services" > "Credentials"
2. Click "Create Credentials" > "API Key"
3. Copy the API key
4. Click on the API key to edit it
5. Under "Application restrictions":
   - Select "HTTP referrers (websites)"
   - Add these referrers:
     - `https://script.google.com/*`
     - `https://*.googleusercontent.com/*`
6. Under "API restrictions":
   - Select "Restrict key"
   - Select "Google Picker API"

### Step 4: Configure the App

#### Method A: Script Properties (Secure - Recommended)

1. In your Google Apps Script editor, go to Project Settings
2. Scroll down to "Script Properties"
3. Add these properties:
   - Property: `GOOGLE_PICKER_API_KEY`, Value: `[Your API Key]`
   - Property: `GOOGLE_PICKER_APP_ID`, Value: `[Your Project Number]`

#### Method B: Direct Configuration (Less Secure)

1. Edit `JavaScript.html`
2. Find the `GOOGLE_PICKER_CONFIG` object
3. Update:
   ```javascript
   const GOOGLE_PICKER_CONFIG = {
     API_KEY: 'YOUR_API_KEY_HERE',
     APP_ID: 'YOUR_PROJECT_NUMBER_HERE',
     ENABLED: true
   };
   ```

### Step 5: Deploy and Test

1. Save all files
2. Deploy your web app
3. Test the folder picker functionality

## Troubleshooting

### Common Issues:

1. **Blank Modal**: Usually means API key or App ID is incorrect
2. **Permission Denied**: Check API key restrictions
3. **API Not Enabled**: Ensure Google Picker API is enabled in Cloud Console

### Debug Steps:

1. Open browser console (F12)
2. Look for error messages when clicking "Choose Folder"
3. Check that configuration is loaded: `console.log(GOOGLE_PICKER_CONFIG)`

## Security Notes

- Never commit API keys to public repositories
- Use Script Properties method for production deployments
- Restrict API keys to your specific domains
- Regularly rotate API keys

## Need Help?

If the Google Picker setup is too complex, the fallback folder selector works great and requires no configuration!