# Application CMS - Google Apps Script Setup Instructions

## 🚀 Complete Setup Guide

### Step 1: Create Google Apps Script

1. Go to [Google Apps Script](https://script.google.com)
2. Click "New Project"
3. Copy the code from `GOOGLE_APPS_SCRIPT_TEMPLATE.js`
4. Paste it into the script editor

### Step 2: Configure Shared Secret

1. In the Apps Script editor, select `setup` function from the dropdown
2. Click "Run" 
3. Change `YOUR_SECRET_KEY_HERE` to a secure random string
4. Run the `setup` function again to set your secret

### Step 3: Deploy as Web App

1. Click "Deploy" → "New Deployment"
2. Select "Web App"
3. Configuration:
   - Description: "Application CMS Integration"
   - Execute as: "Me"
   - Who has access: "Anyone"
4. Click "Deploy"
5. Copy the Web App URL

### Step 4: Create Google Sheet

1. Create a new Google Sheet
2. Copy the Sheet ID from the URL: `.../spreadsheets/d/{SHEET_ID}/edit`
3. Share the sheet with your Apps Script email

### Step 5: Create Drive Folder

1. Create a folder in Google Drive for case files
2. Copy the Folder ID from the URL: `.../drive/folders/{FOLDER_ID}`
3. Share the folder with your Apps Script email

### Step 6: Configure in Application CMS

1. Open your Application CMS
2. Go to Settings → Google Apps Script Integration
3. Enter:
   - Apps Script Web App URL (from Step 3)
   - Shared Secret (from Step 2)
   - Google Sheet ID (from Step 4)
   - Drive Folder ID (from Step 5)
4. Click "Save Settings"
5. Click "Test Connection"

### Step 7: Enable Auto-Sync

1. Toggle "Enable automatic sync on case changes"
2. Test by creating a new case
3. Check your Google Sheet for the synced data

## 🔧 Troubleshooting

### Connection Test Fails
- Verify the Web App URL is correct
- Check the shared secret matches exactly
- Ensure the Apps Script is deployed properly

### Sync Not Working
- Check if auto-sync is enabled
- Verify Google Sheet ID is correct
- Ensure the sheet is shared with Apps Script email

### File Upload Issues
- Verify Drive Folder ID is correct
- Ensure the folder is shared with Apps Script email
- Check file size limits

## 📊 Data Structure

### Cases Sheet Columns
- ID, Date/Time, Case No, Application Name, TM No, Class, Status, Sub Status, IPO Office, Delivery Tracking, Notes, Created, Updated, Sync Status

### Audit Sheet Columns  
- Case ID, Action, Old Values, New Values, Timestamp, Actor

## 🔐 Security Notes

- Keep your shared secret secure
- Regularly rotate the secret key
- Monitor access logs in Google Apps Script
- Limit sheet access to authorized users only
