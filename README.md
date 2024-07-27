# Email-Details

This project automates the processing of emails and uploads relevant data to Google Sheets and Google Drive using GitHub Actions. 
The workflow fetches emails from a specified account, processes attachments, and stores email details in a Google Sheet and ttachments are uploaded to Google Drive.

## Features

- **Fetches emails**: Retrieves all emails from an IMAP email account.
- **Processes email content**: Extracts email subject, sender, date, and body.
- **Handles attachments**: Saves attachments to a specified Google Drive folder.
- **Updates Google Sheets**: Logs email details and attachment links to Google Sheets.

## Setup

### 1. Create a Google Cloud Project
1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Sign in with your Google account.
3. Click the project drop-down and select "New Project."
4. Enter the project name and click "Create."

### 2. Enable Required APIs
1. In the Google Cloud Console, select your project.
2. Navigate to `APIs & Services` > `Library`.
3. Enable the following APIs:
   - **Gmail API**
   - **Google Sheets API**
   - **Google Drive API**

### 3. Set Up the OAuth Consent Screen
1. In the Google Cloud Console, navigate to `APIs & Services` > `OAuth consent screen`.
2. Choose "External" and click "Create".
3. Fill in the required fields such as `App name`, `User support email`.
4. In the `Scopes for Google APIs` section, add the following scopes:
   - **Google Sheets API**: `https://www.googleapis.com/auth/spreadsheets`
   - **Google Drive API**: `https://www.googleapis.com/auth/drive`
   - **Gmail API**: `https://www.googleapis.com/auth/gmail.readonly`
5. Complete the OAuth consent screen setup by clicking "Save and Continue."

### 4. Create a Service Account and 
1. In the Google Cloud Console, navigate to `APIs & Services` > `Credentials`.
2. Click on "Create Credentials" and select "Service Account."
3. Enter the service account name and click "Create and Continue."
4. Assign the required role `Editor`.
5. Click "Done" after assigning roles.
   
### 5. Download the JSON Key File
1. In the Service Accounts list, click on the service account you just created.
2. Go to the `Keys` tab and click "Add Key" > "Create New Key."
3. Choose `JSON` and click "Create." The JSON key file will be downloaded to your computer(Save it securely).

### 6. Share the Google Sheets document and Google Drive folder with the email address associated with your Google service account.

### 7. Generate an App Password

1. Go to the [Google Account Security page](https://myaccount.google.com/security).
2. Enable Two-Factor Authentication. 
3. Under the "Signing in to Google" section, find and click on "App passwords."
4. Under the "Select app" drop-down menu, choose "Custom name)."
5. Enter a name for the app password and click "Generate."
6. Google will generate a 16-character app password (Copy this password and store it securely).

### 8. Store Credentials

**Add GitHub Secrets**:
   - Go to your GitHub repository.
   - Navigate to `Settings` > `Secrets and variables` > `Actions`.
   - Add the following secrets:
     - **`GOOGLE_CREDENTIALS_JSON`**: Paste the single-line JSON string from the previous step.
     - **`EMAIL_USERNAME`**: Your email address.
     - **`EMAIL_PASSWORD`**: Your email password (use an app-specific password if you have 2-factor authentication enabled).

### 9. Create GitHub Actions Workflow

1. **Add Workflow File**:
   - In your repository, create the following directory structure:
     ```
     .github/
       workflows/
         process-emails.yml
     ```

2. **Workflow Process**:
   The process-emails.yml workflow in GitHub Actions triggers every 5 minutes or manually, sets up Python, installs dependencies, handles credentials securely, and runs python script.
   
### 10. Python Script
The project.py script is responsible for fetching emails, processing content and attachments, and updating Google Sheets and Drive.

## Run the Workflow

**Automatic Runs:** The workflow runs every 5 minutes as scheduled.
**Manual Runs:** Trigger the workflow manually from the GitHub Actions tab.

