# Email-Details

This project automates the processing of emails and uploads relevant data to Google Sheets and Google Drive using GitHub Actions. 
The workflow fetches emails from a specified account, processes attachments, and stores email details in a Google Sheet and ttachments are uploaded to Google Drive.

## Features

- **Fetches emails**: Retrieves all emails from an IMAP email account.
- **Processes email content**: Extracts email subject, sender, date, and body.
- **Handles attachments**: Saves attachments to a specified Google Drive folder.
- **Updates Google Sheets**: Logs email details and attachment links to Google Sheets.

## Prerequisites

1. **Google Account**: For Google Sheets and Drive API access.
2. **Gmail Account**: For accessing and processing emails.

## Setup

### 1. Prepare Your Environment

1. **Create a Google Service Account**:
   - Go to the [Google Cloud Console](https://console.cloud.google.com/).
   - Create a new project or select an existing project.
   - Navigate to `APIs & Services` > `Credentials`.
   - Create a new service account and download the JSON key file.

2. **Share Google Sheets**:
   - Share the Google Sheets document with the email address associated with your Google service account.

3. **Enable APIs**:
   - Enable the Google Sheets API and Google Drive API for your project.

### 2. Store Credentials

**Add GitHub Secrets**:
   - Go to your GitHub repository.
   - Navigate to `Settings` > `Secrets and variables` > `Actions`.
   - Add the following secrets:
     - **`GOOGLE_CREDENTIALS_JSON`**: Paste the single-line JSON string from the previous step.
     - **`EMAIL_USERNAME`**: Your email address.
     - **`EMAIL_PASSWORD`**: Your email password (use an app-specific password if you have 2-factor authentication enabled).

### 3. Create GitHub Actions Workflow

1. **Add Workflow File**:
   - In your repository, create the following directory structure:
     ```
     .github/
       workflows/
         process-emails.yml
     ```

2. **Workflow Process**:
   The process-emails.yml workflow in GitHub Actions triggers every 5 minutes or manually, sets up Python, installs dependencies, handles credentials securely, and runs python script.
   
### 4. Python Script
The project.py script is responsible for fetching emails, processing content and attachments, and updating Google Sheets and Drive.

## Run the Workflow

**Automatic Runs:** The workflow runs every 5 minutes as scheduled.
**Manual Runs:** Trigger the workflow manually from the GitHub Actions tab.

