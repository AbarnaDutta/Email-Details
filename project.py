import os
import imaplib
import email
from email.header import decode_header
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.service_account import Credentials

# Account credentials
username = os.getenv('EMAIL_USERNAME')  # Fetch username from environment variables for security
password = os.getenv('EMAIL_PASSWORD')  # Fetch password from environment variables for security

# Google Sheets and Drive API setup
scope = ["https://spreadsheets.google.com/feeds", 
         "https://www.googleapis.com/auth/drive"]

# Authorize using the credentials.json file
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Initialize Google Drive API service
drive_creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
drive_service = build('drive', 'v3', credentials=drive_creds)

# Open the Google Sheets document using its URL
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1saKVvG0D-serGiqPYkPmcJKDoyojwbsZLjM2nvIm5RQ/edit?gid=0")

# Create an IMAP4 class with SSL for secure email connection
mail = imaplib.IMAP4_SSL("imap.gmail.com")

# Authenticate using the provided username and password
mail.login(username, password)

# Select the mailbox
mail.select("inbox")
status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()

# Decode email subject
def decode_subject(subject):
    decoded, encoding = decode_header(subject)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(encoding if encoding else "utf-8")
    return decoded

# Decode email date
def decode_date(date_):
    if 'GMT' in date_:
        date_ = date_.replace('GMT', '+0000')
    return datetime.strptime(date_, '%a, %d %b %Y %H:%M:%S %z')

# Create a folder in Google Drive
def create_drive_folder(folder_name, parent_folder_id):
    query = f"name='{folder_name}' and '{parent_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
    results = drive_service.files().list(q=query, fields="files(id, webViewLink)").execute()
    items = results.get('files', [])
    if items:
        return items[0]['id'], items[0]['webViewLink']
    else:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_folder_id]
        }
        folder = drive_service.files().create(body=file_metadata, fields='id, webViewLink').execute()
        return folder.get('id'), folder.get('webViewLink')

# Upload a file to Google Drive
def upload_to_drive(file_data, file_name, folder_id):
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaIoBaseUpload(io.BytesIO(file_data), mimetype='application/octet-stream')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f'File {file_name} uploaded to Google Drive with ID: {file.get("id")}')

# Google Drive parent folder ID
drive_folder_id = '1V8PmM2wLhuv8iWJbm_MqKszlhWZ6iZ5b'

# Initialize email_folder_link
email_folder_link = "None"

# Fetch and process each email
for email_id in email_ids:
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            subject = decode_subject(msg["Subject"])
            from_ = msg.get("From")
            date_ = msg.get("Date")
            email_time = decode_date(date_).strftime("%H:%M:%S")
            email_date = decode_date(date_).strftime("%Y-%m-%d")

            # Get or create the worksheet for this date
            try:
                ws = spreadsheet.worksheet(email_date)
            except gspread.exceptions.WorksheetNotFound:
                ws = spreadsheet.add_worksheet(title=email_date, rows="100", cols="20")
                ws.append_row(["Time", "From", "Subject", "Details", "Attachment"])
            
            # Update the worksheet headers if necessary
            headers = ws.row_values(1)
            if "Attachment" not in headers:
                headers.append("Attachment")
                ws.delete_row(1)
                ws.insert_row(headers, 1)

            # Check if the email is already recorded
            records = ws.get_all_records()
            already_recorded = any(record['Subject'] == subject and record['Time'] == email_time for record in records)
            if already_recorded:
                continue
            
            # Flag to check if email has attachments
            has_attachment = False
            
            details = ""
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    try:
                        body = part.get_payload(decode=True).decode()
                    except:
                        pass
                    if "attachment" in content_disposition:
                        filename = part.get_filename()
                        if filename:
                            if not has_attachment:
                                # Create a new folder for this email's attachments
                                email_folder_id, email_folder_link = create_drive_folder(subject, drive_folder_id)
                                has_attachment = True
                            file_data = part.get_payload(decode=True)
                            file_name = f"{filename}"
                            upload_to_drive(file_data, file_name, email_folder_id)
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        details = body
            else:
                content_type = msg.get_content_type()
                body = msg.get_payload(decode=True).decode()
                if content_type == "text/plain":
                    details = body
            
            attachment_link = email_folder_link if has_attachment else "None"
            
            # Append the details to the worksheet
            ws.append_row([email_time, from_, subject, details, attachment_link])

# Close the connection and logout
mail.close()
mail.logout()

print("Email details and attachments uploaded to Google Drive and saved to Google Sheets")
