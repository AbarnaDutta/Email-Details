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
import logging
import re
# Set up logging
logging.basicConfig(filename='email_automation.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Account credentials
username = os.environ.get('EMAIL_USERNAME')
password = os.environ.get('EMAIL_PASSWORD')

# Google Sheets and Drive API setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(os.environ.get('CREDENTIALS_PATH'), scope)
client = gspread.authorize(creds)

# Initialize Google Drive API service
drive_creds = Credentials.from_service_account_file(os.environ.get('CREDENTIALS_PATH'), scopes=scope)
drive_service = build('drive', 'v3', credentials=drive_creds)

# Open the Google Sheets document
spreadsheet = client.open_by_url(os.environ.get('SPREADSHEET_URL'))

# Create an IMAP4 class with SSL
mail = imaplib.IMAP4_SSL("imap.gmail.com")

# Authenticate
try:
    mail.login(username, password)
except imaplib.IMAP4.error as e:
    logging.error(f"IMAP Authentication failed: {e}")
    raise

mail.select("inbox")
status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()

# Function to decode email subject
def decode_subject(subject):
    decoded, encoding = decode_header(subject)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(encoding if encoding else "utf-8")
    return decoded

# Function to decode email date
def decode_date(date_):
    # Remove any extraneous timezone information
    if 'GMT' in date_:
        date_ = date_.replace('GMT', '+0000')
    elif '(' in date_:
        date_ = date_.split('(')[0].strip()

    try:
        return datetime.strptime(date_, '%a, %d %b %Y %H:%M:%S %z')
    except ValueError:
        logging.error(f"Date parsing error: {date_}")
        return None

# Create a folder in Google Drive
def create_drive_folder(folder_name, parent_folder_id):
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id]
    }
    try:
        folder = drive_service.files().create(body=file_metadata, fields='id, webViewLink').execute()
        return folder.get('id'), folder.get('webViewLink')
    except Exception as e:
        logging.error(f"Failed to create folder: {e}")
        raise

# Upload a file to Google Drive
def upload_to_drive(file_data, file_name, folder_id):
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaIoBaseUpload(io.BytesIO(file_data), mimetype='application/octet-stream')
    try:
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        logging.info(f'File {file_name} uploaded to Google Drive with ID: {file.get("id")}')
    except Exception as e:
        logging.error(f"Failed to upload file {file_name}: {e}")
        raise

# Google Drive parent folder ID
drive_folder_id = os.environ.get('DRIVE_FOLDER_ID')

# Function to process each part of the email
def process_part(part):
    content_disposition = str(part.get("Content-Disposition"))
    has_attachment = False
    details = ""
    filename = None

    if "attachment" in content_disposition or part.get_filename():
        filename = part.get_filename()
        if filename:
            logging.info(f"Processing attachment: {filename}")
            file_data = part.get_payload(decode=True)
            has_attachment = True

    if part.get_content_type() == "text/plain" and not "attachment" in content_disposition:
        try:
            details = part.get_payload(decode=True).decode()
        except:
            pass

    return has_attachment, details, filename, part.get_payload(decode=True)

# Function to process each part of the email
def process_part(part):
    content_disposition = str(part.get("Content-Disposition"))
    content_type = part.get_content_type()
    content_transfer_encoding = part.get("Content-Transfer-Encoding", "")

    has_attachment = False
    details = ""
    filename = None
    file_data = None

    # Check if the part is an attachment
    if "attachment" in content_disposition or part.get_filename():
        filename = part.get_filename()
        if filename:
            print(f"Attachment found: {filename}")
            file_data = part.get_payload(decode=True)
            if file_data is not None:
                has_attachment = True
            else:
                logging.error(f"Failed to decode attachment: {filename}")

    # Handle text/plain and text/html parts separately
    if content_type in ["text/plain", "text/html"] and not "attachment" in content_disposition:
        try:
            payload = part.get_payload(decode=True)
            if payload is not None:
                details += payload.decode()
        except Exception as e:
            print(f"Failed to decode text part: {e}")

    # Log part details for debugging
    print(f"Content Type: {content_type}")
    print(f"Content-Disposition: {content_disposition}")
    print(f"Content Transfer Encoding: {content_transfer_encoding}")

    return has_attachment, details, filename, file_data if has_attachment else None

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

            # Check if the email is already recorded
            records = ws.get_all_records()
            already_recorded = any(record['Subject'] == subject and record['Time'] == email_time for record in records)
            if already_recorded:
                continue

            has_attachment = False
            details = ""
            email_folder_id, email_folder_link = None, None

            if msg.is_multipart():
                for part in msg.walk():
                    part_has_attachment, part_details, filename, file_data = process_part(part)
                    if part_has_attachment:
                        if not email_folder_id:
                            # Create a new folder for this email's attachments
                            email_folder_id, email_folder_link = create_drive_folder(subject, drive_folder_id)
                        upload_to_drive(file_data, filename, email_folder_id)
                        has_attachment = True
                    if part_details:
                        details += part_details

            else:
                has_attachment, details, filename, file_data = process_part(msg)
                if has_attachment:
                    # Create a new folder for this email's attachments
                    email_folder_id, email_folder_link = create_drive_folder(subject, drive_folder_id)
                    upload_to_drive(file_data, filename, email_folder_id)

            # Check for Google Drive links in email body
            attachment_link = "None"
            for part in msg.walk():
                if part.get_content_type() in ["text/plain", "text/html"]:
                    body = part.get_payload(decode=True)
                    if body:
                        body_str = body.decode()
                        links = re.findall(r'(https://drive\.google\.com[^\s]+)', body_str)
                        if links:
                            attachment_link = " | ".join(links)
                        else:
                            attachment_link = email_folder_link if has_attachment else "None"

            # Append the details to the worksheet
            ws.append_row([email_time, from_, subject, details, attachment_link])

# Close the connection and logout
mail.close()
mail.logout()

logging.info("Email details and attachments uploaded to Google Drive and saved to Google Sheets")
