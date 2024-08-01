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
import time
import re
import os

# Set up logging
logging.basicConfig(filename='email_automation.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Account credentials
username = os.getenv('EMAIL_USERNAME')
password = os.getenv('EMAIL_PASSWORD')

# Google Sheets and Drive API setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(os.getenv('CREDENTIALS_PATH'), scope)
client = gspread.authorize(creds)

# Initialize Google Drive API service
drive_creds = Credentials.from_service_account_file(os.getenv('CREDENTIALS_PATH'), scopes=scope)
drive_service = build('drive', 'v3', credentials=drive_creds)

# Open the Google Sheets document
spreadsheet = client.open_by_url(os.getenv('SPREADSHEET_URL'))

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

MAX_CELL_SIZE = 50000  # Maximum allowed characters in a Google Sheets cell

# Function to split large text into chunks
def split_text(text, max_size):
    return [text[i:i+max_size] for i in range(0, len(text), max_size)]

def open_sheet_with_retry(url, retries=5, delay=10):
    for attempt in range(retries):
        try:
            return client.open_by_url(url)
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:  # Rate limit error
                logging.warning(f"Rate limit exceeded. Retrying in {delay} seconds...")
                time.sleep(delay)
                delay *= 2  # Exponential backoff
            else:
                raise e
    raise Exception("Failed to open Google Sheets after multiple attempts")

# Function to decode email subject
def decode_subject(subject):
    decoded, encoding = decode_header(subject)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(encoding if encoding else "utf-8")
    return decoded

# Function to decode email date
def decode_date(date_):
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
    logging.info(f'File {file_name} uploaded to Google Drive with ID: {file.get("id")}')

# Extract file ID from Google Drive URL
def extract_file_id(drive_url):
    if 'drive.google.com' in drive_url:
        try:
            return drive_url.split('/d/')[1].split('/')[0]
        except IndexError:
            logging.error(f"Failed to extract file ID from URL: {drive_url}")
    return None

# Google Drive parent folder ID
drive_folder_id = os.getenv('DRIVE_FOLDER_ID')

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
            logging.info(f"Attachment found: {filename}")
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
            logging.error(f"Failed to decode text part: {e}")

    # Log part details for debugging
    logging.debug(f"Content Type: {content_type}")
    logging.debug(f"Content-Disposition: {content_disposition}")
    logging.debug(f"Content Transfer Encoding: {content_transfer_encoding}")

    return has_attachment, details, filename, file_data if has_attachment else None

def read_worksheet_with_retry(ws, retries=5, delay=10):
    for attempt in range(retries):
        try:
            return ws.get_all_records()
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:  # Rate limit error
                logging.warning(f"Rate limit exceeded while reading worksheet. Retrying in {delay} seconds...")
                time.sleep(delay)
                delay *= 2  # Exponential backoff
            else:
                raise e
    raise Exception("Failed to read worksheet after multiple attempts")

# Fetch and process each email
for email_id in email_ids:
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            subject = decode_subject(msg["Subject"])
            from_ = msg.get("From")
            date_ = msg.get("Date")
            decoded_date = decode_date(date_)
            if decoded_date:
                email_time = decoded_date.strftime("%H:%M:%S")
                email_date = decoded_date.strftime("%Y-%m-%d")
            else:
                logging.error(f"Date decoding failed for email ID {email_id}. Skipping email.")
                continue

            try:
                ws = spreadsheet.worksheet(email_date)
            except gspread.exceptions.WorksheetNotFound:
                ws = spreadsheet.add_worksheet(title=email_date, rows="100", cols="20")
                ws.append_row(["Time", "From", "Subject", "Details", "Attachment"])

            try:
                records = read_worksheet_with_retry(ws)
            except Exception as e:
                logging.error(f"Failed to read worksheet records: {e}")
                continue

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
                            email_folder_id, email_folder_link = create_drive_folder(subject, drive_folder_id)
                        upload_to_drive(file_data, filename, email_folder_id)
                        has_attachment = True
                    if part_details:
                        details += part_details
            else:
                has_attachment, details, filename, file_data = process_part(msg)
                if has_attachment:
                    email_folder_id, email_folder_link = create_drive_folder(subject, drive_folder_id)
                    upload_to_drive(file_data, filename, email_folder_id)

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

            # Handle long details text
            details_chunks = split_text(details, MAX_CELL_SIZE)
            for chunk in details_chunks:
                try:
                    ws.append_row([email_time, from_, subject, chunk, attachment_link])
                except gspread.exceptions.APIError as e:
                    logging.error(f"Failed to append row to worksheet: {e}")
                    time.sleep(10)  # Short delay before retrying

# Close the connection and logout
mail.close()
mail.logout()

logging.info("Email details and attachments uploaded to Google Drive and saved to Google Sheets")
