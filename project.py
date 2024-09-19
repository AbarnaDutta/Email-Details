import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.service_account import Credentials
import logging
import time
import re
import random
import os
import json
from pathlib import Path
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
import sys
from azure.core.exceptions import HttpResponseError


sys.stdout.reconfigure(encoding='utf-8')

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
spreadsheet_url = os.getenv('SPREADSHEET_URL')
spreadsheet = client.open_by_url(spreadsheet_url)

# Google Drive parent folder ID
parent_folder_id = os.getenv('DRIVE_FOLDER_ID')

# Create a cache for worksheets
worksheet_cache = {ws.title: ws for ws in spreadsheet.worksheets()}

def get_or_create_worksheet(sheet_name):
    if sheet_name in worksheet_cache:
        return worksheet_cache[sheet_name]
    else:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
        ws.append_row(["Date", "Time", "From", "Subject", "Invoice Number","Invoice Date","Invoice Amount", "Vendor Name", "Attachment"])
        worksheet_cache[sheet_name] = ws
        return ws

def retry_api_call(func, *args, retries=5, delay=10, backoff=2):
    for attempt in range(retries):
        try:
            return func(*args)
        except gspread.exceptions.APIError as e:
            if e.response.status_code == 429:  # Rate limit error
                logging.warning(f"Rate limit exceeded. Retrying in {delay} seconds...")
                time.sleep(delay)
                delay *= backoff  # Exponential backoff
                delay += random.uniform(0, 1)  # Add jitter
            else:
                raise e
    raise Exception("Failed to complete API call after multiple attempts")

def decode_subject(subject):
    decoded, encoding = decode_header(subject)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(encoding if encoding else "utf-8")
    return decoded

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

def create_drive_folder(folder_name, parent_folder_id):
    file_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id]
    }
    folder = drive_service.files().create(body=file_metadata, fields='id, webViewLink').execute()
    return folder.get('id'), folder.get('webViewLink')

def upload_to_drive(file_data, file_name, folder_id):
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaIoBaseUpload(io.BytesIO(file_data), mimetype='application/octet-stream')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f'File {file_name} uploaded to Google Drive with ID: {file.get("id")}')

def extract_file_id(drive_url):
    if 'drive.google.com' in drive_url:
        try:
            return drive_url.split('/d/')[1].split('/')[0]
        except IndexError:
            logging.error(f"Failed to extract file ID from URL: {drive_url}")
    return None

def get_or_create_monthly_folder(year_month):
    query = f"name='{year_month}' and mimeType='application/vnd.google-apps.folder' and '{parent_folder_id}' in parents"
    results = drive_service.files().list(q=query, fields="files(id)").execute()
    folders = results.get('files', [])
    if folders:
        return folders[0]['id']

    file_metadata = {
        'name': year_month,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_folder_id]
    }
    folder = drive_service.files().create(body=file_metadata, fields='id').execute()
    return folder.get('id')

def process_part(part):
    content_disposition = str(part.get("Content-Disposition", ""))
    content_type = part.get_content_type()
    content_transfer_encoding = part.get("Content-Transfer-Encoding", "")

    has_attachment = False
    filename = None
    file_data = None

    if "attachment" in content_disposition or part.get_filename():
        filename = part.get_filename()
        if filename:
            print(f"Attachment found: {filename}")
            file_data = part.get_payload(decode=True)
            if file_data is not None:
                has_attachment = True
            else:
                logging.error(f"Failed to decode attachment: {filename}")

    if content_type in ["text/plain", "text/html"] and not "attachment" in content_disposition:
        try:
            payload = part.get_payload(decode=True)
            if payload is not None:
                pass
        except Exception as e:
            print(f"Failed to decode text part: {e}")

    print(f"Content Type: {content_type}")
    print(f"Content-Disposition: {content_disposition}")
    print(f"Content Transfer Encoding: {content_transfer_encoding}")

    return has_attachment, filename, file_data


class APIRateLimiter:
    def __init__(self, min_interval_seconds):
        self.min_interval_seconds = min_interval_seconds
        self.last_call_time = None

    def wait_if_needed(self):
        if self.last_call_time:
            elapsed_time = (datetime.now() - self.last_call_time).total_seconds()
            wait_time = self.min_interval_seconds - elapsed_time
            if wait_time > 0:
                time.sleep(wait_time)

    def record_call(self):
        self.last_call_time = datetime.now()

# Rate Limiter Setup: 3 seconds between calls
recognizer_rate_limiter = APIRateLimiter(min_interval_seconds=3)

class DocumentExtractor:
    def __init__(self, endpoint: str, key: str, invoice_model: str, receipt_model: str):
        self.document_analysis_client = DocumentAnalysisClient(
            endpoint=endpoint, credential=AzureKeyCredential(key)
        )
        self.invoice_model_id = invoice_model
        self.receipt_model_id = receipt_model

    def extract_document_data(self, file_path: Path) -> dict:
        document_data = {
            "invoice_number": None,
            "invoice_date": None,
            "invoice_amount": None,
            "vendor_name": None,
        }

        # Try extracting with the invoice model first
        invoice_data = self.extract_data_with_model(file_path, self.invoice_model_id)
        
        # If any field is missing, attempt to extract from the receipt model
        receipt_data = self.extract_data_with_model(file_path, self.receipt_model_id)

        # Use receipt data to fill in missing fields from the invoice data
        for field in document_data.keys():
            if not invoice_data.get(field):
                document_data[field] = receipt_data.get(field)
            else:
                document_data[field] = invoice_data[field]

        return document_data

    def extract_data_with_model(self, file_path: Path, model_id: str) -> dict:
        extracted_data = {
            "invoice_number": None,
            "invoice_date": None,
            "invoice_amount": None,
            "vendor_name": None,
        }

        max_retries = 3
        retry_delay = 60  # Initial delay of 60 seconds for backoff

        for attempt in range(max_retries):
            try:
                recognizer_rate_limiter.wait_if_needed()
                
                with open(file_path, "rb") as f:
                    poller = self.document_analysis_client.begin_analyze_document(
                        model_id, document=f, locale="en-US"
                    )
                recognizer_rate_limiter.record_call()
                documents = poller.result()

                # Additional logic for extracting invoice number with the invoice model
                with open(file_path, "rb") as f:
                    poller = self.document_analysis_client.begin_analyze_document(
                        self.invoice_model_id, document=f, locale="en-US"
                    )
                invoice_documents = poller.result()
                    
                for document in invoice_documents.documents:
                    extracted_data["invoice_number"] = (
                        document.fields.get("InvoiceId").value if document.fields.get("InvoiceId") else None
                    )

                # Extract fields specific to the model used
                for document in documents.documents:
                    if model_id == self.invoice_model_id:
                        extracted_data["invoice_number"] = (
                            document.fields.get("InvoiceId").value if document.fields.get("InvoiceId") else None
                        )
                        extracted_data["invoice_date"] = (
                            document.fields.get("InvoiceDate").value.strftime("%Y-%m-%d") if document.fields.get("InvoiceDate") and document.fields.get("InvoiceDate").value else None
                        )
                        if document.fields.get("InvoiceTotal"):
                            invoice_total_text = document.fields.get("InvoiceTotal").content
                            amount = document.fields.get("InvoiceTotal").value.amount
                            currency_symbol = self.get_currency_symbol(invoice_total_text)
                            extracted_data["invoice_amount"] = f"{currency_symbol}{amount}"
    
                        extracted_data["vendor_name"] = (
                            document.fields.get("VendorName").value if document.fields.get("VendorName") else None
                        )
    
                    elif model_id == self.receipt_model_id:
                        extracted_data["invoice_date"] = (
                            document.fields.get("TransactionDate").value.strftime("%Y-%m-%d") if document.fields.get("TransactionDate") and document.fields.get("TransactionDate").value else None
                        )
                        if document.fields.get("Total"):
                            total_text = document.fields.get("Total").content
                            amount = document.fields.get("Total").value
                            currency_symbol = self.get_currency_symbol(total_text)
                            extracted_data["invoice_amount"] = f"{currency_symbol}{amount}"
    
                        extracted_data["vendor_name"] = (
                            document.fields.get("MerchantName").value if document.fields.get("MerchantName") else None
                        )
                return extracted_data

            except HttpResponseError as e:
                if e.status_code == 403:
                    logging.error(f"Quota exceeded. Retrying in {retry_delay} seconds... (Attempt {attempt + 1}/{max_retries})")
                    time.sleep(retry_delay)
                    retry_delay *= 2  # Exponential backoff
                else:
                    raise e

        raise Exception("Quota exceeded and retry attempts failed.")



    def get_currency_symbol(self, text: str) -> str:
        """
        Extract the currency symbol from the text and convert any Unicode escape sequences.
        """
        # Regex to match common currency symbols or Unicode sequences just before the amount
        currency_symbols_pattern = r"([€£$¥₹])|\\u([0-9a-fA-F]{4})"
        
        # Search for a currency symbol or Unicode sequence in the text
        match = re.search(currency_symbols_pattern, text)
        
        if match:
            if match.group(1):
                return match.group(1)  # Directly return the symbol (€, £, $, etc.)
            elif match.group(2):
                # Convert the Unicode sequence to an actual character
                return chr(int(match.group(2), 16))
        return ""  # Return empty if no symbol is found


# Initialize DocumentExtractor with Azure OCR credentials and model IDs
document_extractor = DocumentExtractor(
    endpoint=os.getenv("AZURE_ENDPOINT"),
    key=os.getenv("AZURE_KEY"),
    invoice_model="prebuilt-invoice",
    receipt_model="prebuilt-receipt"
)


def update_total_invoice_amount(ws):
    # Get all values to search for "Total Amount" row and determine the number of rows
    all_values = ws.get_all_values()
    
    # Delete the "Total Amount" row if it exists, before processing any records
    for i in range(len(all_values)):
        if all_values[i][0] == "Total Amount":
            ws.delete_rows(i + 1)  # Row numbers are 1-indexed
            break  # Delete only the first occurrence and stop
            
    records = ws.get_all_records()
    print("All records:", records)

    # Initialize a dictionary to hold totals for different currencies
    currency_totals = {}
    no_currency_total = 0.0

    # Start from row 2 (skip header)
    for record in records[1:]:
        invoice_amount = str(record["Invoice Amount"])

        # Debug: Print the raw invoice amount
        print(f"Processing invoice amount: {invoice_amount}")

        # Detect currency symbol and extract the amount
        match = re.match(r'([€£$₹]?)([\d,\.]+)', invoice_amount)
        if match:
            currency_symbol = match.group(1) if match.group(1) else ''  # Default to '' if no symbol
            amount_str = match.group(2).replace(",", "")

            try:
                amount = float(amount_str)

                # Debug: Print extracted amount and symbol
                print(f"Extracted amount: {amount} with symbol: '{currency_symbol}'")

                if currency_symbol == '':
                    # If no currency symbol, treat as separate 'No Currency' amount
                    no_currency_total += amount
                else:
                    # Accumulate totals for each currency symbol
                    if currency_symbol in currency_totals:
                        currency_totals[currency_symbol] += amount
                    else:
                        currency_totals[currency_symbol] = amount
            except ValueError:
                print(f"Skipping invalid amount: {amount_str}")
                continue  # Skip invalid amounts

    # Debug: Print accumulated totals for verification
    print("Accumulated currency totals:", currency_totals)
    print("No currency total:", no_currency_total)

    # Prepare the total amount text for all currencies
    currency_totals_text = " + ".join([f"{symbol}{total:.2f}" for symbol, total in currency_totals.items()])
    
    # Add no currency total if it exists
    if no_currency_total > 0:
        currency_totals_text += f" + {no_currency_total:.2f}"

    # Debug: Print final total amount text
    print(f"Final total amount text: {currency_totals_text}")

    # Append new "Total Amount" row with placeholders for G, H, I
    ws.append_row(["Total Amount", "", "", "", "", "", currency_totals_text, "", ""])

    # Get the new last row number
    new_last_row = len(ws.get_all_values())  # Update after appending

    # Merge the first six columns (A:F) and the last three columns (G:I)
    ws.merge_cells(new_last_row, 1, new_last_row, 6)  # Merge A:F
    ws.merge_cells(new_last_row, 7, new_last_row, 9)  # Merge G:I

    print(f"Total invoice amount updated: {currency_totals_text}")










def process_email_attachment(email_date, email_time, from_, subject, part, extracted_data):
    # Normalize the extracted data for comparison
    normalized_extracted_data = json.dumps(
        extracted_data, ensure_ascii=False, indent=None, separators=(',', ':')
    )
    normalized_extracted_data = json.loads(normalized_extracted_data)

    # Convert extracted data fields to strings
    invoice_number = str(normalized_extracted_data.get('invoice_number', ''))
    invoice_date = str(normalized_extracted_data.get('invoice_date', ''))
    invoice_amount = str(normalized_extracted_data.get('invoice_amount', ''))
    vendor_name = str(normalized_extracted_data.get('vendor_name', ''))
    
    if not invoice_date:
        print("No invoice date found in the attachment.")
        return

    # Determine year-month from the invoice date
    year_month = datetime.strptime(invoice_date, "%Y-%m-%d").strftime("%Y-%m")

    # Get or create the correct worksheet
    ws = get_or_create_worksheet(year_month)

    # Retrieve all records from the worksheet
    records = ws.get_all_records()
    

    # Flag to determine if a match was found
    match_found = False

    # Iterate through existing records to find a match
    for record in records:
        # Convert each field from the Google Sheet's record to string
        normalized_record = {
            'Invoice Number': str(record.get('Invoice Number', '')).strip(),
            'Invoice Date': str(record.get('Invoice Date', '')).strip(),
            'Invoice Amount': str(record.get('Invoice Amount', '')).strip(),
            'Vendor Name': str(record.get('Vendor Name', '')).strip()
        }


        # Print comparisons with string values
        print(f"Comparing Invoice Number: {normalized_record['Invoice Number']} vs {invoice_number}")
        print(f"Comparing Invoice Date: {normalized_record['Invoice Date']} vs {invoice_date}")
        print(f"Comparing Invoice Amount: {normalized_record['Invoice Amount']} vs {invoice_amount}")
        print(f"Comparing Vendor Name: {normalized_record['Vendor Name']} vs {vendor_name}")

        if (
            normalized_record['Invoice Number'] == invoice_number and
            normalized_record['Invoice Date'] == invoice_date and
            normalized_record['Invoice Amount'] == invoice_amount and
            normalized_record['Vendor Name'] == vendor_name
        ):
            print("Match found with the existing record.")
            match_found = True
            break
    # If no exact match is found, update the record
    if not match_found:
        print("No exact match found, appending new record to the sheet.")
        # Get or create the corresponding month folder in Google Drive
        month_folder_id = get_or_create_monthly_folder(year_month)

        # Create a folder for the email subject within the month folder
        email_folder_id, email_folder_link = create_drive_folder(subject, month_folder_id)

        # Upload the attachment to the Drive folder
        filename = part.get_filename()
        file_data = part.get_payload(decode=True)
        upload_to_drive(file_data, filename, email_folder_id)


        # Update the Google Sheet
        ws.append_row([email_date, email_time, from_, subject, invoice_number, invoice_date, invoice_amount, vendor_name, email_folder_link])
        # Update the total invoice amount at the bottom of the sheet
        update_total_invoice_amount(ws)
        
        print(f"Record updated for email from {from_} with subject {subject}.")
    else:
        print(f"Email from {from_} with subject {subject} already exists with the same details.")



    
# Fetch and process each email
mail = imaplib.IMAP4_SSL("imap.gmail.com")
mail.login(username, password)
mail.select("inbox")
status, messages = mail.search(None, "UNSEEN")
email_ids = messages[0].split()

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

            if msg.is_multipart():
                for part in msg.walk():
                    part_has_attachment, filename, file_data = process_part(part)
                    if part_has_attachment:
                        temp_path = Path(f"temp_{filename}")
                        with open(temp_path, "wb") as temp_file:
                            temp_file.write(file_data)
                        extracted_data = document_extractor.extract_document_data(temp_path)
                        temp_path.unlink()  # Remove temporary file
                        process_email_attachment(email_date, email_time, from_, subject, part, extracted_data)

            else:
                has_attachment, filename, file_data = process_part(msg)
                if has_attachment:
                    temp_path = Path(f"temp_{filename}")
                    with open(temp_path, "wb") as temp_file:
                        temp_file.write(file_data)
                    extracted_data = document_extractor.extract_document_data(temp_path)
                    temp_path.unlink()

                    process_email_attachment(email_date, email_time, from_, subject, msg, extracted_data)
    # Mark the email as seen
    mail.store(email_id, '+FLAGS', '\\Seen')

# Close the connection and logout
mail.close()
mail.logout()

print("Email details and attachments uploaded to Google Drive and recorded in Google Sheets.")
