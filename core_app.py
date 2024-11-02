import win32com.client
import re
import os
import tempfile
import pandas as pd
from PyPDF2 import PdfReader
import logging
import easyocr
import shutil
import numpy as np
import pdfplumber
import pythoncom
import time
import threading

class GUIHandler(logging.Handler):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        log_entry = self.format(record)
        self.callback(log_entry)

class OutlookProcessor:
    required_columns = ['Mailbox', 'Folder_destination', 'Filter_1', 'Filter_2', 'Filter_3', 'Filter_4', 'Filter_5']

    def __init__(self, excel_path, log_callback=None):
        self.excel_path = excel_path
        self.log_callback = log_callback

        # Setup logging
        self.logger = logging.getLogger('OutlookProcessor')
        self.logger.setLevel(logging.INFO)
        file_handler = logging.FileHandler('outlook_process_log.log')
        file_handler.setFormatter(logging.Formatter('%(asctime)s:%(levelname)s:%(message)s'))
        self.logger.addHandler(file_handler)
        
        if self.log_callback:
            gui_handler = GUIHandler(self.log_callback)
            gui_handler.setFormatter(logging.Formatter('%(asctime)s:%(levelname)s:%(message)s'))
            self.logger.addHandler(gui_handler)
        
    def load_configuration(self):
        try:
            self.config = pd.read_excel(self.excel_path)
            self.logger.info("Configuration loaded successfully.")
            self.validate_configuration()
        except Exception as e:
            self.logger.error(f"Failed to load configuration from {self.excel_path}: {e}")
            self.config = None
    
    def validate_configuration(self):
        if not all(col in self.config.columns for col in self.required_columns):
            self.logger.error("Some required columns are missing in the configuration file.")
            self.config = None
    
    def connect_to_outlook(self):
        pythoncom.CoInitialize()  # Initialize COM library
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            self.logger.info("Successfully connected to Outlook.")
        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {e}. It may be necessary to close Outlook and try again.")
    
    def disconnect_from_outlook(self):
        """Uninitialize COM library to clean up resources."""
        pythoncom.CoUninitialize()
        self.logger.info("Disconnected from Outlook.")

    def extract_text_with_pdf_reader(self, file_path):
        text = ""
        try:
            with open(file_path, 'rb') as f:
                pdf = PdfReader(f)
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text
            self.logger.info(f"Text extracted from {os.path.basename(file_path)} using PdfReader.")
        except Exception as e:
            self.logger.error(f"Error reading PDF file {file_path}: {e}")
            return None
        return text

    def extract_text_with_easyocr(self, file_path):
        text = ""
        try:
            abs_path = os.path.abspath(file_path)
            if not os.path.exists(abs_path):
                self.logger.error(f"File does not exist: {abs_path}")
                return None

            reader = easyocr.Reader(['en'])

            with pdfplumber.open(abs_path) as pdf:
                for page in pdf.pages:
                    im = page.to_image(resolution=300)
                    pil_images = im.original

                    results = reader.readtext(np.array(pil_images), detail=0)
                    text += " ".join(results) + " "

            self.logger.info(f"Text extracted from {os.path.basename(file_path)} using easyocr.")
        except Exception as e:
            self.logger.error(f"Error with OCR on file {abs_path}: {e}")
            return None

        return text.strip()

    def handle_extracted_text(self, text, message, config_rows, inbox):
        patterns_matched = []
        for _, config_row in config_rows.iterrows():
            filters = [config_row[col] for col in self.required_columns[2:] if not pd.isna(config_row[col])]
            self.logger.info(f"Checking extracted text from email '{message.Subject}' against filters.")
            match = all(re.search(pattern, text, re.IGNORECASE) for pattern in filters)
            if match:
                destination_path = config_row['Folder_destination']
                target_folder = self.find_target_folder(inbox, destination_path)
                if target_folder:
                    patterns_matched.append((message, target_folder))
                    self.logger.info(f"Email with subject '{message.Subject}' matched all filters and will be moved to {destination_path}.")
                    break
        if not patterns_matched:
            self.logger.info(f"No matching patterns found for the email with subject '{message.Subject}'.")
        return patterns_matched

    def find_target_folder(self, base_folder, path):
        current_folder = base_folder
        try:
            for folder_name in path.split('/'):
                current_folder = current_folder.Folders[folder_name]
            self.logger.info(f"Successfully navigated to folder: {path}")
        except Exception as e:
            self.logger.error(f"Failed to navigate to {path}: {e}")
            return None
        return current_folder

    def get_mailbox(self, mailbox_name):
        if mailbox_name.lower() == 'main':
            return self.outlook.GetDefaultFolder(6)  # Inbox for main mailbox
        else:
            for store in self.outlook.Folders:
                if store.Name.lower() == mailbox_name.lower():
                    self.logger.info(f"Accessing shared mailbox '{mailbox_name}'.")
                    return store
            try:
                recipient = self.outlook.CreateRecipient(mailbox_name)
                recipient.Resolve()
                if recipient.Resolved:
                    shared_folder = self.outlook.GetSharedDefaultFolder(recipient, 6)  # Inbox for shared mailbox
                    if shared_folder:
                        self.logger.info(f"Accessed shared mailbox '{mailbox_name}'.")
                        return shared_folder
            except Exception as e:
                self.logger.error(f"Failed to access shared mailbox '{mailbox_name}': {e}")
        self.logger.error(f"Mailbox '{mailbox_name}' not found.")
        return None

    def list_all_folders(self, folder, indent=0):
        folder_list = []
        try:
            for subfolder in folder.Folders:
                folder_path = f"{' ' * indent}{subfolder.Name}"
                folder_list.append(folder_path)
                self.logger.info(f"Found folder: {folder_path}")
                folder_list.extend(self.list_all_folders(subfolder, indent + 2))
        except Exception as e:
            self.logger.error(f"Failed to list folders in '{folder.Name}': {e}")
        return folder_list

    def list_folders_in_sharebox(self, mailbox_name):
        mailbox = self.get_mailbox(mailbox_name)
        if mailbox is not None:
            self.logger.info(f"Listing folders in mailbox '{mailbox_name}'")
            all_folders = self.list_all_folders(mailbox)
            self.logger.info(f"Available folders in mailbox '{mailbox_name}':\n" + "\n".join(all_folders))
            return all_folders
        else:
            self.logger.error(f"Mailbox '{mailbox_name}' not found.")
            return []

    def process_emails(self):
        if self.config is None:
            self.logger.error("Invalid configuration. Exiting process.")
            return

        try:
            self.connect_to_outlook()
            
            email_count = 0
            pdfs_for_ocr = []
            temp_dir = tempfile.mkdtemp()
            move_operations = []

            try:
                mailboxes = self.config['Mailbox'].unique()
                for mailbox_name in mailboxes:
                    inbox = self.get_mailbox(mailbox_name)
                    if inbox is None:
                        self.logger.error(f"Skipping mailbox '{mailbox_name}' as it was not found.")
                        continue

                    available_folders = self.list_all_folders(inbox)
                    self.logger.info(f"Available folders in mailbox '{mailbox_name}':\n" + "\n".join(available_folders))

                    if mailbox_name.lower() != 'main':
                        folder_names = ["Posteingang", "Boîte de réception", "Inbox"]
                        for name in folder_names:
                            try:
                                inbox = inbox.Folders[name]
                                self.logger.info(f"Navigated to '{name}' folder in mailbox '{mailbox_name}'.")
                                break
                            except Exception as e:
                                self.logger.error(f"Failed to navigate to '{name}' folder in mailbox '{mailbox_name}': {e}")
                                continue
                        else:
                            self.logger.error(f"No recognized inbox folder found in mailbox '{mailbox_name}'.")
                            continue

                    self.logger.info(f"Processing mailbox: {mailbox_name}")

                    mailbox_config = self.config[self.config['Mailbox'] == mailbox_name]

                    for message in inbox.Items:
                        for attachment in message.Attachments:
                            if attachment.FileName.lower().endswith('.pdf'):
                                temp_file = os.path.join(temp_dir, attachment.FileName)
                                attachment.SaveAsFile(temp_file)
                                text = self.extract_text_with_pdf_reader(temp_file)
                                if not text:
                                    # If text extraction fails, mark for OCR processing
                                    pdfs_for_ocr.append((temp_file, message, mailbox_config, inbox))
                                    self.logger.info(f"Text extraction failed for {attachment.FileName}, scheduled for OCR.")
                                else:
                                    self.logger.info(f"Text extracted successfully from {attachment.FileName} using PdfReader.")
                                    patterns_matched = self.handle_extracted_text(text, message, mailbox_config, inbox)
                                    if patterns_matched:
                                        move_operations.extend(patterns_matched)
                                        email_count += 1
                                        break

                # Perform OCR on marked PDFs
                for temp_file, message, mailbox_config, inbox in pdfs_for_ocr:
                    if os.path.exists(temp_file):
                        text = self.extract_text_with_easyocr(temp_file)
                        if not text:
                            self.logger.error(f"OCR failed for {os.path.basename(temp_file)}.")
                            continue

                        self.logger.info(f"Text extracted successfully from {os.path.basename(temp_file)} using easyocr.")
                        patterns_matched = self.handle_extracted_text(text, message, mailbox_config, inbox)
                        if patterns_matched:
                            move_operations.extend(patterns_matched)
                            email_count += 1

                    if os.path.exists(temp_file):
                        os.remove(temp_file)

                for message, target_folder in move_operations:
                    try:
                        message.Move(target_folder)
                        self.logger.info(f"Email with subject '{message.Subject}' moved to {target_folder.Name}")
                    except Exception as e:
                        self.logger.error(f"Failed to move email with subject '{message.Subject}' to {target_folder.Name}: {e}")

            finally:
                shutil.rmtree(temp_dir)

        finally:
            self.disconnect_from_outlook()

        self.logger.info(f"Processed {email_count} emails.")

    def execute_once(self):
        self.process_emails()

    def execute_periodically(self, interval=600):
        self.stop_event = threading.Event()  # Create a stop event to allow stopping the loop
        try:
            while not self.stop_event.is_set():
                self.process_emails()  # Execute the email processing
                self.logger.info("Processing completed. Starting countdown for the next execution.")

                countdown = interval
                while countdown > 0 and not self.stop_event.is_set():
                    time.sleep(1)
                    countdown -= 1
                    if countdown % 120 == 0:  # Log every 2 minutes (120 seconds)
                        minutes_left = countdown // 60
                        self.logger.info(f"{minutes_left} minutes left until the next execution.")
        except Exception as e:
            self.logger.error(f"An error occurred during periodic execution: {e}")
        finally:
            self.logger.info("Periodic execution has been stopped.")
