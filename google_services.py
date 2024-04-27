import os
import pickle
from time import sleep

import gspread
from googleapiclient.http import MediaFileUpload
from email.mime.multipart import MIMEMultipart
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

from constants import INTERN_SHEET_KEY, SIGNATURE_TEMPLATE_PATH
from utils import *


class GoogleServices:
    SCOPES = [
        "https://www.googleapis.com/auth/gmail.compose",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    def __init__(self):
        self.creds = None
        self.gmail_service = None
        self.drive_service = None
        self.sheet_service = None
        self.gmail = None
        self.sender_name = None
        self.phone = None
        self.charge = None
        self.gc = None
        self.login()
        self.signature = self.get_signature()
        
    def login(self):
        self.google_auth()

        clear_display()
        while True:
            print("\n")
            print_center(f"Last Logged-in Account : {self.email}")
            print("\n")
            print_center("Do you want to use it (y/n) : ")
            while True:
                choice = getch()
                print(choice)
                if choice == "y" or choice == "n":
                    print("\nLoading...")
                    break
                sleep(0.1)
                
            if choice == "n":
                self.google_auth(True)
            if self.gmail in self.interns.keys():
                self.sender_name = self.interns[self.gmail]["Name"]
                self.phone = self.interns[self.gmail]["Phone"]
                self.charge = self.interns[self.gmail]["Charge"]
                self.image = self.interns[self.gmail]["Image"]
                self.calendly = self.interns[self.gmail]["Calendly"]
                break
            else:
                clear_display()
                print("\n\n\nInvalid Sender-Email")
                print("Use any of these Account :")
                for i in self.interns:
                    print(f"  {i}   ({self.interns[i]['Email']})")
                print()

    def google_auth(self, new=False):
        """This will help to create service for the object"""
        if new:
            self.creds = None
        else:
            try:
                # The file secret_token.pickle stores the user's access and refresh tokens
                if os.path.exists('secret_token.pickle'):
                    with open('secret_token.pickle', 'rb') as token:
                        self.creds = pickle.load(token)
            except Exception as e:
                pass

        # If there are no (valid) credentials available, let the user log in
        if not self.creds or not self.creds.valid:
            try:
                self.creds.refresh(Request())
            except:
                # Use the JSON file containing your OAuth2 credentials
                flow = InstalledAppFlow.from_client_config(get_client_secret(), self.SCOPES)
                self.creds = flow.run_local_server(port=0)

            # Save the credentials for the next run
            with open('secret_token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.gmail_service = build('gmail', 'v1', credentials=self.creds)
        self.drive_service = build('drive', 'v3', credentials=self.creds)
        self.sheet_service = build('sheets', 'v4', credentials=self.creds)
        profile = self.gmail_service.users().getProfile(userId='me').execute()
        self.gmail = profile['emailAddress']
        self.interns = self.get_interns_info()
        self.email = self.get_user_email()
        self.gc = gspread.authorize(self.creds)

    def get_interns_info(self):
        gc = gspread.authorize(self.creds)
        spreadsheet = gc.open_by_key(INTERN_SHEET_KEY) 
        worksheet = spreadsheet.get_worksheet(0)

        interns_data = worksheet.get_all_records()

        interns_dict = {}
        for intern in interns_data:
            interns_dict[intern['Gmail']] = {
                'Name': intern['Name'],
                'Phone': intern['Phone'],
                'Email': intern['Email'],
                'Charge': intern['Charge'],
                'Image':intern['Image'],
                'Calendly':intern['Calendly']
            }
        return interns_dict
    
    def get_user_email(self):
        try:
            email =  self.interns[self.gmail]["Email"]
            if email == "":
                return self.gmail
            return email
        except:
            return self.gmail
        
    def get_signature(self):
        with open(SIGNATURE_TEMPLATE_PATH, 'r', encoding="utf-8") as file:
            html_template = file.read()
        
        signature = html_template.format(
            image_link = self.image,
            name = self.sender_name,
            charge = self.charge,
            phone = self.phone,
            email = self.email,
            calendly_link = self.calendly
        )
        return signature
    
    def send_email(self, message: MIMEMultipart):
        try:
            status = self.gmail_service.users().messages().send(
                userId=self.gmail, body=message).execute()
            if status:
                print("\nEmail sent successfully.")
                return status
        except Exception as e:
            print(f"Error sending email: {e}")
            return None
    
    
    def print_details(self):
        clear_display()
        print("\n")
        print_center(
            f"  Account : {self.email}  ")
        print_center(
            f"  Sender : {self.sender_name}  ")
        print()
        
    def create_folder(self, name, parent_id):
        """Create a folder on Google Drive or use an existing one, and return its ID."""
        # Search for existing folder with the same name in the specified parent directory
        query = f"name='{name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        response = self.drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])

        if files:
            return files[0]['id']
        else:
            folder_metadata = {
                'name': name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_id]
            }
            folder = self.drive_service.files().create(body=folder_metadata, fields='id').execute()
            return folder.get('id')


    def upload_folder(self, folder_path, parent_folder_id):
        folder_name = os.path.basename(folder_path)
        new_folder_id = self.create_folder(folder_name, parent_folder_id)
        print(f"Folder --- {folder_name}")
        for item in os.listdir(folder_path):
            file_path = os.path.join(folder_path, item)
            if os.path.isfile(file_path):
                self.upload_file(file_path, new_folder_id)
    
    def upload_file(self, file_path, folder_id):
        """Upload a file to Google Drive within the specified folder, skip if file already exists."""
        file_name = os.path.basename(file_path)
        # Search for existing file with the same name in the specified folder
        query = f"name='{file_name}' and '{folder_id}' in parents and trashed=false"
        response = self.drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])

        if not files:
            # No existing file, proceed with upload
            file_metadata = {
                'name': file_name,
                'parents': [folder_id]
            }
            media = MediaFileUpload(file_path, resumable=True)
            file = self.drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            print(f"Upload          |- {file_name}")
        else:
            # File already exists, skip the upload
            print(f"Skip            |- {file_name}")
            
    def delete_file_by_name(self, file_name, folder_id):
        """Delete a file by name within the specified folder."""
        query = f"name='{file_name}' and '{folder_id}' in parents and trashed=false"
        response = self.drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = response.get('files', [])

        if files:
            for file in files:
                # Perform the deletion
                self.drive_service.files().delete(fileId=file['id']).execute()
                print(f"Deleted         |- {file_name}")
        else:
            print(f"File not found  |- {file_name}")
