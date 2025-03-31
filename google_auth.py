import os
import json
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from config import Config

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

class GoogleAuth:
    def __init__(self, config: Config):
        self.config = config
        self._credentials = None
        self._sheets_service = None
        self._drive_service = None

    def get_credentials(self) -> Credentials:
        if self._credentials is None:
            self._credentials = self._load_or_refresh_credentials()
        return self._credentials

    def _load_or_refresh_credentials(self) -> Credentials:
        creds = None
        if os.path.exists(self.config.token_path):
            with open(self.config.token_path) as token:
                creds = Credentials.from_authorized_user_info(json.load(token), SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.config.client_secret_path, SCOPES)
                creds = flow.run_local_server(port=0)
            
            with open(self.config.token_path, 'w') as token:
                token.write(creds.to_json())

        return creds

    def get_service(self):
        if self._sheets_service is None:
            creds = self.get_credentials()
            self._sheets_service = build('sheets', 'v4', credentials=creds)
        return self._sheets_service

    def get_drive_service(self):
        if self._drive_service is None:
            creds = self.get_credentials()
            self._drive_service = build('drive', 'v3', credentials=creds)
        return self._drive_service 