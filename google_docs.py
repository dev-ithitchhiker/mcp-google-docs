from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from typing import List, Dict, Any, Optional
from google_auth import GoogleAuth
import json

class GoogleDocs:
    def __init__(self, auth: GoogleAuth):
        self.service = build('docs', 'v1', credentials=auth.get_credentials())
        
    def create_document(self, title: str) -> str:
        """Create a new Google Doc."""
        try:
            document = {
                'title': title
            }
            document = self.service.documents().create(body=document).execute()
            return document.get('documentId')
        except HttpError as error:
            print(f'An error occurred: {error}')
            return None

    def insert_text(self, document_id: str, text: str, index: int = 1) -> bool:
        """Insert text at a specific index in the document."""
        try:
            requests = [
                {
                    'insertText': {
                        'location': {
                            'index': index
                        },
                        'text': text
                    }
                }
            ]
            
            body = {
                'requests': requests
            }
            
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def insert_heading(self, document_id: str, text: str, level: int = 1, index: int = 1) -> bool:
        """Insert a heading at a specific index in the document."""
        try:
            requests = [
                {
                    'insertText': {
                        'location': {
                            'index': index
                        },
                        'text': text
                    }
                },
                {
                    'updateParagraphStyle': {
                        'range': {
                            'startIndex': index,
                            'endIndex': index + len(text)
                        },
                        'paragraphStyle': {
                            'namedStyleType': f'HEADING_{level}'
                        },
                        'fields': 'namedStyleType'
                    }
                }
            ]
            
            body = {
                'requests': requests
            }
            
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def insert_image(self, document_id: str, image_url: str, index: int = 1) -> bool:
        """Insert an image at a specific index in the document."""
        try:
            requests = [
                {
                    'insertInlineImage': {
                        'location': {
                            'index': index
                        },
                        'uri': image_url,
                        'objectSize': {
                            'predefinedSize': 'MEDIUM'
                        }
                    }
                }
            ]
            
            body = {
                'requests': requests
            }
            
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def get_document(self, document_id: str) -> Optional[Dict[str, Any]]:
        """Get document details."""
        try:
            document = self.service.documents().get(
                documentId=document_id
            ).execute()
            return document
        except HttpError as error:
            print(f'An error occurred: {error}')
            return None

    def delete_document(self, document_id: str) -> bool:
        """Delete a document."""
        try:
            self.service.documents().delete(
                documentId=document_id
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def create_table(self, document_id: str, rows: int, columns: int, index: int = 1) -> bool:
        """Create a table in the document."""
        try:
            requests = [
                {
                    'insertTable': {
                        'location': {
                            'index': index
                        },
                        'rows': rows,
                        'columns': columns,
                        'endOfSegmentLocation': True
                    }
                }
            ]
            
            body = {
                'requests': requests
            }
            
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False 