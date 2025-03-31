from typing import List, Dict, Any
from googleapiclient.discovery import build
from config import Config
from google_auth import GoogleAuth
import logging

logger = logging.getLogger(__name__)

class GoogleDrive:
    def __init__(self, auth: GoogleAuth):
        self.auth = auth
        self.service = auth.get_service()
        self.sheets_service = build('sheets', 'v4', credentials=auth.get_credentials())
        self.folder_id = auth.config.folder_id

    def list_files(self) -> List[Dict[str, Any]]:
        """List files in the specified folder."""
        try:
            results = self.service.files().list(
                q=f"'{self.folder_id}' in parents and trashed=false",
                fields="files(id, name, mimeType, createdTime, modifiedTime)",
                orderBy="modifiedTime desc"
            ).execute()
            files = results.get('files', [])
            
            if not files:
                logger.info("Folder is empty")
                return [{"message": "Folder is empty"}]
                
            logger.info(f"Found {len(files)} files in folder")
            return files
        except Exception as e:
            logger.error(f"Error listing files: {str(e)}", exc_info=True)
            return []

    def copy_file(self, file_id: str, new_name: str) -> Dict[str, Any]:
        """Copy a file."""
        try:
            file_metadata = {'name': new_name}
            result = self.service.files().copy(
                fileId=file_id,
                body=file_metadata
            ).execute()
            logger.info(f"Successfully copied file {file_id} to {new_name}")
            return result
        except Exception as e:
            logger.error(f"Error copying file {file_id}: {str(e)}", exc_info=True)
            return {}

    def rename_file(self, file_id: str, new_name: str) -> Dict[str, Any]:
        """Rename a file."""
        try:
            file_metadata = {'name': new_name}
            result = self.service.files().update(
                fileId=file_id,
                body=file_metadata
            ).execute()
            logger.info(f"Successfully renamed file {file_id} to {new_name}")
            return result
        except Exception as e:
            logger.error(f"Error renaming file {file_id}: {str(e)}", exc_info=True)
            return {}

    def create_spreadsheet(self, title: str) -> Dict[str, Any]:
        """Create an empty spreadsheet."""
        try:
            logger.info(f"Creating new spreadsheet with title: {title}")
            # Create spreadsheet
            spreadsheet = {
                'properties': {
                    'title': title
                },
                'sheets': [
                    {
                        'properties': {
                            'title': 'Sheet1'
                        }
                    }
                ]
            }
            
            # Create spreadsheet
            spreadsheet = self.sheets_service.spreadsheets().create(
                body=spreadsheet,
                fields='spreadsheetId'
            ).execute()
            
            spreadsheet_id = spreadsheet['spreadsheetId']
            logger.info(f"Created spreadsheet with ID: {spreadsheet_id}")
            
            try:
                # Get current parent folder information of the created spreadsheet
                file = self.service.files().get(
                    fileId=spreadsheet_id,
                    fields='parents'
                ).execute()
                
                # Remove previous parent folders and move to new folder
                previous_parents = ",".join(file.get('parents', []))
                
                if previous_parents:
                    # Move file to new folder
                    self.service.files().update(
                        fileId=spreadsheet_id,
                        addParents=self.folder_id,
                        removeParents=previous_parents,
                        fields='id, parents'
                    ).execute()
                    logger.info(f"Moved spreadsheet {spreadsheet_id} to folder {self.folder_id}")
                else:
                    # Add to new folder if no parent folder exists
                    self.service.files().update(
                        fileId=spreadsheet_id,
                        addParents=self.folder_id,
                        fields='id, parents'
                    ).execute()
                    logger.info(f"Added spreadsheet {spreadsheet_id} to folder {self.folder_id}")
            except Exception as e:
                logger.warning(f"Failed to move spreadsheet to folder: {str(e)}")
                # Continue even if folder move fails since spreadsheet is created
            
            result = {
                'spreadsheetId': spreadsheet_id,
                'title': title
            }
            logger.info(f"Successfully created spreadsheet: {result}")
            return result
        except Exception as e:
            logger.error(f"Error creating spreadsheet '{title}': {str(e)}", exc_info=True)
            return {}

    def create_spreadsheet_from_template(self, template_id: str, title: str) -> Dict[str, Any]:
        """Create a new spreadsheet from a template."""
        try:
            logger.info(f"Creating new spreadsheet from template {template_id} with title: {title}")
            # Copy template file
            file_metadata = {
                'name': title,
                'parents': [self.folder_id]
            }
            result = self.service.files().copy(
                fileId=template_id,
                body=file_metadata
            ).execute()
            
            if result:
                response = {
                    'spreadsheetId': result.get('id'),
                    'title': title
                }
                logger.info(f"Successfully created spreadsheet from template: {response}")
                return response
            logger.warning(f"Failed to create spreadsheet from template {template_id}")
            return {}
        except Exception as e:
            logger.error(f"Error creating spreadsheet from template {template_id}: {str(e)}", exc_info=True)
            return {}

    def create_spreadsheet_from_existing(self, source_id: str, title: str) -> Dict[str, Any]:
        """Create a new spreadsheet by copying an existing one."""
        try:
            logger.info(f"Creating new spreadsheet from existing {source_id} with title: {title}")
            file_metadata = {
                'name': title,
                'parents': [self.folder_id]
            }
            result = self.service.files().copy(
                fileId=source_id,
                body=file_metadata
            ).execute()
            
            if result:
                response = {
                    'spreadsheetId': result.get('id'),
                    'title': title
                }
                logger.info(f"Successfully created spreadsheet from existing: {response}")
                return response
            logger.warning(f"Failed to create spreadsheet from existing {source_id}")
            return {}
        except Exception as e:
            logger.error(f"Error creating spreadsheet from existing {source_id}: {str(e)}", exc_info=True)
            return {} 