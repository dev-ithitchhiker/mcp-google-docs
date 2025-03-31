from typing import List, Dict, Any
from googleapiclient.discovery import build
from config import Config
from google_auth import GoogleAuth
import logging

logger = logging.getLogger(__name__)

class GoogleSheets:
    def __init__(self, auth: GoogleAuth):
        self.auth = auth
        self.service = auth.get_service()

    def list_sheets(self, spreadsheet_id: str) -> List[Dict[str, Any]]:
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        return spreadsheet.get('sheets', [])

    def duplicate_sheet(self, spreadsheet_id: str, sheet_id: int, new_name: str) -> Dict[str, Any]:
        """Create a new sheet by duplicating an existing one."""
        try:
            logger.info(f"Duplicating sheet {sheet_id} to new sheet with name: {new_name}")
            # Request sheet copy
            copy_request = {
                'destinationSpreadsheetId': spreadsheet_id
            }
            
            # Execute sheet copy
            result = self.service.spreadsheets().sheets().copyTo(
                spreadsheetId=spreadsheet_id,
                sheetId=sheet_id,
                body=copy_request
            ).execute()
            
            # Get new sheet ID
            new_sheet_id = result.get('sheetId')
            if not new_sheet_id:
                logger.error(f"Failed to get new sheet ID after duplicating sheet {sheet_id}")
                return {
                    'success': False,
                    'error': 'Failed to get new sheet ID',
                    'details': result
                }
            
            # Rename new sheet
            rename_result = self.rename_sheet(spreadsheet_id, new_sheet_id, new_name)
            if not rename_result:
                logger.error(f"Failed to rename new sheet {new_sheet_id}")
                return {
                    'success': False,
                    'error': 'Failed to rename new sheet',
                    'details': {'copy_result': result, 'rename_result': rename_result}
                }
            
            logger.info(f"Successfully duplicated sheet {sheet_id} to new sheet {new_sheet_id} with name: {new_name}")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error duplicating sheet {sheet_id}: {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def rename_sheet(self, spreadsheet_id: str, sheet_id: int, new_name: str) -> Dict[str, Any]:
        """Rename a sheet."""
        try:
            logger.info(f"Renaming sheet {sheet_id} to '{new_name}'")
            body = {
                'requests': [{
                    'updateSheetProperties': {
                        'properties': {
                            'sheetId': sheet_id,
                            'title': new_name
                        },
                        'fields': 'title'
                    }
                }]
            }
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to rename sheet {sheet_id}")
                return {
                    'success': False,
                    'error': 'Failed to rename sheet',
                    'details': result
                }
            
            logger.info(f"Successfully renamed sheet {sheet_id} to '{new_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error renaming sheet {sheet_id}: {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def get_sheet_data(self, spreadsheet_id: str, sheet_name: str, range_name: str) -> Dict[str, Any]:
        """Get data from a sheet."""
        try:
            logger.info(f"Getting data from sheet '{sheet_name}' range '{range_name}'")
            result = self.service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!{range_name}"
            ).execute()
            
            if not result or 'values' not in result:
                logger.error(f"Failed to get data from sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to get data',
                    'details': result
                }
            
            logger.info(f"Successfully retrieved data from sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result.get('values', [])
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error getting data from sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def add_rows(self, spreadsheet_id: str, sheet_name: str, values: List[List[Any]]) -> Dict[str, Any]:
        """Add rows to a sheet."""
        try:
            logger.info(f"Adding rows to sheet '{sheet_name}'")
            body = {
                'values': values
            }
            result = self.service.spreadsheets().values().append(
                spreadsheetId=spreadsheet_id,
                range=sheet_name,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            if not result or 'updates' not in result:
                logger.error(f"Failed to add rows to sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to add rows',
                    'details': result
                }
            
            logger.info(f"Successfully added rows to sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error adding rows to sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def add_columns(self, spreadsheet_id: str, sheet_name: str, values: List[List[Any]]) -> Dict[str, Any]:
        """Add columns to a sheet."""
        try:
            logger.info(f"Adding columns to sheet '{sheet_name}'")
            # Get sheet ID
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is None:
                logger.error(f"Sheet '{sheet_name}' not found")
                return {
                    'success': False,
                    'error': f"Sheet '{sheet_name}' not found"
                }

            # Add columns
            body = {
                'requests': [{
                    'appendDimension': {
                        'sheetId': sheet_id,
                        'dimension': 'COLUMNS',
                        'length': len(values[0]) if values else 1
                    }
                }]
            }
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to add columns to sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to add columns',
                    'details': result
                }
            
            logger.info(f"Successfully added columns to sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error adding columns to sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def update_cells(self, values: List[List[Any]], range_name: str, sheet_name: str, spreadsheet_id: str, format: Dict[str, Any] = None) -> Dict[str, Any]:
        """Update cells in a sheet."""
        try:
            logger.info(f"Updating cells in sheet '{sheet_name}' range '{range_name}'")
            
            # Get sheet ID
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is None:
                logger.error(f"Sheet '{sheet_name}' not found")
                return {
                    'success': False,
                    'error': f"Sheet '{sheet_name}' not found"
                }
            
            # Parse range_name
            range_info = self._parse_range(range_name)
            
            # Data dimension validation
            expected_cols = range_info['endColumnIndex'] - range_info['startColumnIndex']
            if any(len(row) != expected_cols for row in values):
                logger.error(f"Data dimensions do not match range. Expected {expected_cols} columns, got {len(values[0])}")
                return {
                    'success': False,
                    'error': f"Data dimensions do not match range. Expected {expected_cols} columns, got {len(values[0])}"
                }
            
            # format parameter handling
            if isinstance(format, str):
                try:
                    import json
                    format = json.loads(format)
                except json.JSONDecodeError:
                    logger.error(f"Invalid format parameter: {format}")
                    return {
                        'success': False,
                        'error': f"Invalid format parameter: {format}"
                    }
            
            # Create update request
            updates = [{
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': range_info['startRowIndex'],
                    'endRowIndex': range_info['endRowIndex'] + 1,  # Add 1 because endRowIndex is exclusive
                    'startColumnIndex': range_info['startColumnIndex'],
                    'endColumnIndex': range_info['endColumnIndex'] + 1  # Add 1 because endColumnIndex is exclusive
                },
                'values': values,
                'merge': False,
                'format': format
            }]
            
            # Use batch_update_cells
            result = self.batch_update_cells(spreadsheet_id, sheet_name, updates)
            
            if not result['success']:
                logger.error(f"Failed to update cells: {result['error']}")
                return result
            
            logger.info(f"Successfully updated cells: {result}")
            return result
            
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error updating cells: {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def _parse_html_tags(self, text: str) -> Dict[str, Any]:
        """Convert HTML tags to Google Sheets style."""
        style = {}
        content = text
        
        # Remove HTML tags and extract styles
        while '<' in content and '>' in content:
            start = content.find('<')
            end = content.find('>', start)
            if end == -1:
                break
                
            tag = content[start:end+1]
            content = content[:start] + content[end+1:]
            
            # Handle closing tags
            if tag.startswith('</'):
                continue
                
            # Extract tag content
            tag_name = tag[1:].split()[0].lower()
            tag_attrs = {}
            
            # Extract attributes
            if ' ' in tag:
                attrs = tag[tag.find(' ')+1:-1].split()
                for attr in attrs:
                    if '=' in attr:
                        key, value = attr.split('=', 1)
                        tag_attrs[key.lower()] = value.strip("'\"")
            
            # Apply styles
            if tag_name in ['b', 'strong']:
                style['bold'] = True
            elif tag_name in ['i', 'em']:
                style['italic'] = True
            elif tag_name in ['s', 'strike', 'del']:
                style['strikethrough'] = True
            elif tag_name == 'u':
                style['underline'] = True
            elif tag_name in ['h1', 'h2', 'h3']:
                style['fontSize'] = {'h1': 24, 'h2': 20, 'h3': 16}[tag_name]
            elif tag_name == 'small':
                style['fontSize'] = 10
            elif tag_name == 'font':
                if 'color' in tag_attrs:
                    color = tag_attrs['color']
                    if color.startswith('#'):
                        color = color[1:]
                    r = int(color[0:2], 16) / 255
                    g = int(color[2:4], 16) / 255
                    b = int(color[4:6], 16) / 255
                    style['foregroundColor'] = {'red': r, 'green': g, 'blue': b, 'alpha': 1}
            elif tag_name == 'bg':
                if 'color' in tag_attrs:
                    color = tag_attrs['color']
                    if color.startswith('#'):
                        color = color[1:]
                    r = int(color[0:2], 16) / 255
                    g = int(color[2:4], 16) / 255
                    b = int(color[4:6], 16) / 255
                    style['backgroundColor'] = {'red': r, 'green': g, 'blue': b, 'alpha': 1}
            elif tag_name in ['center', 'left', 'right']:
                style['horizontalAlignment'] = tag_name.upper()
        
        return {
            'text': content.strip(),
            'style': style if style else None
        }

    def _parse_range(self, range_str: str) -> Dict[str, int]:
        """Parse range string to convert to indices."""
        # Example: 'A1:D1' -> {'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 3}
        if ':' in range_str:
            start_range, end_range = range_str.split(':')
            # Extract column letter (use only the first character)
            start_col = start_range[0].upper()
            start_row = int(start_range[1:]) - 1  # 0-based index
            end_col = end_range[0].upper()
            end_row = int(end_range[1:]) - 1  # 0-based index
            
            # Convert column letter to index (A=0, B=1, ...)
            start_col_index = ord(start_col) - ord('A')
            end_col_index = ord(end_col) - ord('A')
        else:
            # Handle single cell range
            start_col = range_str[0].upper()
            start_row = int(range_str[1:]) - 1  # 0-based index
            start_col_index = ord(start_col) - ord('A')
            end_col_index = start_col_index + 1
            end_row = start_row + 1
        
        return {
            'startRowIndex': start_row,
            'endRowIndex': end_row,
            'startColumnIndex': start_col_index,
            'endColumnIndex': end_col_index
        }

    def batch_update_cells(self, spreadsheet_id: str, sheet_name: str, updates: List[Dict[str, Any]]) -> Dict[str, Any]:
        """Batch update cells in a sheet. Supports HTML tags."""
        try:
            logger.info(f"Batch updating cells in sheet '{sheet_name}'")
            # Get sheet ID
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            # Create new sheet if sheet does not exist
            if sheet_id is None:
                logger.info(f"Sheet '{sheet_name}' not found, creating new sheet")
                create_result = self.add_sheet(spreadsheet_id, sheet_name)
                if not create_result['success']:
                    logger.error(f"Failed to create new sheet '{sheet_name}'")
                    return create_result
                sheet_id = create_result['result']['sheet_id']

            # Create update requests
            requests = []
            for update in updates:
                if 'range' in update and 'values' in update:
                    # Handle range as string
                    if isinstance(update['range'], str):
                        range_info = self._parse_range(update['range'])
                    else:
                        range_info = update['range']
                    
                    values = update['values']
                    merge = update.get('merge', False)
                    format_config = update.get('format', {})

                    # Create cell update request
                    cell_request = {
                        'updateCells': {
                            'range': {
                                'sheetId': sheet_id,
                                'startRowIndex': range_info['startRowIndex'],
                                'endRowIndex': range_info['endRowIndex'] + 1,  # Add 1 because endRowIndex is exclusive
                                'startColumnIndex': range_info['startColumnIndex'],
                                'endColumnIndex': range_info['endColumnIndex'] + 1  # Add 1 because endColumnIndex is exclusive
                            },
                            'rows': [],
                            'fields': '*'
                        }
                    }

                    # Add data and style for each row
                    for row in values:
                        row_data = {'values': []}
                        for value in row:
                            # Handle empty string
                            if value == "":
                                cell_data = {'userEnteredValue': None}
                            else:
                                # Parse HTML tags
                                parsed = self._parse_html_tags(str(value))
                                cell_data = {'userEnteredValue': {'stringValue': parsed['text']}}
                                
                                # Apply styles
                                cell_data['userEnteredFormat'] = {}
                                
                                # Apply default styles
                                if format_config:
                                    if 'textFormat' in format_config:
                                        cell_data['userEnteredFormat']['textFormat'] = format_config['textFormat']
                                    if 'backgroundColor' in format_config:
                                        cell_data['userEnteredFormat']['backgroundColor'] = format_config['backgroundColor']
                                
                                # Apply HTML tag styles
                                if parsed['style']:
                                    if 'textFormat' not in cell_data['userEnteredFormat']:
                                        cell_data['userEnteredFormat']['textFormat'] = {}
                                    
                                    # Text styles
                                    if 'bold' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['bold'] = parsed['style']['bold']
                                    if 'italic' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['italic'] = parsed['style']['italic']
                                    if 'strikethrough' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['strikethrough'] = parsed['style']['strikethrough']
                                    if 'underline' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['underline'] = parsed['style']['underline']
                                    
                                    # Font size
                                    if 'fontSize' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['fontSize'] = parsed['style']['fontSize']
                                    
                                    # Font color
                                    if 'foregroundColor' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['foregroundColor'] = parsed['style']['foregroundColor']
                                    
                                    # Background color
                                    if 'backgroundColor' in parsed['style']:
                                        cell_data['userEnteredFormat']['backgroundColor'] = parsed['style']['backgroundColor']
                                    
                                    # Alignment
                                    if 'horizontalAlignment' in parsed['style']:
                                        cell_data['userEnteredFormat']['horizontalAlignment'] = parsed['style']['horizontalAlignment']
                            
                            row_data['values'].append(cell_data)
                        cell_request['updateCells']['rows'].append(row_data)

                    requests.append(cell_request)

                    # Create cell merge request
                    if merge:
                        merge_request = {
                            'mergeCells': {
                                'range': {
                                    'sheetId': sheet_id,
                                    'startRowIndex': range_info['startRowIndex'],
                                    'endRowIndex': range_info['endRowIndex'] + 1,
                                    'startColumnIndex': range_info['startColumnIndex'],
                                    'endColumnIndex': range_info['endColumnIndex'] + 1
                                },
                                'mergeType': 'MERGE_ALL'
                            }
                        }
                        requests.append(merge_request)

            if not requests:
                logger.warning("No valid update requests found")
                return {
                    'success': False,
                    'error': 'No valid update requests found'
                }

            body = {
                'requests': requests
            }
            
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to batch update cells in sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to batch update cells',
                    'details': result
                }
            
            logger.info(f"Successfully batch updated cells in sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error batch updating cells in sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def delete_rows(self, spreadsheet_id: str, sheet_name: str, start_index: int, end_index: int) -> Dict[str, Any]:
        """Delete rows from a sheet."""
        try:
            logger.info(f"Deleting rows {start_index} to {end_index} from sheet '{sheet_name}'")
            # Get sheet ID
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is None:
                logger.error(f"Sheet '{sheet_name}' not found")
                return {
                    'success': False,
                    'error': f"Sheet '{sheet_name}' not found"
                }

            body = {
                'requests': [{
                    'deleteDimension': {
                        'range': {
                            'sheetId': sheet_id,
                            'dimension': 'ROWS',
                            'startIndex': start_index,
                            'endIndex': end_index
                        }
                    }
                }]
            }
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to delete rows from sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to delete rows',
                    'details': result
                }
            
            logger.info(f"Successfully deleted rows from sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error deleting rows from sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def delete_columns(self, spreadsheet_id: str, sheet_name: str, start_index: int, end_index: int) -> Dict[str, Any]:
        """Delete columns from a sheet."""
        try:
            logger.info(f"Deleting columns {start_index} to {end_index} from sheet '{sheet_name}'")
            # Get sheet ID
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is None:
                logger.error(f"Sheet '{sheet_name}' not found")
                return {
                    'success': False,
                    'error': f"Sheet '{sheet_name}' not found"
                }

            body = {
                'requests': [{
                    'deleteDimension': {
                        'range': {
                            'sheetId': sheet_id,
                            'dimension': 'COLUMNS',
                            'startIndex': start_index,
                            'endIndex': end_index
                        }
                    }
                }]
            }
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to delete columns from sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to delete columns',
                    'details': result
                }
            
            logger.info(f"Successfully deleted columns from sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error deleting columns from sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def create_chart(self, chart_type: str, range_name: str, sheet_name: str, spreadsheet_id: str, title: str = None) -> Dict[str, Any]:
        """Create a chart.
        
        Args:
            chart_type: Chart type ('LINE', 'COLUMN', 'PIE', 'SCATTER', 'BAR')
            range_name: Data range (e.g., 'A1:B10')
            sheet_name: Sheet name
            spreadsheet_id: Spreadsheet ID
            title: Chart title (optional)
        """
        try:
            # Remove quotes from sheet_name
            sheet_name = sheet_name.strip('"')
            logger.info(f"Creating chart in sheet '{sheet_name}'")
            
            # Get sheet ID
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is None:
                logger.error(f"Sheet '{sheet_name}' not found")
                return {
                    'success': False,
                    'error': f"Sheet '{sheet_name}' not found"
                }

            # Parse range_name
            range_info = self._parse_range(range_name)

            # Create chart request
            chart_request = {
                'addChart': {
                    'chart': {
                        'spec': {
                            'title': title or sheet_name,
                            'basicChart': {
                                'chartType': chart_type,
                                'legendPosition': 'RIGHT_LEGEND',
                                'domains': [{
                                    'domain': {
                                        'sourceRange': {
                                            'sources': [{
                                                'sheetId': sheet_id,
                                                'startRowIndex': range_info['startRowIndex'],
                                                'endRowIndex': range_info['endRowIndex'],
                                                'startColumnIndex': range_info['startColumnIndex'],
                                                'endColumnIndex': range_info['startColumnIndex'] + 1
                                            }]
                                        }
                                    }
                                }],
                                'series': []
                            }
                        },
                        'position': {
                            'overlayPosition': {
                                'anchorCell': {
                                    'sheetId': sheet_id,
                                    'rowIndex': range_info['startRowIndex'],
                                    'columnIndex': range_info['startColumnIndex']
                                },
                                'widthPixels': 600,
                                'heightPixels': 371
                            }
                        }
                    }
                }
            }

            # Add all series
            for i in range(1, range_info['endColumnIndex'] - range_info['startColumnIndex']):
                series = {
                    'series': {
                        'sourceRange': {
                            'sources': [{
                                'sheetId': sheet_id,
                                'startRowIndex': range_info['startRowIndex'],
                                'endRowIndex': range_info['endRowIndex'],
                                'startColumnIndex': range_info['startColumnIndex'] + i,
                                'endColumnIndex': range_info['startColumnIndex'] + i + 1
                            }]
                        }
                    },
                    'targetAxis': 'LEFT_AXIS'
                }
                chart_request['addChart']['chart']['spec']['basicChart']['series'].append(series)

            body = {
                'requests': [chart_request]
            }
            
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to create chart in sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to create chart',
                    'details': result
                }
            
            logger.info(f"Successfully created chart in sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error creating chart in sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def update_chart(self, spreadsheet_id: str, sheet_name: str, chart_id: int, chart_config: Dict[str, Any]) -> Dict[str, Any]:
        """Update an existing chart."""
        try:
            logger.info(f"Updating chart {chart_id} in sheet '{sheet_name}'")
            # Get sheet ID
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is None:
                logger.error(f"Sheet '{sheet_name}' not found")
                return {
                    'success': False,
                    'error': f"Sheet '{sheet_name}' not found"
                }

            # Create chart update request
            chart_request = {
                'updateChartSpec': {
                    'chartId': chart_id,
                    'spec': {
                        'title': chart_config['title'],
                        'basicChart': {
                            'type': chart_config['type'],
                            'domains': [{
                                'domain': {
                                    'sourceRange': {
                                        'sources': [{
                                            'sheetId': sheet_id,
                                            'startRowIndex': chart_config['data_range']['start_row'],
                                            'endRowIndex': chart_config['data_range']['end_row'],
                                            'startColumnIndex': chart_config['data_range']['start_col'],
                                            'endColumnIndex': chart_config['data_range']['end_col']
                                        }]
                                    }
                                }
                            }],
                            'series': [{
                                'series': {
                                    'sourceRange': {
                                        'sources': [{
                                            'sheetId': sheet_id,
                                            'startRowIndex': chart_config['data_range']['start_row'],
                                            'endRowIndex': chart_config['data_range']['end_row'],
                                            'startColumnIndex': chart_config['data_range']['start_col'] + 1,
                                            'endColumnIndex': chart_config['data_range']['end_col']
                                        }]
                                    }
                                }
                            }]
                        }
                    }
                }
            }

            # Apply chart options
            if 'options' in chart_config:
                options = chart_config['options']
                if 'width' in options:
                    chart_request['updateChartSpec']['spec']['titleTextFormat']['fontSize'] = options['width']
                if 'height' in options:
                    chart_request['updateChartSpec']['spec']['titleTextFormat']['fontSize'] = options['height']
                if 'legend_position' in options:
                    chart_request['updateChartSpec']['spec']['basicChart']['legendPosition'] = options['legend_position']
                if 'axis_title' in options:
                    if 'x' in options['axis_title']:
                        chart_request['updateChartSpec']['spec']['basicChart']['axis'] = [{
                            'position': 'BOTTOM_AXIS',
                            'title': options['axis_title']['x']
                        }]
                    if 'y' in options['axis_title']:
                        chart_request['updateChartSpec']['spec']['basicChart']['axis'].append({
                            'position': 'LEFT_AXIS',
                            'title': options['axis_title']['y']
                        })
                if 'series' in options:
                    for i, series in enumerate(options['series']):
                        if i < len(chart_request['updateChartSpec']['spec']['basicChart']['series']):
                            if 'color' in series:
                                chart_request['updateChartSpec']['spec']['basicChart']['series'][i]['color'] = series['color']
                            if 'line_width' in series:
                                chart_request['updateChartSpec']['spec']['basicChart']['series'][i]['lineStyle'] = {
                                    'width': series['line_width']
                                }
                            if 'point_size' in series:
                                chart_request['updateChartSpec']['spec']['basicChart']['series'][i]['pointStyle'] = {
                                    'size': series['point_size']
                                }

            body = {
                'requests': [chart_request]
            }
            
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to update chart {chart_id} in sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to update chart',
                    'details': result
                }
            
            logger.info(f"Successfully updated chart {chart_id} in sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error updating chart {chart_id} in sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def delete_chart(self, spreadsheet_id: str, sheet_name: str, chart_id: int) -> Dict[str, Any]:
        """Delete a chart."""
        try:
            logger.info(f"Deleting chart {chart_id} from sheet '{sheet_name}'")
            body = {
                'requests': [{
                    'deleteEmbeddedObject': {
                        'objectId': chart_id
                    }
                }]
            }
            
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to delete chart {chart_id} from sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to delete chart',
                    'details': result
                }
            
            logger.info(f"Successfully deleted chart {chart_id} from sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error deleting chart {chart_id} from sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def add_sheet(self, spreadsheet_id: str, sheet_name: str) -> Dict[str, Any]:
        """Add a new sheet."""
        try:
            logger.info(f"Adding new sheet '{sheet_name}' to spreadsheet")
            
            # Create new sheet request
            body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name,
                            'gridProperties': {
                                'rowCount': 1000,  # Default row count
                                'columnCount': 26  # Default column count (A-Z)
                            }
                        }
                    }
                }]
            }
            
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to add new sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to add new sheet',
                    'details': result
                }
            
            # Get ID of newly created sheet
            new_sheet_id = result['replies'][0]['addSheet']['properties']['sheetId']
            
            logger.info(f"Successfully added new sheet '{sheet_name}' with ID: {new_sheet_id}")
            return {
                'success': True,
                'result': {
                    'sheet_id': new_sheet_id,
                    'sheet_name': sheet_name
                }
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error adding new sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

    def update_cell_format(self, spreadsheet_id: str, sheet_name: str, range_name: str, format: Dict[str, Any]) -> Dict[str, Any]:
        """Change cell format style.
        
        Args:
            spreadsheet_id: Spreadsheet ID
            sheet_name: Sheet name
            range_name: Range (e.g., 'A1:B5')
            format: Format setting information. It has the following format:
                {
                    'textFormat': {  # Text format
                        'fontFamily': str,  # Font
                        'fontSize': int,  # Font size
                        'bold': bool,  # Bold
                        'italic': bool,  # Italic
                        'strikethrough': bool,  # Strikethrough
                        'underline': bool,  # Underline
                        'foregroundColor': {  # Font color
                            'red': float,  # 0.0 ~ 1.0
                            'green': float,
                            'blue': float,
                            'alpha': float
                        }
                    },
                    'backgroundColor': {  # Background color
                        'red': float,
                        'green': float,
                        'blue': float,
                        'alpha': float
                    },
                    'horizontalAlignment': str,  # Horizontal alignment ('LEFT', 'CENTER', 'RIGHT')
                    'verticalAlignment': str,  # Vertical alignment ('TOP', 'MIDDLE', 'BOTTOM')
                    'padding': {  # Padding
                        'top': int,  # Top padding
                        'right': int,  # Right padding
                        'bottom': int,  # Bottom padding
                        'left': int  # Left padding
                    },
                    'wrapText': bool,  # Auto wrap text
                    'textRotation': {  # Text rotation
                        'angle': int  # Rotation angle (0 ~ 360)
                    }
                }
        """
        try:
            logger.info(f"Updating cell format in sheet '{sheet_name}' range '{range_name}'")
            
            # Get sheet ID
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            if sheet_id is None:
                logger.error(f"Sheet '{sheet_name}' not found")
                return {
                    'success': False,
                    'error': f"Sheet '{sheet_name}' not found"
                }

            # Parse range_name
            range_info = self._parse_range(range_name)
            
            # Create format update request
            format_request = {
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': range_info['startRowIndex'],
                        'endRowIndex': range_info['endRowIndex'] + 1,
                        'startColumnIndex': range_info['startColumnIndex'],
                        'endColumnIndex': range_info['endColumnIndex'] + 1
                    },
                    'cell': {
                        'userEnteredFormat': {}
                    },
                    'fields': 'userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment,padding,wrapText,textRotation)'
                }
            }

            # Apply format settings
            if 'textFormat' in format:
                format_request['repeatCell']['cell']['userEnteredFormat']['textFormat'] = format['textFormat']
            if 'backgroundColor' in format:
                format_request['repeatCell']['cell']['userEnteredFormat']['backgroundColor'] = format['backgroundColor']
            if 'horizontalAlignment' in format:
                format_request['repeatCell']['cell']['userEnteredFormat']['horizontalAlignment'] = format['horizontalAlignment']
            if 'verticalAlignment' in format:
                format_request['repeatCell']['cell']['userEnteredFormat']['verticalAlignment'] = format['verticalAlignment']
            if 'padding' in format:
                format_request['repeatCell']['cell']['userEnteredFormat']['padding'] = format['padding']
            if 'wrapText' in format:
                format_request['repeatCell']['cell']['userEnteredFormat']['wrapText'] = format['wrapText']
            if 'textRotation' in format:
                format_request['repeatCell']['cell']['userEnteredFormat']['textRotation'] = format['textRotation']

            body = {
                'requests': [format_request]
            }
            
            result = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body=body
            ).execute()
            
            if not result or 'replies' not in result or not result['replies']:
                logger.error(f"Failed to update cell format in sheet '{sheet_name}'")
                return {
                    'success': False,
                    'error': 'Failed to update cell format',
                    'details': result
                }
            
            logger.info(f"Successfully updated cell format in sheet '{sheet_name}'")
            return {
                'success': True,
                'result': result
            }
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Error updating cell format in sheet '{sheet_name}': {error_msg}", exc_info=True)
            return {
                'success': False,
                'error': error_msg
            }

def create_chart(chart_type: str, range_name: str, sheet_name: str, spreadsheet_id: str, title: str = None) -> Dict[str, Any]:
    """Create a chart.
    
    Args:
        chart_type: Chart type ('LINE', 'COLUMN', 'PIE', 'SCATTER', 'BAR')
        range_name: Data range (e.g., 'A1:B10')
        sheet_name: Sheet name
        spreadsheet_id: Spreadsheet ID
        title: Chart title (optional)
    """
    auth = GoogleAuth()
    sheets = GoogleSheets(auth)
    return sheets.create_chart(chart_type, range_name, sheet_name, spreadsheet_id, title)

def duplicate_sheet(values: List[List[Any]], range_name: str, sheet_name: str, spreadsheet_id: str, source_sheet_id: int, new_name: str) -> Dict[str, Any]:
    """Create a new sheet by duplicating an existing one.
    
    Args:
        values: Sheet data (not used)
        range_name: Data range (not used)
        sheet_name: Sheet name (not used)
        spreadsheet_id: Spreadsheet ID
        source_sheet_id: ID of the original sheet to copy
        new_name: New sheet name
    """
    auth = GoogleAuth()
    sheets = GoogleSheets(auth)
    return sheets.duplicate_sheet(spreadsheet_id, source_sheet_id, new_name)

def add_sheet(values: List[List[Any]], range_name: str, sheet_name: str, spreadsheet_id: str) -> Dict[str, Any]:
    """Add a new sheet.
    
    Args:
        values: Sheet data (not used)
        range_name: Data range (not used)
        sheet_name: Sheet name
        spreadsheet_id: Spreadsheet ID
    """
    auth = GoogleAuth()
    sheets = GoogleSheets(auth)
    return sheets.add_sheet(spreadsheet_id, sheet_name) 