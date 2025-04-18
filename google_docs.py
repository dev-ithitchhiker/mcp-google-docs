from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from typing import List, Dict, Any, Optional, Union
from google_auth import GoogleAuth
import json

class GoogleDocs:
    def __init__(self, auth: GoogleAuth):
        self.service = build('docs', 'v1', credentials=auth.get_credentials())
        self.drive_service = build('drive', 'v3', credentials=auth.get_credentials())
        self.last_insert_index = 1  # Track the last insertion index
        
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

    def insert_text(self, document_id: str, text: str, index: int = None,
                   font_family: str = None, font_size: float = None,
                   bold: bool = None, italic: bool = None,
                   underline: bool = None, strikethrough: bool = None,
                   foreground_color: str = None, background_color: str = None,
                   alignment: str = None, line_spacing: float = None,
                   space_before: float = None, space_after: float = None,
                   first_line_indent: float = None, bullet: bool = None,
                   numbered_list: bool = None) -> bool:
        """Insert text at a specific index in the document with formatting options."""
        try:
            requests = []
            
            # Insert text at the end of the document
            requests.append({
                'insertText': {
                    'text': text,
                    'endOfSegmentLocation': {}
                }
            })
            
            # Get the document to calculate the text range
            document = self.service.documents().get(documentId=document_id).execute()
            content = document.get('body', {}).get('content', [])
            
            # Calculate the actual insertion index
            actual_index = 1  # Start from the beginning of the document
            for element in content:
                if 'paragraph' in element:
                    for paragraph_element in element['paragraph']['elements']:
                        if 'textRun' in paragraph_element:
                            actual_index += len(paragraph_element['textRun']['content'])
                        elif 'inlineObjectElement' in paragraph_element:
                            actual_index += 1
                elif 'table' in element:
                    # Skip tables for now as they have their own indexing
                    continue
            
            # Calculate text range
            start_index = actual_index
            end_index = actual_index + len(text)
            
            # Apply text style if any text formatting options are provided
            if any([font_family, font_size, bold, italic, underline, strikethrough, foreground_color, background_color]):
                text_style = {}
                fields = []
                
                if font_family:
                    text_style['weightedFontFamily'] = {'fontFamily': font_family}
                    fields.append('weightedFontFamily')
                if font_size:
                    text_style['fontSize'] = {'magnitude': font_size, 'unit': 'PT'}
                    fields.append('fontSize')
                if bold is not None:
                    text_style['bold'] = bold
                    fields.append('bold')
                if italic is not None:
                    text_style['italic'] = italic
                    fields.append('italic')
                if underline is not None:
                    text_style['underline'] = underline
                    fields.append('underline')
                if strikethrough is not None:
                    text_style['strikethrough'] = strikethrough
                    fields.append('strikethrough')
                if foreground_color:
                    rgb = self._parse_color(foreground_color)
                    text_style['foregroundColor'] = {'color': {'rgbColor': rgb}}
                    fields.append('foregroundColor')
                if background_color:
                    rgb = self._parse_color(background_color)
                    text_style['backgroundColor'] = {'color': {'rgbColor': rgb}}
                    fields.append('backgroundColor')
                
                requests.append({
                    'updateTextStyle': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'textStyle': text_style,
                        'fields': ','.join(fields)
                    }
                })
            
            # Apply paragraph style if any paragraph formatting options are provided
            if any([alignment, line_spacing, space_before, space_after, first_line_indent, bullet, numbered_list]):
                paragraph_style = {}
                fields = []
                
                if alignment:
                    paragraph_style['alignment'] = alignment
                    fields.append('alignment')
                if line_spacing:
                    paragraph_style['lineSpacing'] = {'magnitude': line_spacing, 'unit': 'PT'}
                    fields.append('lineSpacing')
                if space_before:
                    paragraph_style['spaceBefore'] = {'magnitude': space_before, 'unit': 'PT'}
                    fields.append('spaceBefore')
                if space_after:
                    paragraph_style['spaceAfter'] = {'magnitude': space_after, 'unit': 'PT'}
                    fields.append('spaceAfter')
                if first_line_indent:
                    paragraph_style['firstLineIndent'] = {'magnitude': first_line_indent, 'unit': 'PT'}
                    fields.append('firstLineIndent')
                
                if paragraph_style:
                    requests.append({
                        'updateParagraphStyle': {
                            'range': {
                                'startIndex': start_index,
                                'endIndex': end_index
                            },
                            'paragraphStyle': paragraph_style,
                            'fields': ','.join(fields)
                        }
                    })
                
                if bullet:
                    requests.append({
                        'createParagraphBullets': {
                            'range': {
                                'startIndex': start_index,
                                'endIndex': end_index
                            },
                            'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                        }
                    })
                elif numbered_list:
                    requests.append({
                        'createParagraphBullets': {
                            'range': {
                                'startIndex': start_index,
                                'endIndex': end_index
                            },
                            'bulletPreset': 'NUMBERED_DECIMAL_ALPHA_ROMAN'
                        }
                    })
            
            body = {'requests': requests}
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
        """Delete a document using Google Drive API."""
        try:
            # First, verify the document exists and is accessible
            try:
                self.service.documents().get(documentId=document_id).execute()
            except HttpError as error:
                print(f'Error accessing document: {error}')
                if hasattr(error, 'content'):
                    print(f'Error content: {error.content}')
                return False

            # Try to delete the document
            try:
                self.drive_service.files().delete(
                    fileId=document_id,
                    supportsAllDrives=True  # Support shared drives
                ).execute()
                return True
            except HttpError as error:
                print(f'Error deleting document: {error}')
                if hasattr(error, 'content'):
                    print(f'Error content: {error.content}')
                return False
        except Exception as e:
            print(f'Unexpected error during document deletion: {str(e)}')
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
                        'columns': columns
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

    def update_text_style(self, document_id: str, start_index: int, end_index: int,
                         font_family: str = None, font_size: float = None,
                         bold: bool = None, italic: bool = None,
                         underline: bool = None, strikethrough: bool = None,
                         foreground_color: str = None, background_color: str = None) -> bool:
        """Update text style for a specific range.
        
        Args:
            document_id: Document ID
            start_index: Start index of the text range
            end_index: End index of the text range
            font_family: Font family name
            font_size: Font size in points
            bold: Whether to make text bold
            italic: Whether to make text italic
            underline: Whether to underline text
            strikethrough: Whether to strikethrough text
            foreground_color: Text color in hex format (e.g., '#FF0000')
            background_color: Background color in hex format (e.g., '#FFFF00')
        """
        try:
            text_style = {}
            fields = []
            
            if font_family:
                text_style['weightedFontFamily'] = {'fontFamily': font_family}
                fields.append('weightedFontFamily')
            if font_size:
                text_style['fontSize'] = {'magnitude': font_size, 'unit': 'PT'}
                fields.append('fontSize')
            if bold is not None:
                text_style['bold'] = bold
                fields.append('bold')
            if italic is not None:
                text_style['italic'] = italic
                fields.append('italic')
            if underline is not None:
                text_style['underline'] = underline
                fields.append('underline')
            if strikethrough is not None:
                text_style['strikethrough'] = strikethrough
                fields.append('strikethrough')
            if foreground_color:
                rgb = self._parse_color(foreground_color)
                text_style['foregroundColor'] = {'color': {'rgbColor': rgb}}
                fields.append('foregroundColor')
            if background_color:
                rgb = self._parse_color(background_color)
                text_style['backgroundColor'] = {'color': {'rgbColor': rgb}}
                fields.append('backgroundColor')
            
            requests = [{
                'updateTextStyle': {
                    'range': {
                        'startIndex': start_index,
                        'endIndex': end_index
                    },
                    'textStyle': text_style,
                    'fields': ','.join(fields)
                }
            }]
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def update_paragraph_style(self, document_id: str, start_index: int, end_index: int,
                             alignment: str = None, line_spacing: float = None,
                             space_before: float = None, space_after: float = None,
                             first_line_indent: float = None, bullet: bool = None,
                             numbered_list: bool = None) -> bool:
        """Update paragraph style for a specific range.
        
        Args:
            document_id: Document ID
            start_index: Start index of the paragraph range
            end_index: End index of the paragraph range
            alignment: Text alignment ('START', 'CENTER', 'END', 'JUSTIFIED')
            line_spacing: Line spacing multiplier
            space_before: Space before paragraph in points
            space_after: Space after paragraph in points
            first_line_indent: First line indent in points
            bullet: Whether to add bullet points
            numbered_list: Whether to add numbered list
        """
        try:
            paragraph_style = {}
            
            if alignment:
                paragraph_style['alignment'] = alignment
            if line_spacing:
                paragraph_style['lineSpacing'] = {'magnitude': line_spacing, 'unit': 'PT'}
            if space_before:
                paragraph_style['spaceBefore'] = {'magnitude': space_before, 'unit': 'PT'}
            if space_after:
                paragraph_style['spaceAfter'] = {'magnitude': space_after, 'unit': 'PT'}
            if first_line_indent:
                paragraph_style['firstLineIndent'] = {'magnitude': first_line_indent, 'unit': 'PT'}
            
            requests = [{
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': start_index,
                        'endIndex': end_index
                    },
                    'paragraphStyle': paragraph_style,
                    'fields': ','.join(paragraph_style.keys())
                }
            }]
            
            if bullet:
                requests.append({
                    'createParagraphBullets': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'bulletPreset': 'BULLET_DISC_CIRCLE_SQUARE'
                    }
                })
            elif numbered_list:
                requests.append({
                    'createParagraphBullets': {
                        'range': {
                            'startIndex': start_index,
                            'endIndex': end_index
                        },
                        'bulletPreset': 'NUMBERED_DECIMAL_ALPHA_ROMAN'
                    }
                })
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def insert_page_break(self, document_id: str, index: int) -> bool:
        """Insert a page break at the specified index."""
        try:
            requests = [{
                'insertPageBreak': {
                    'location': {
                        'index': index
                    }
                }
            }]
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def insert_horizontal_rule(self, document_id: str, index: int = None) -> bool:
        """Insert a horizontal rule at the specified index."""
        try:
            # If index is not provided, use the last insertion index
            if index is None:
                index = self.last_insert_index
            
            # Get the current document to calculate the correct index
            document = self.service.documents().get(documentId=document_id).execute()
            content = document.get('body', {}).get('content', [])
            
            # Calculate the actual insertion index
            actual_index = 1  # Start from the beginning of the document
            for element in content:
                if 'paragraph' in element:
                    for paragraph_element in element['paragraph']['elements']:
                        if 'textRun' in paragraph_element:
                            actual_index += len(paragraph_element['textRun']['content'])
                        elif 'inlineObjectElement' in paragraph_element:
                            actual_index += 1
                elif 'table' in element:
                    # Skip tables for now as they have their own indexing
                    continue
                
                if actual_index >= index:
                    break
            
            # Create a horizontal rule using text
            horizontal_rule = "―" * 50  # Using em dash for a horizontal line
            
            requests = [
                {
                    'insertText': {
                        'location': {
                            'index': actual_index
                        },
                        'text': horizontal_rule + '\n'
                    }
                },
                {
                    'updateParagraphStyle': {
                        'range': {
                            'startIndex': actual_index,
                            'endIndex': actual_index + len(horizontal_rule) + 1
                        },
                        'paragraphStyle': {
                            'alignment': 'CENTER'
                        },
                        'fields': 'alignment'
                    }
                }
            ]
            
            # Update the last insertion index
            self.last_insert_index = actual_index + len(horizontal_rule) + 1
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            if hasattr(error, 'content'):
                print(f'Error content: {error.content}')
            return False

    def _parse_color(self, color: str) -> Dict[str, float]:
        """Parse color string (hex or RGB) into RGB color object."""
        if color.startswith('#'):
            color = color[1:]
        if len(color) == 6:
            r = int(color[0:2], 16) / 255.0
            g = int(color[2:4], 16) / 255.0
            b = int(color[4:6], 16) / 255.0
        else:
            raise ValueError("Invalid color format. Use hex (#RRGGBB) or RGB (r,g,b)")
        return {'red': r, 'green': g, 'blue': b}

    def update_table_cell_content(self, document_id: str, table_id: str, 
                                row_index: int, column_index: int, 
                                content: str) -> bool:
        """Update content of a specific table cell.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            row_index: Row index (0-based)
            column_index: Column index (0-based)
            content: Content to insert
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Get the document with specific fields to reduce payload
            document = self.service.documents().get(
                documentId=document_id,
                fields='body.content.table(tableId,tableRows(tableCells(startIndex,endIndex)))'
            ).execute()
            
            # Find the table with matching ID
            table = None
            for element in document.get('body', {}).get('content', []):
                tbl = element.get('table')
                if tbl and tbl.get('tableId') == table_id:
                    table = tbl
                    break
            
            if not table:
                print(f"No table with ID {table_id} found")
                return False
            
            # Get the cell's start index
            try:
                table_rows = table.get('tableRows', [])
                if row_index >= len(table_rows):
                    print(f"Row index {row_index} is out of bounds")
                    return False
                    
                table_cells = table_rows[row_index].get('tableCells', [])
                if column_index >= len(table_cells):
                    print(f"Column index {column_index} is out of bounds")
                    return False
                    
                cell = table_cells[column_index]
                start_index = cell.get('startIndex')
                end_index = cell.get('endIndex')
                
                if not start_index or not end_index:
                    print("Could not find cell indices")
                    return False
                
                requests = []
                
                # Only delete content if the cell is not empty
                if end_index - start_index > 1:
                    requests.append({
                        'deleteContentRange': {
                            'range': {
                                'startIndex': start_index + 1,
                                'endIndex': end_index - 1
                            }
                        }
                    })
                
                # Insert new content
                requests.append({
                    'insertText': {
                        'location': {
                            'index': start_index + 1
                        },
                        'text': content
                    }
                })
                
                if requests:
                    body = {'requests': requests}
                    self.service.documents().batchUpdate(
                        documentId=document_id,
                        body=body
                    ).execute()
                return True
            except IndexError as e:
                print(f"Invalid row or column index: {e}")
                return False
        except HttpError as error:
            print(f'An error occurred: {error}')
            if hasattr(error, 'content'):
                print(f'Error content: {error.content}')
            return False

    def update_table_cell_style(self, document_id: str, table_id: str,
                              row_index: int, column_index: int,
                              background_color: str = None,
                              border_color: str = None,
                              border_width: float = None,
                              padding: Dict[str, float] = None) -> bool:
        """Update style of a specific table cell.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            row_index: Row index (0-based)
            column_index: Column index (0-based)
            background_color: Background color in hex format
            border_color: Border color in hex format
            border_width: Border width in points
            padding: Dictionary with 'top', 'right', 'bottom', 'left' padding values
        """
        try:
            requests = []
            table_cell_style = {}
            
            if background_color:
                table_cell_style['backgroundColor'] = {
                    'color': {'rgbColor': self._parse_color(background_color)}
                }
            
            if border_color or border_width:
                table_cell_style['borderBottom'] = {
                    'color': {'rgbColor': self._parse_color(border_color or '#000000')},
                    'width': {'magnitude': border_width or 1.0, 'unit': 'PT'}
                }
                table_cell_style['borderTop'] = table_cell_style['borderBottom']
                table_cell_style['borderLeft'] = table_cell_style['borderBottom']
                table_cell_style['borderRight'] = table_cell_style['borderBottom']
            
            if padding:
                table_cell_style['padding'] = {
                    'top': {'magnitude': padding.get('top', 0), 'unit': 'PT'},
                    'right': {'magnitude': padding.get('right', 0), 'unit': 'PT'},
                    'bottom': {'magnitude': padding.get('bottom', 0), 'unit': 'PT'},
                    'left': {'magnitude': padding.get('left', 0), 'unit': 'PT'}
                }
            
            if table_cell_style:
                requests.append({
                    'updateTableCellStyle': {
                        'tableCellStyle': table_cell_style,
                        'tableRange': {
                            'tableCellLocation': {
                                'tableId': table_id,
                                'rowIndex': row_index,
                                'columnIndex': column_index
                            }
                        },
                        'fields': ','.join(table_cell_style.keys())
                    }
                })
            
            if requests:
                body = {'requests': requests}
                self.service.documents().batchUpdate(
                    documentId=document_id,
                    body=body
                ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def update_table_row_style(self, document_id: str, table_id: str,
                             row_index: int,
                             background_color: str = None,
                             height: float = None) -> bool:
        """Update style of a table row.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            row_index: Row index (0-based)
            background_color: Background color in hex format
            height: Row height in points
        """
        try:
            requests = []
            table_row_style = {}
            
            if background_color:
                table_row_style['backgroundColor'] = {
                    'color': {'rgbColor': self._parse_color(background_color)}
                }
            
            if height:
                table_row_style['minRowHeight'] = {
                    'magnitude': height,
                    'unit': 'PT'
                }
            
            if table_row_style:
                requests.append({
                    'updateTableRowStyle': {
                        'tableRowStyle': table_row_style,
                        'rowIndices': [row_index],
                        'tableStartLocation': {
                            'index': 1  # This should be the actual table start index
                        },
                        'fields': ','.join(table_row_style.keys())
                    }
                })
            
            if requests:
                body = {'requests': requests}
                self.service.documents().batchUpdate(
                    documentId=document_id,
                    body=body
                ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def update_table_column_style(self, document_id: str, table_id: str,
                                column_index: int,
                                width: float = None) -> bool:
        """Update style of a table column.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            column_index: Column index (0-based)
            width: Column width in points
        """
        try:
            if width:
                requests = [{
                    'updateTableColumnProperties': {
                        'columnIndices': [column_index],
                        'tableStartLocation': {
                            'index': 1  # This should be the actual table start index
                        },
                        'tableColumnProperties': {
                            'width': {
                                'magnitude': width,
                                'unit': 'PT'
                            }
                        },
                        'fields': 'width'
                    }
                }]
                
                body = {'requests': requests}
                self.service.documents().batchUpdate(
                    documentId=document_id,
                    body=body
                ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def merge_table_cells(self, document_id: str, table_id: str,
                         start_row: int, start_column: int,
                         end_row: int, end_column: int) -> bool:
        """Merge table cells.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            start_row: Start row index (0-based)
            start_column: Start column index (0-based)
            end_row: End row index (0-based)
            end_column: End column index (0-based)
        """
        try:
            requests = [{
                'mergeTableCells': {
                    'tableRange': {
                        'tableCellLocation': {
                            'tableId': table_id,
                            'rowIndex': start_row,
                            'columnIndex': start_column
                        },
                        'rowSpan': end_row - start_row + 1,
                        'columnSpan': end_column - start_column + 1
                    }
                }
            }]
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def insert_table_row(self, document_id: str, table_id: str,
                        row_index: int, num_rows: int = 1) -> bool:
        """Insert rows into a table.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            row_index: Index where to insert the row(s)
            num_rows: Number of rows to insert
        """
        try:
            requests = [{
                'insertTableRow': {
                    'tableCellLocation': {
                        'tableId': table_id,
                        'rowIndex': row_index,
                        'columnIndex': 0
                    },
                    'insertBelow': True,
                    'number': num_rows
                }
            }]
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def insert_table_column(self, document_id: str, table_id: str,
                          column_index: int, num_columns: int = 1) -> bool:
        """Insert columns into a table.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            column_index: Index where to insert the column(s)
            num_columns: Number of columns to insert
        """
        try:
            requests = [{
                'insertTableColumn': {
                    'tableCellLocation': {
                        'tableId': table_id,
                        'rowIndex': 0,
                        'columnIndex': column_index
                    },
                    'insertRight': True,
                    'number': num_columns
                }
            }]
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def delete_table_row(self, document_id: str, table_id: str,
                        row_index: int, num_rows: int = 1) -> bool:
        """Delete rows from a table.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            row_index: Index of the first row to delete
            num_rows: Number of rows to delete
        """
        try:
            requests = [{
                'deleteTableRow': {
                    'tableCellLocation': {
                        'tableId': table_id,
                        'rowIndex': row_index,
                        'columnIndex': 0
                    },
                    'number': num_rows
                }
            }]
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def delete_table_column(self, document_id: str, table_id: str,
                          column_index: int, num_columns: int = 1) -> bool:
        """Delete columns from a table.
        
        Args:
            document_id: Document ID
            table_id: Table ID
            column_index: Index of the first column to delete
            num_columns: Number of columns to delete
        """
        try:
            requests = [{
                'deleteTableColumn': {
                    'tableCellLocation': {
                        'tableId': table_id,
                        'rowIndex': 0,
                        'columnIndex': column_index
                    },
                    'number': num_columns
                }
            }]
            
            body = {'requests': requests}
            self.service.documents().batchUpdate(
                documentId=document_id,
                body=body
            ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def update_document_style(self, document_id: str,
                            default_font_family: str = None,
                            default_font_size: float = None,
                            default_line_spacing: float = None,
                            default_margin_top: float = None,
                            default_margin_bottom: float = None,
                            default_margin_left: float = None,
                            default_margin_right: float = None,
                            default_page_color: str = None) -> bool:
        """Update document-wide layout settings.
        
        Args:
            document_id: Document ID
            default_margin_top: Top margin in points
            default_margin_bottom: Bottom margin in points
            default_margin_left: Left margin in points
            default_margin_right: Right margin in points
            default_page_color: Page background color in hex format
        """
        try:
            document_style = {}
            fields = []
            
            # Update margins
            if default_margin_top:
                document_style['marginTop'] = {'magnitude': default_margin_top, 'unit': 'PT'}
                fields.append('marginTop')
            if default_margin_bottom:
                document_style['marginBottom'] = {'magnitude': default_margin_bottom, 'unit': 'PT'}
                fields.append('marginBottom')
            if default_margin_left:
                document_style['marginLeft'] = {'magnitude': default_margin_left, 'unit': 'PT'}
                fields.append('marginLeft')
            if default_margin_right:
                document_style['marginRight'] = {'magnitude': default_margin_right, 'unit': 'PT'}
                fields.append('marginRight')
            
            # Update page background color
            if default_page_color:
                rgb = self._parse_color(default_page_color)
                document_style['background'] = {
                    'color': {
                        'color': {
                            'rgbColor': rgb
                        }
                    }
                }
                fields.append('background.color')
            
            if document_style:
                requests = [{
                    'updateDocumentStyle': {
                        'documentStyle': document_style,
                        'fields': ','.join(fields)
                    }
                }]
                
                body = {'requests': requests}
                self.service.documents().batchUpdate(
                    documentId=document_id,
                    body=body
                ).execute()
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            return False

    def create_table_and_insert_text(self, document_id: str, rows: int, columns: int, 
                                   content: Dict[str, str]) -> bool:
        """Create a table and insert text into specific cells.
        
        Args:
            document_id: Document ID
            rows: Number of rows in the table
            columns: Number of columns in the table
            content: Dictionary mapping cell coordinates to content
                   Format: {'row,column': 'content'}
                   Example: {'0,0': 'Header', '1,1': 'Data'}
        
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # First, create the table
            create_table_request = {
                'insertTable': {
                    'rows': rows,
                    'columns': columns,
                    'location': {
                        'index': 1
                    }
                }
            }
            
            # Execute table creation
            self.service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': [create_table_request]}
            ).execute()
            
            # Get the newly created table's structure
            document = self.service.documents().get(
                documentId=document_id,
                fields='body.content.table(tableId,tableRows(tableCells(startIndex,endIndex)))'
            ).execute()
            
            # Find the most recently created table
            table = None
            for element in document.get('body', {}).get('content', []):
                tbl = element.get('table')
                if tbl:
                    table = tbl
                    break
            
            if not table:
                print("Failed to find created table")
                return False
            
            table_id = table.get('tableId')
            requests = []
            
            # Process each cell content
            for cell_coord, cell_content in content.items():
                try:
                    row, col = map(int, cell_coord.split(','))
                    if row >= rows or col >= columns:
                        print(f"Invalid cell coordinates: {cell_coord}")
                        continue
                        
                    table_rows = table.get('tableRows', [])
                    if row >= len(table_rows):
                        continue
                        
                    table_cells = table_rows[row].get('tableCells', [])
                    if col >= len(table_cells):
                        continue
                        
                    cell = table_cells[col]
                    start_index = cell.get('startIndex')
                    
                    if start_index:
                        requests.append({
                            'insertText': {
                                'location': {
                                    'index': start_index + 1
                                },
                                'text': cell_content
                            }
                        })
                except (ValueError, IndexError) as e:
                    print(f"Error processing cell {cell_coord}: {e}")
                    continue
            
            if requests:
                self.service.documents().batchUpdate(
                    documentId=document_id,
                    body={'requests': requests}
                ).execute()
            
            return True
        except HttpError as error:
            print(f'An error occurred: {error}')
            if hasattr(error, 'content'):
                print(f'Error content: {error.content}')
            return False 