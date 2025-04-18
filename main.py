import os
import json
import logging
import signal
import sys
import warnings
from mcp.server.fastmcp import FastMCP
from config import Config
from google_auth import GoogleAuth
from google_sheets import GoogleSheets
from google_drive import GoogleDrive
from google_slides import GoogleSlides
from google_docs import GoogleDocs
from typing import List, Dict, Any

# Ignore warnings
warnings.filterwarnings('ignore', message='file_cache is only supported with oauth2client<4.0.0')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Set googleapiclient logging level to WARNING
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.WARNING)

# Initialize MCP server
mcp = FastMCP("Google Spreadsheet MCP")

# Initialize configuration and authentication
config = Config.from_env()
auth = GoogleAuth(config)
drive = GoogleDrive(auth)
sheets = GoogleSheets(auth)
slides = GoogleSlides(auth)
docs = GoogleDocs(auth)

# Store current spreadsheet ID as a global variable
current_spreadsheet_id = None

@mcp.tool()
def list_files() -> List[Dict[str, Any]]:
    """List files in Google Drive."""
    return drive.list_files()

@mcp.tool()
def copy_file(file_id: str, new_name: str) -> Dict[str, Any]:
    """Copy a file."""
    return drive.copy_file(file_id, new_name)

@mcp.tool()
def rename_file(file_id: str, new_name: str) -> Dict[str, Any]:
    """Rename a file."""
    return drive.rename_file(file_id, new_name)

@mcp.tool()
def create_spreadsheet(title: str) -> Dict[str, Any]:
    """Create an empty google sheets spreadsheet.
    
    Args:
        title: Title of the new spreadsheet
    """
    global current_spreadsheet_id
    logger.info(f"Creating new spreadsheet with title: {title}")
    result = drive.create_spreadsheet(title)
    if result:
        spreadsheet_id = result.get('spreadsheetId') or result.get('id')
        if spreadsheet_id:
            current_spreadsheet_id = spreadsheet_id
            logger.info(f"Successfully created spreadsheet with ID: {current_spreadsheet_id}")
        else:
            logger.error("Failed to get spreadsheet ID from result")
    else:
        logger.error("Failed to create spreadsheet")
    return result

@mcp.tool()
def create_spreadsheet_from_template(template_id: str, title: str) -> Dict[str, Any]:
    """Create a new google sheets spreadsheet from a template.

    Args:
        template_id: Template ID
        title: Title of the new spreadsheet
    """
    global current_spreadsheet_id
    result = drive.create_spreadsheet_from_template(template_id, title)
    if result:
        spreadsheet_id = result.get('spreadsheetId') or result.get('id')
        if spreadsheet_id:
            current_spreadsheet_id = spreadsheet_id
            logger.info(f"Successfully created spreadsheet from template with ID: {current_spreadsheet_id}")
        else:
            logger.error("Failed to get spreadsheet ID from result")
    return result

@mcp.tool()
def create_spreadsheet_from_existing(source_id: str, title: str) -> Dict[str, Any]:
    """Create a new google sheets spreadsheet by copying an existing one."""
    global current_spreadsheet_id
    result = drive.create_spreadsheet_from_existing(source_id, title)
    if result:
        spreadsheet_id = result.get('spreadsheetId') or result.get('id')
        if spreadsheet_id:
            current_spreadsheet_id = spreadsheet_id
            logger.info(f"Successfully created spreadsheet from existing with ID: {current_spreadsheet_id}")
        else:
            logger.error("Failed to get spreadsheet ID from result")
    return result

@mcp.tool()
def list_sheets(spreadsheet_id: str = None) -> List[Dict[str, Any]]:
    """List sheets in a google sheets spreadsheet."""
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.list_sheets(spreadsheet_id)

@mcp.tool()
def add_sheet(spreadsheet_id: str, sheet_name: str) -> Dict[str, Any]:
    """Create a new sheet in a google sheets spreadsheet."""
    global current_spreadsheet_id
    return sheets.add_sheet(spreadsheet_id, sheet_name)

@mcp.tool()
def duplicate_sheet(spreadsheet_id: str, sheet_id: int, new_name: str) -> Dict[str, Any]:
    """Create a new sheet in a google sheets spreadsheet by duplicating an existing one.
    
    Args:
        values: Sheet data
        range_name: Data range (e.g., 'A1:B5')
        sheet_name: Sheet name
        spreadsheet_id: Spreadsheet ID
        source_sheet_id: Source sheet ID to duplicate
        new_name: New sheet name
    """
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.duplicate_sheet(spreadsheet_id, sheet_id, new_name)

@mcp.tool()
def rename_sheet(spreadsheet_id: str, sheet_id: int, new_name: str) -> Dict[str, Any]:
    """Rename a sheet in a google sheets spreadsheet."""
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.rename_sheet(spreadsheet_id, sheet_id, new_name)

@mcp.tool()
def get_sheet_data(spreadsheet_id: str, sheet_name: str, range_name: str) -> List[List[Any]]:
    """Get data from a sheet in a google sheets spreadsheet."""
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.get_sheet_data(spreadsheet_id, sheet_name, range_name)

@mcp.tool()
def add_rows(spreadsheet_id: str, sheet_name: str, values: List[List[Any]]) -> Dict[str, Any]:
    """Add rows to a sheet in a google sheets spreadsheet."""
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.add_rows(spreadsheet_id, sheet_name, values)

@mcp.tool()
def add_columns(spreadsheet_id: str, sheet_name: str, values: List[List[Any]]) -> Dict[str, Any]:
    """Add columns to a sheet in a google sheets spreadsheet."""
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.add_columns(spreadsheet_id, sheet_name, values)

@mcp.tool()
def update_cells(spreadsheet_id: str, sheet_name: str, range_name: str, values: List[List[Any]]) -> Dict[str, Any]:
    """Update cells in a sheet in a google sheets spreadsheet. Supports HTML tags and formatting.
        
        Args:
            values: Values to update
            range_name: Range (e.g., 'A1:B5')
            sheet_name: Sheet name
            spreadsheet_id: Spreadsheet ID
            format: Format settings. Has the following structure:
                {
                    'textFormat': {  # Text formatting
                        'fontFamily': str,  # Font family
                        'fontSize': int,  # Font size
                        'bold': bool,  # Bold
                        'italic': bool,  # Italic
                        'strikethrough': bool,  # Strikethrough
                        'underline': bool,  # Underline
                        'foregroundColor': {  # Text color
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
                    'wrapText': bool,  # Text wrapping
                    'textRotation': {  # Text rotation
                        'angle': int  # Rotation angle (0 ~ 360)
                    }
                }
        """
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.update_cells(spreadsheet_id, sheet_name, range_name, values)

@mcp.tool()
def batch_update_cells(spreadsheet_id: str, sheet_name: str, updates: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Update multiple cells in a sheet at once. Supports HTML tags.
        
        Args:
            updates: Values to update
            sheet_name: Sheet name
            spreadsheet_id: Spreadsheet ID
            format: Format settings. Has the following structure:
                {
                    'textFormat': {  # Text formatting
                        'fontFamily': str,  # Font family
                        'fontSize': int,  # Font size
                        'bold': bool,  # Bold
                        'italic': bool,  # Italic
                        'strikethrough': bool,  # Strikethrough
                        'underline': bool,  # Underline
                        'foregroundColor': {  # Text color
                            'red': float,  # 0.0 ~ 1.0
                            'green': float,
                            'blue': float,
                            'alpha': float
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
                        'wrapText': bool,  # Text wrapping
                        'textRotation': {  # Text rotation
                            'angle': int  # Rotation angle (0 ~ 360)
                        }
                    }
                }
        """

    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.batch_update_cells(spreadsheet_id, sheet_name, updates)

@mcp.tool()
def delete_rows(spreadsheet_id: str, sheet_name: str, start_index: int, end_index: int) -> Dict[str, Any]:
    """Delete rows from a google sheets spreadsheet."""
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.delete_rows(spreadsheet_id, sheet_name, start_index, end_index)

@mcp.tool()
def delete_columns(spreadsheet_id: str, sheet_name: str, start_index: int, end_index: int) -> Dict[str, Any]:
    """Delete columns from a sheet in a google sheets spreadsheet.

    Args:
        spreadsheet_id: Spreadsheet ID
        sheet_name: Sheet name
        start_index: Start index of columns to delete
        end_index: End index of columns to delete
    """
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")
    return sheets.delete_columns(spreadsheet_id, sheet_name, start_index, end_index)

@mcp.tool()
def create_chart(chart_type: str, range_name: str, sheet_name: str, spreadsheet_id: str, title: str = None) -> Dict[str, Any]:
    """Create a chart in a google sheets spreadsheet.
        
        Args:
            chart_type: Chart type ('LINE', 'COLUMN', 'PIE', 'SCATTER', 'BAR')
            range_name: Data range (e.g., 'A1:B10')
            sheet_name: Sheet name
            spreadsheet_id: Spreadsheet ID
            title: Chart title (optional)
        """
    global current_spreadsheet_id
    if spreadsheet_id is None:
        spreadsheet_id = current_spreadsheet_id
        if current_spreadsheet_id is None:
            logger.warning("No current spreadsheet ID is set")
        else:
            logger.info(f"Using current spreadsheet ID: {current_spreadsheet_id}")

    return sheets.create_chart(chart_type,range_name,sheet_name,spreadsheet_id,title)

@mcp.tool()
def create_presentation(title: str) -> Dict[str, Any]:
    """Create a new google slides presentation.
    
    Args:
        title: Title of the new presentation
    """
    presentation_id = slides.create_presentation(title)
    if presentation_id:
        # Get the first slide's dimensions
        dimensions = slides.get_slide_dimensions(presentation_id, "p")
        return {
            "success": True,
            "presentation_id": presentation_id,
            "message": f"Created presentation: {title}",
            "dimensions": dimensions
        }
    return {
        "success": False,
        "message": "Failed to create presentation"
    }

@mcp.tool()
def add_slide_to_presentation(presentation_id: str, title: str, content: str) -> Dict[str, Any]:
    """Add a new slide to a google slides presentation."""
    try:
        # Handle content formatting
        if isinstance(content, str):
            # If content is a JSON string, parse it
            try:
                # Remove backticks if present
                content = content.strip('`')
                # Parse the content as JSON
                parsed_content = json.loads(content)
                # Convert back to string with proper newlines
                content = parsed_content
            except json.JSONDecodeError:
                # If not valid JSON, use the content as is
                pass

            # Replace any remaining escaped newlines
            content = content.replace('\\n', '\n')
            
        result = slides.add_slide(presentation_id, title, content)
        if result and result.get('success'):
            return {
                "success": True,
                "message": f"Added slide: {title}",
                "slide_id": result.get('slide_id'),
                "dimensions": result.get('dimensions')
            }
        return {
            "success": False,
            "message": "Failed to add slide"
        }
    except Exception as e:
        logger.error(f"Error adding slide: {str(e)}")
        return {
            "success": False,
            "message": f"Error: {str(e)}"
        }

@mcp.tool()
def add_image_to_slide(presentation_id: str, slide_id: str, image_url: str,
                      x: float = 100, y: float = 100,
                      width: float = 400, height: float = 300,
                      rotation: float = 0.0) -> Dict[str, Any]:
    """Add an image to a specific slide of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide to add the image to
        image_url (str): URL of the image to add
        x (float, optional): X position in points (default: 100)
        y (float, optional): Y position in points (default: 100)
        width (float, optional): Width in points (default: 400)
        height (float, optional): Height in points (default: 300)
        rotation (float, optional): Rotation angle in degrees (default: 0)
        
    Returns:
        Dict[str, Any]: Response containing success status and image ID if successful
    """
    try:
        image_id = slides.add_image(
            presentation_id=presentation_id,
            slide_id=slide_id,
            image_url=image_url,
            x=x,
            y=y,
            width=width,
            height=height,
            rotation=rotation
        )
        
        if image_id:
            return {
                'success': True,
                'message': 'Image added successfully',
                'image_id': image_id
            }
        else:
            return {
                'success': False,
                'message': 'Failed to add image'
            }
    except Exception as e:
        logger.error(f"Error adding image: {str(e)}")
        return {
            'success': False,
            'message': f'Error adding image: {str(e)}'
        }

@mcp.tool()
def get_presentation_details(presentation_id: str) -> Dict[str, Any]:
    """Get presentation details of a google slides presentation.
    
    Args:
        presentation_id: ID of the presentation
    """
    presentation = slides.get_presentation(presentation_id)
    if presentation:
        return {
            "success": True,
            "presentation": presentation
        }
    return {
        "success": False,
        "message": "Failed to get presentation details"
    }

@mcp.tool()
def delete_presentation(presentation_id: str) -> Dict[str, Any]:
    """Delete a google slides presentation.
    
    Args:
        presentation_id: ID of the presentation
    """
    success = slides.delete_presentation(presentation_id)
    if success:
        return {
            "success": True,
            "message": "Deleted presentation"
        }
    return {
        "success": False,
        "message": "Failed to delete presentation"
    }

@mcp.tool()
def create_document(title: str) -> Dict[str, Any]:
    """Create a new google docs document.
    
    Args:
        title: Title of the new document
    """
    document_id = docs.create_document(title)
    if document_id:
        return {
            "success": True,
            "document_id": document_id,
            "message": f"Created document: {title}"
        }
    return {
        "success": False,
        "message": "Failed to create document"
    }

@mcp.tool()
def insert_text_to_document(document_id: str, text: str, index: int = 1,
                          font_family: str = None, font_size: float = None,
                          bold: bool = None, italic: bool = None,
                          underline: bool = None, strikethrough: bool = None,
                          foreground_color: str = None, background_color: str = None,
                          alignment: str = None, line_spacing: float = None,
                          space_before: float = None, space_after: float = None,
                          first_line_indent: float = None, bullet: bool = None,
                          numbered_list: bool = None) -> Dict[str, Any]:
    """Insert text into a google docs document with formatting options.
    
    Args:
        document_id: ID of the document
        text: Text to insert
        index: Index of the text to insert
        font_family: Font family name
        font_size: Font size in points
        bold: Whether to make text bold
        italic: Whether to make text italic
        underline: Whether to underline text
        strikethrough: Whether to strikethrough text
        foreground_color: Text color in hex format (e.g., '#FF0000')
        background_color: Background color in hex format (e.g., '#FFFF00')
        alignment: Text alignment ('START', 'CENTER', 'END', 'JUSTIFIED')
        line_spacing: Line spacing multiplier
        space_before: Space before paragraph in points
        space_after: Space after paragraph in points
        first_line_indent: First line indent in points
        bullet: Whether to add bullet points
        numbered_list: Whether to add numbered list
    """
    success = docs.insert_text(
        document_id=document_id,
        text=text,
        index=index,
        font_family=font_family,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        strikethrough=strikethrough,
        foreground_color=foreground_color,
        background_color=background_color,
        alignment=alignment,
        line_spacing=line_spacing,
        space_before=space_before,
        space_after=space_after,
        first_line_indent=first_line_indent,
        bullet=bullet,
        numbered_list=numbered_list
    )
    if success:
        return {
            "success": True,
            "message": "Inserted text into document with formatting"
        }
    return {
        "success": False,
        "message": "Failed to insert text"
    }

@mcp.tool()
def insert_heading_to_document(document_id: str, text: str, level: int = 1, index: int = 1) -> Dict[str, Any]:
    """Insert a heading into a google docs document.
    
    Args:
        document_id: ID of the document
        text: Text to insert
        level: Level of the heading
    """
    success = docs.insert_heading(document_id, text, level, index)
    if success:
        return {
            "success": True,
            "message": f"Inserted heading level {level}"
        }
    return {
        "success": False,
        "message": "Failed to insert heading"
    }

@mcp.tool()
def insert_image_to_document(document_id: str, image_url: str, index: int = 1) -> Dict[str, Any]:
    """Insert an image into a google docs document.
    
    Args:
        document_id: ID of the document
        image_url: URL of the image
        index: Index of the image to insert
    """
    success = docs.insert_image(document_id, image_url, index)
    if success:
        return {
            "success": True,
            "message": "Inserted image into document"
        }
    return {
        "success": False,
        "message": "Failed to insert image"
    }

@mcp.tool()
def get_document_details(document_id: str) -> Dict[str, Any]:
    """Get document details of a google docs document.
    
    Args:
        document_id: ID of the document
    """
    document = docs.get_document(document_id)
    if document:
        return {
            "success": True,
            "document": document
        }
    return {
        "success": False,
        "message": "Failed to get document details"
    }

@mcp.tool()
def delete_document(document_id: str) -> Dict[str, Any]:
    """Delete a google docs document."""
    success = docs.delete_document(document_id)
    if success:
        return {
            "success": True,
            "message": "Deleted document"
        }
    return {
        "success": False,
        "message": "Failed to delete document"
    }

@mcp.tool()
def create_table_in_document(document_id: str, rows: int, columns: int, index: int = 1) -> Dict[str, Any]:
    """Create a table in a google docs document.
    
    Args:
        document_id: ID of the document
        rows: Number of rows in the table
        columns: Number of columns in the table
        index: Index of the table to create
    """
    success = docs.create_table(document_id, rows, columns, index)
    if success:
        return {
            "success": True,
            "message": f"Created {rows}x{columns} table"
        }
    return {
        "success": False,
        "message": "Failed to create table"
    }

@mcp.tool()
def search_slide_elements(presentation_id: str, slide_id: str, element_type: str = None) -> Dict[str, Any]:
    """Search for elements in a specific slide of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide to search in
        element_type (str, optional): Type of elements to search for ('shape', 'text', etc.)
    
    Returns:
        dict: {
            'success': bool,
            'elements': list of dict (if success is True),
            'message': str (if success is False)
        }
    """
    try:
        elements = slides.search_elements(presentation_id, slide_id, element_type)
        return {
            'success': True,
            'elements': elements
        }
    except Exception as e:
        logger.error(f"Error searching elements: {str(e)}")
        return {
            'success': False,
            'message': str(e)
        }

@mcp.tool()
def update_text_style(presentation_id: str, slide_id: str, element_id: str,
                     font_family: str = None, font_size: float = None,
                     font_weight: str = None, font_style: str = None,
                     foreground_color: str = None, background_color: str = None) -> Dict[str, Any]:
    """Update text style of an element in a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide containing the element
        element_id (str): ID of the text element to update
        font_family (str, optional): Font family name (e.g., 'Arial', 'Times New Roman')
        font_size (float, optional): Font size in points
        font_weight (str, optional): Font weight ('NORMAL', 'BOLD', 'LIGHT', etc.)
        font_style (str, optional): Font style ('NORMAL', 'ITALIC')
        foreground_color (str, optional): Text color in hex format (e.g., '#FF0000')
        background_color (str, optional): Background color in hex format (e.g., '#FFFFFF')
    
    Returns:
        dict: {
            'success': bool,
            'message': str
        }
    """
    try:
        success = slides.update_text_style(
            presentation_id=presentation_id,
            slide_id=slide_id,
            element_id=element_id,
            font_family=font_family,
            font_size=font_size,
            font_weight=font_weight,
            font_style=font_style,
            foreground_color=foreground_color,
            background_color=background_color
        )

        return {
            'success': success,
            'message': 'Text style updated successfully' if success else 'Failed to update text style'
        }
    except Exception as e:
        logger.error(f"Error updating text style: {str(e)}")
        return {
            'success': False,
            'message': str(e)
        }

@mcp.tool()
def update_shape_style(presentation_id: str, slide_id: str, element_id: str,
                      width: float = None, height: float = None,
                      x: float = None, y: float = None,
                      fill_color: str = None, border_color: str = None,
                      border_width: float = None) -> Dict[str, Any]:
    """Update shape style (size, position, colors, border) of an element in a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide containing the element
        element_id (str): ID of the shape element to update
        width (float, optional): Width of the shape in points
        height (float, optional): Height of the shape in points
        x (float, optional): X position of the shape in points
        y (float, optional): Y position of the shape in points
        fill_color (str, optional): Fill color in hex format (e.g., '#FF0000')
        border_color (str, optional): Border color in hex format (e.g., '#000000')
        border_width (float, optional): Border width in points
    
    Returns:
        dict: {
            'success': bool,
            'message': str
        }
    """
    try:
        success = slides.update_shape_style(
            presentation_id=presentation_id,
            slide_id=slide_id,
            element_id=element_id,
            width=width,
            height=height,
            x=x,
            y=y,
            fill_color=fill_color,
            border_color=border_color,
            border_width=border_width
        )

        return {
            'success': success,
            'message': 'Shape style updated successfully' if success else 'Failed to update shape style'
        }
    except Exception as e:
        logger.error(f"Error updating shape style: {str(e)}")
        return {
            'success': False,
            'message': str(e)
        }

@mcp.tool()
def delete_slide_element(presentation_id: str, slide_id: str, element_id: str) -> Dict[str, Any]:
    """Delete an element from a slide of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide containing the element
        element_id (str): ID of the element to delete
    
    Returns:
        dict: {
            'success': bool,
            'message': str
        }
    """
    try:
        success = slides.delete_element(presentation_id, slide_id, element_id)

        return {
            'success': success,
            'message': 'Element deleted successfully' if success else 'Failed to delete element'
        }
    except Exception as e:
        logger.error(f"Error deleting element: {str(e)}")
        return {
            'success': False,
            'message': str(e)
        }

@mcp.tool()
def add_shape_to_slide(presentation_id: str, slide_id: str, shape_type: str, x: float, y: float,
                      width: float, height: float, fill_color: str = None,
                      border_color: str = None, border_width: float = None) -> Dict[str, Any]:
    """Add a shape to a slide of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide to add the shape to
        shape_type (str): Type of shape ('RECTANGLE', 'TRIANGLE', 'ELLIPSE', etc.)
        x (float): X position in points
        y (float): Y position in points
        width (float): Width in points
        height (float): Height in points
        fill_color (str, optional): Fill color in hex format (e.g., '#FF0000')
        border_color (str, optional): Border color in hex format (e.g., '#000000')
        border_width (float, optional): Border width in points
        
    Returns:
        Dict[str, Any]: Response containing success status and shape ID if successful
    """
    try:
        shape_id = slides.add_shape(
            presentation_id=presentation_id,
            slide_id=slide_id,
            shape_type=shape_type,
            x=x,
            y=y,
            width=width,
            height=height,
            fill_color=fill_color,
            border_color=border_color,
            border_width=border_width
        )
        
        if shape_id:
            return {
                'success': True,
                'message': 'Shape added successfully',
                'shape_id': shape_id
            }
        else:
            return {
                'success': False,
                'message': 'Failed to add shape'
            }
    except Exception as e:
        return {
            'success': False,
            'message': f'Error adding shape: {str(e)}'
        }

@mcp.tool()
def add_line_to_slide(presentation_id: str, slide_id: str, start_x: float, start_y: float,
                     end_x: float, end_y: float, line_color: str = '#000000',
                     line_width: float = 1.0, line_type: str = 'STRAIGHT') -> Dict[str, Any]:
    """Add a line to a slide of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide to add the line to
        start_x (float): Starting X position in points
        start_y (float): Starting Y position in points
        end_x (float): Ending X position in points
        end_y (float): Ending Y position in points
        line_color (str, optional): Line color in hex format (e.g., '#000000')
        line_width (float, optional): Line width in points
        line_type (str, optional): Type of line ('STRAIGHT', 'CURVED', 'ELBOW', 'BENT')
        
    Returns:
        Dict[str, Any]: Response containing success status and line ID if successful
    """
    try:
        line_id = slides.add_line(
            presentation_id=presentation_id,
            slide_id=slide_id,
            start_x=start_x,
            start_y=start_y,
            end_x=end_x,
            end_y=end_y,
            line_color=line_color,
            line_width=line_width,
            line_type=line_type
        )
        
        if line_id:
            return {
                'success': True,
                'message': 'Line added successfully',
                'line_id': line_id
            }
        else:
            return {
                'success': False,
                'message': 'Failed to add line'
            }
    except Exception as e:
        return {
            'success': False,
            'message': f'Error adding line: {str(e)}'
        }

@mcp.tool()
def update_slide_background(presentation_id: str, slide_id: str,
                          background_color: str = None,
                          background_image_url: str = None) -> Dict[str, Any]:
    """Update slide background with color or image of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide to update
        background_color (str, optional): Background color in hex format (e.g., '#FFFFFF')
        background_image_url (str, optional): URL of the background image
        
    Returns:
        Dict[str, Any]: Response containing success status
    """
    try:
        success = slides.update_slide_background(
            presentation_id=presentation_id,
            slide_id=slide_id,
            background_color=background_color,
            background_image_url=background_image_url
        )
        
        return {
            'success': success,
            'message': 'Slide background updated successfully' if success else 'Failed to update slide background'
        }
    except Exception as e:
        logger.error(f"Error updating slide background: {str(e)}")
        return {
            'success': False,
            'message': str(e)
        }

@mcp.tool()
def update_slide_layout(presentation_id: str, slide_id: str,
                       layout_type: str) -> Dict[str, Any]:
    """Update slide layout of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide to update
        layout_type (str): Type of layout ('TITLE', 'TITLE_AND_BODY', 'MAIN_POINT', etc.)
        
    Returns:
        Dict[str, Any]: Response containing success status
    """
    try:
        success = slides.update_slide_layout(
            presentation_id=presentation_id,
            slide_id=slide_id,
            layout_type=layout_type
        )
        
        return {
            'success': success,
            'message': 'Slide layout updated successfully' if success else 'Failed to update slide layout'
        }
    except Exception as e:
        logger.error(f"Error updating slide layout: {str(e)}")
        return {
            'success': False,
            'message': str(e)
        }

@mcp.tool()
def update_slide_transition(presentation_id: str, slide_id: str,
                          transition_type: str = 'FADE',
                          duration: str = 'SLOW') -> Dict[str, Any]:
    """Update slide transition effect of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide to update
        transition_type (str): Type of transition ('FADE', 'SLIDE', 'ZOOM', etc.)
        duration (str): Duration of transition ('SLOW', 'MEDIUM', 'FAST')
        
    Returns:
        Dict[str, Any]: Response containing success status
    """
    try:
        success = slides.update_slide_transition(
            presentation_id=presentation_id,
            slide_id=slide_id,
            transition_type=transition_type,
            duration=duration
        )
        
        return {
            'success': success,
            'message': 'Slide transition updated successfully' if success else 'Failed to update slide transition'
        }
    except Exception as e:
        logger.error(f"Error updating slide transition: {str(e)}")
        return {
            'success': False,
            'message': str(e)
        }

@mcp.tool()
def add_slide_notes(presentation_id: str, slide_id: str,
                   notes_text: str) -> Dict[str, Any]:
    """Add or update speaker notes for a slide of a google slides presentation.
    
    Args:
        presentation_id (str): Google Slides presentation ID
        slide_id (str): ID of the slide to update
        notes_text (str): Text for speaker notes
        
    Returns:
        Dict[str, Any]: Response containing success status
    """
    try:
        success = slides.add_slide_notes(
            presentation_id=presentation_id,
            slide_id=slide_id,
            notes_text=notes_text
        )
        
        return {
            'success': success,
            'message': 'Slide notes updated successfully' if success else 'Failed to update slide notes'
        }
    except Exception as e:
        logger.error(f"Error updating slide notes: {str(e)}")
        return {
            'success': False,
            'message': str(e)
        }

@mcp.tool()
def update_text_style_in_document(document_id: str, start_index: int, end_index: int,
                                font_family: str = None, font_size: float = None,
                                bold: bool = None, italic: bool = None,
                                underline: bool = None, strikethrough: bool = None,
                                foreground_color: str = None, background_color: str = None) -> Dict[str, Any]:
    """Update text style in a google docs document.
    
    Args:
        document_id: ID of the document
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
    success = docs.update_text_style(
        document_id=document_id,
        start_index=start_index,
        end_index=end_index,
        font_family=font_family,
        font_size=font_size,
        bold=bold,
        italic=italic,
        underline=underline,
        strikethrough=strikethrough,
        foreground_color=foreground_color,
        background_color=background_color
    )
    
    if success:
        return {
            "success": True,
            "message": "Text style updated successfully"
        }
    return {
        "success": False,
        "message": "Failed to update text style"
    }

@mcp.tool()
def update_paragraph_style_in_document(document_id: str, start_index: int, end_index: int,
                                     alignment: str = None, line_spacing: float = None,
                                     space_before: float = None, space_after: float = None,
                                     first_line_indent: float = None, bullet: bool = None,
                                     numbered_list: bool = None) -> Dict[str, Any]:
    """Update paragraph style in a google docs document.
    
    Args:
        document_id: ID of the document
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
    success = docs.update_paragraph_style(
        document_id=document_id,
        start_index=start_index,
        end_index=end_index,
        alignment=alignment,
        line_spacing=line_spacing,
        space_before=space_before,
        space_after=space_after,
        first_line_indent=first_line_indent,
        bullet=bullet,
        numbered_list=numbered_list
    )
    
    if success:
        return {
            "success": True,
            "message": "Paragraph style updated successfully"
        }
    return {
        "success": False,
        "message": "Failed to update paragraph style"
    }

@mcp.tool()
def insert_page_break_in_document(document_id: str, index: int) -> Dict[str, Any]:
    """Insert a page break in a google docs document.
    
    Args:
        document_id: ID of the document
        index: Index where to insert the page break
    """
    success = docs.insert_page_break(document_id, index)
    
    if success:
        return {
            "success": True,
            "message": "Page break inserted successfully"
        }
    return {
        "success": False,
        "message": "Failed to insert page break"
    }

@mcp.tool()
def insert_horizontal_rule_in_document(document_id: str, index: int) -> Dict[str, Any]:
    """Insert a horizontal rule in a google docs document.
    
    Args:
        document_id: ID of the document
        index: Index where to insert the horizontal rule
    """
    success = docs.insert_horizontal_rule(document_id, index)
    
    if success:
        return {
            "success": True,
            "message": "Horizontal rule inserted successfully"
        }
    return {
        "success": False,
        "message": "Failed to insert horizontal rule"
    }

@mcp.tool()
def update_table_cell_content(document_id: str, table_id: str, row_index: int, column_index: int, content: str) -> Dict[str, Any]:
    """Update content of a specific table cell.
    
    Args:
        document_id: Document ID
        table_id: Table ID
        row_index: Row index (0-based)
        column_index: Column index (0-based)
        content: Content to insert
    """
    success = docs.update_table_cell_content(
        document_id=document_id,
        table_id=table_id,
        row_index=row_index,
        column_index=column_index,
        content=content
    )

    if success:
        return {
            'success': True,
            'message': 'Updated table cell content'
        }
    return {
        'success': False,
        'message': 'Failed to update table cell content'
    }

@mcp.tool()
def update_table_cell_style(document_id: str, table_id: str, row_index: int, column_index: int,
                          background_color: str = None, border_color: str = None,
                          border_width: float = None, padding: Dict[str, float] = None) -> Dict[str, Any]:
    """Update style of a specific table cell in a google docs document.
    
    Args:
        document_id: ID of the document
        table_id: Table ID
        row_index: Row index (0-based)
        column_index: Column index (0-based)
        background_color: Background color in hex format
        border_color: Border color in hex format
        border_width: Border width in points
        padding: Dictionary with 'top', 'right', 'bottom', 'left' padding values
    """
    success = docs.update_table_cell_style(
        document_id=document_id,
        table_id=table_id,
        row_index=row_index,
        column_index=column_index,
        background_color=background_color,
        border_color=border_color,
        border_width=border_width,
        padding=padding
    )
    
    if success:
        return {
            "success": True,
            "message": "Table cell style updated successfully"
        }
    return {
        "success": False,
        "message": "Failed to update table cell style"
    }

@mcp.tool()
def update_table_row_style(document_id: str, table_id: str, row_index: int,
                         background_color: str = None, height: float = None) -> Dict[str, Any]:
    """Update style of a table row in a google docs document.
    
    Args:
        document_id: ID of the document
        table_id: Table ID
        row_index: Row index (0-based)
        background_color: Background color in hex format
        height: Row height in points
    """
    success = docs.update_table_row_style(
        document_id=document_id,
        table_id=table_id,
        row_index=row_index,
        background_color=background_color,
        height=height
    )
    
    if success:
        return {
            "success": True,
            "message": "Table row style updated successfully"
        }
    return {
        "success": False,
        "message": "Failed to update table row style"
    }

@mcp.tool()
def update_table_column_style(document_id: str, table_id: str, column_index: int,
                            width: float = None) -> Dict[str, Any]:
    """Update style of a table column in a google docs document.
    
    Args:
        document_id: ID of the document
        table_id: Table ID
        column_index: Column index (0-based)
        width: Column width in points
    """
    success = docs.update_table_column_style(
        document_id=document_id,
        table_id=table_id,
        column_index=column_index,
        width=width
    )
    
    if success:
        return {
            "success": True,
            "message": "Table column style updated successfully"
        }
    return {
        "success": False,
        "message": "Failed to update table column style"
    }

@mcp.tool()
def merge_table_cells(document_id: str, table_id: str, start_row: int, start_column: int,
                     end_row: int, end_column: int) -> Dict[str, Any]:
    """Merge table cells in a google docs document.
    
    Args:
        document_id: ID of the document
        table_id: Table ID
        start_row: Start row index (0-based)
        start_column: Start column index (0-based)
        end_row: End row index (0-based)
        end_column: End column index (0-based)
    """
    success = docs.merge_table_cells(
        document_id=document_id,
        table_id=table_id,
        start_row=start_row,
        start_column=start_column,
        end_row=end_row,
        end_column=end_column
    )
    
    if success:
        return {
            "success": True,
            "message": "Table cells merged successfully"
        }
    return {
        "success": False,
        "message": "Failed to merge table cells"
    }

@mcp.tool()
def insert_table_row(document_id: str, table_id: str, row_index: int, num_rows: int = 1) -> Dict[str, Any]:
    """Insert rows into a table in a google docs document.
    
    Args:
        document_id: ID of the document
        table_id: Table ID
        row_index: Index where to insert the row(s)
        num_rows: Number of rows to insert
    """
    success = docs.insert_table_row(
        document_id=document_id,
        table_id=table_id,
        row_index=row_index,
        num_rows=num_rows
    )
    
    if success:
        return {
            "success": True,
            "message": "Table rows inserted successfully"
        }
    return {
        "success": False,
        "message": "Failed to insert table rows"
    }

@mcp.tool()
def insert_table_column(document_id: str, table_id: str, column_index: int, num_columns: int = 1) -> Dict[str, Any]:
    """Insert columns into a table in a google docs document.
    
    Args:
        document_id: ID of the document
        table_id: Table ID
        column_index: Index where to insert the column(s)
        num_columns: Number of columns to insert
    """
    success = docs.insert_table_column(
        document_id=document_id,
        table_id=table_id,
        column_index=column_index,
        num_columns=num_columns
    )
    
    if success:
        return {
            "success": True,
            "message": "Table columns inserted successfully"
        }
    return {
        "success": False,
        "message": "Failed to insert table columns"
    }

@mcp.tool()
def delete_table_row(document_id: str, table_id: str, row_index: int, num_rows: int = 1) -> Dict[str, Any]:
    """Delete rows from a table in a google docs document.
    
    Args:
        document_id: ID of the document
        table_id: Table ID
        row_index: Index of the first row to delete
        num_rows: Number of rows to delete
    """
    success = docs.delete_table_row(
        document_id=document_id,
        table_id=table_id,
        row_index=row_index,
        num_rows=num_rows
    )
    
    if success:
        return {
            "success": True,
            "message": "Table rows deleted successfully"
        }
    return {
        "success": False,
        "message": "Failed to delete table rows"
    }

@mcp.tool()
def delete_table_column(document_id: str, table_id: str, column_index: int, num_columns: int = 1) -> Dict[str, Any]:
    """Delete columns from a table in a google docs document.
    
    Args:
        document_id: ID of the document
        table_id: Table ID
        column_index: Index of the first column to delete
        num_columns: Number of columns to delete
    """
    success = docs.delete_table_column(
        document_id=document_id,
        table_id=table_id,
        column_index=column_index,
        num_columns=num_columns
    )
    
    if success:
        return {
            "success": True,
            "message": "Table columns deleted successfully"
        }
    return {
        "success": False,
        "message": "Failed to delete table columns"
    }

@mcp.tool()
def update_document_style(document_id: str,
                        default_font_family: str = None,
                        default_font_size: float = None,
                        default_line_spacing: float = None,
                        default_margin_top: float = None,
                        default_margin_bottom: float = None,
                        default_margin_left: float = None,
                        default_margin_right: float = None,
                        default_page_color: str = None) -> Dict[str, Any]:
    """Update document-wide styles in a google docs document.
    
    Args:
        document_id: ID of the document
        default_font_family: Default font family for the document
        default_font_size: Default font size in points
        default_line_spacing: Default line spacing multiplier
        default_margin_top: Top margin in points
        default_margin_bottom: Bottom margin in points
        default_margin_left: Left margin in points
        default_margin_right: Right margin in points
        default_page_color: Page background color in hex format
    """
    success = docs.update_document_style(
        document_id=document_id,
        default_font_family=default_font_family,
        default_font_size=default_font_size,
        default_line_spacing=default_line_spacing,
        default_margin_top=default_margin_top,
        default_margin_bottom=default_margin_bottom,
        default_margin_left=default_margin_left,
        default_margin_right=default_margin_right,
        default_page_color=default_page_color
    )
    
    if success:
        return {
            "success": True,
            "message": "Document style updated successfully"
        }
    return {
        "success": False,
        "message": "Failed to update document style"
    }

@mcp.tool()
def create_table_with_content(document_id: str, rows: int, columns: int, content: Dict[str, str]) -> Dict[str, Any]:
    """Create a table and insert text into specific cells.
    
    Args:
        document_id: Document ID
        rows: Number of rows in the table
        columns: Number of columns in the table
        content: Dictionary mapping cell coordinates to content
               Format: {'row,column': 'content'}
               Example: {'0,0': 'Header', '1,1': 'Data'}
    """
    success = docs.create_table_and_insert_text(document_id, rows, columns, content)
    if success:
        return {
            "success": True,
            "message": "Created table and inserted content successfully"
        }
    return {
        "success": False,
        "message": "Failed to create table and insert content"
    }

def signal_handler(signum, frame):
    """Signal handler to handle SIGINT and SIGTERM."""
    logger.info("Received signal to terminate")
    sys.exit(0)

if __name__ == "__main__":
    # Register signal handler
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    
    # Run MCP server
    mcp.run() 