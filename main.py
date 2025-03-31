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
    """Create an empty spreadsheet."""
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
    """Create a new spreadsheet from a template."""
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
    """Create a new spreadsheet by copying an existing one."""
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
    """List sheets in a spreadsheet."""
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
    """Create a new sheet."""
    global current_spreadsheet_id
    return sheets.add_sheet(spreadsheet_id, sheet_name)

@mcp.tool()
def duplicate_sheet(spreadsheet_id: str, sheet_id: int, new_name: str) -> Dict[str, Any]:
    """Create a new sheet by duplicating an existing one.
    
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
    """Rename a sheet."""
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
    """Get data from a sheet."""
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
    """Add rows to a sheet."""
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
    """Add columns to a sheet."""
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
    """Update cells in a sheet. Supports HTML tags and formatting.
        
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
    """Delete rows from a sheet."""
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
    """Delete columns from a sheet.

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
    """Create a chart.
        
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


def signal_handler(signum, frame):
    """Signal handler"""
    logger.info("Received signal to terminate")
    sys.exit(0)

if __name__ == "__main__":
    # Register signal handler
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    
    # Run MCP server
    mcp.run() 