# MCP Google Workspace Integration

A comprehensive MCP (Metoro Control Protocol) tool for interacting with Google Workspace services including Google Docs, Sheets, Slides, and Drive.

## Features

### Google Drive Features
- List files
- Copy files
- Rename files
- Create empty spreadsheets
- Create spreadsheets from templates
- Copy existing spreadsheets

### Google Sheets Features
- List sheets
- Copy sheets
- Rename sheets
- Get sheet data
- Add/Delete rows
- Add/Delete columns
- Update cells
- Create/Update/Delete charts
- Update cell formats

### Google Docs Features
- Create documents
- Insert text with formatting
- Add headings
- Insert images
- Create and manage tables
- Insert page breaks
- Add horizontal rules
- Update document styles
- Manage table styles and content

### Google Slides Features
- Create presentations
- Add slides
- Insert images
- Add shapes and lines
- Update text styles
- Modify slide backgrounds
- Update slide layouts
- Add slide transitions
- Add speaker notes

## Installation

### 1. Virtual Environment Setup

#### macOS/Linux
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
source venv/bin/activate
```

#### Windows
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate
```

### 2. Install Required Packages
```bash
pip install -r requirements.txt
```

### 3. Google Cloud Console Setup
1. Create a project in Google Cloud Console
2. Create OAuth 2.0 client ID
3. Enable required APIs:
   - Google Sheets API
   - Google Drive API
   - Google Docs API
   - Google Slides API

### 4. Environment Variables Setup
```bash
export MCPGD_CLIENT_SECRET_PATH="/path/to/client_secret.json"
export MCPGD_FOLDER_ID="your_folder_id"
export MCPGD_TOKEN_PATH="/path/to/token.json"  # Optional
```

## Usage

### 1. Run the Program
```bash
python main.py
```

### 2. Use Tools via MCP

#### Google Drive Examples
```bash
# List files
mcp list_files

# Copy a file
mcp copy_file --file-id "file_id" --new_name "new_name"
```

#### Google Sheets Examples
```bash
# Get sheet data
mcp get_sheet_data --spreadsheet_id "your_spreadsheet_id" --range "Sheet1!A1:D10"

# Create chart
mcp create_chart --chart_type "LINE" --range "A1:B10" --sheet_name "Sheet1" --title "Sales Trend"
```

#### Google Docs Examples
```bash
# Create document
mcp create_document --title "My Document"

# Insert formatted text
mcp insert_text_to_document --document_id "doc_id" --text "Hello World" --font_family "Arial" --font_size 12
```

#### Google Slides Examples
```bash
# Create presentation
mcp create_presentation --title "My Presentation"

# Add slide with content
mcp add_slide_to_presentation --presentation_id "presentation_id" --title "Slide Title" --content "Slide Content"
```

## Environment Variables

- `MCPGD_CLIENT_SECRET_PATH`: Path to Google OAuth 2.0 client secret file
- `MCPGD_FOLDER_ID`: Google Drive folder ID
- `MCPGD_TOKEN_PATH`: Path to token storage file (Optional, Default: ~/.mcp_google_spreadsheet.json)

## License

MIT License
