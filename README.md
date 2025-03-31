# MCP Google Spreadsheet

Google Spreadsheet와 Google Drive를 조작하기 위한 MCP(Metoro Control Protocol) 도구입니다.
A tool for manipulating Google Spreadsheet and Google Drive using MCP (Metoro Control Protocol).

## 기능 (Features)

### Google Drive 기능 (Google Drive Features)
- 파일 목록 조회 (List files)
- 파일 복사 (Copy files)
- 파일 이름 변경 (Rename files)
- 빈 스프레드시트 생성 (Create empty spreadsheets)
- 템플릿 기반 스프레드시트 생성 (Create spreadsheets from templates)
- 기존 스프레드시트 복사 (Copy existing spreadsheets)

### Google Sheets 기능 (Google Sheets Features)
- 시트 목록 조회 (List sheets)
- 시트 복사 (Copy sheets)
- 시트 이름 변경 (Rename sheets)
- 시트 데이터 조회 (Get sheet data)
- 행 추가/삭제 (Add/Delete rows)
- 열 추가/삭제 (Add/Delete columns)
- 셀 업데이트 (Update cells)
- 차트 생성/수정/삭제 (Create/Update/Delete charts)
- 셀 서식 변경 (Update cell formats)

## 설치 방법 (Installation)

### 1. 가상 환경 설정 (Virtual Environment Setup)

#### macOS/Linux
```bash
# 가상 환경 생성 (Create virtual environment)
python -m venv venv

# 가상 환경 활성화 (Activate virtual environment)
source venv/bin/activate
```

#### Windows
```bash
# 가상 환경 생성 (Create virtual environment)
python -m venv venv

# 가상 환경 활성화 (Activate virtual environment)
venv\Scripts\activate
```

### 2. 필요한 패키지 설치 (Install Required Packages):
```bash
pip install -r requirements.txt
```

### 3. Google Cloud Console 설정 (Google Cloud Console Setup)
1. Google Cloud Console에서 프로젝트를 생성합니다. (Create a project in Google Cloud Console)
2. OAuth 2.0 클라이언트 ID를 생성합니다. (Create OAuth 2.0 client ID)
3. 필요한 API를 활성화합니다: (Enable required APIs:)
   - Google Sheets API
   - Google Drive API

### 4. 환경 변수 설정 (Environment Variables Setup):
```bash
export MCPGD_CLIENT_SECRET_PATH="/path/to/client_secret.json"
export MCPGD_FOLDER_ID="your_folder_id"
export MCPGD_TOKEN_PATH="/path/to/token.json"  # 선택사항 (Optional)
```

## 사용 방법 (Usage)

### 1. 프로그램 실행 (Run the Program):
```bash
python main.py
```

### 2. MCP를 통해 도구 사용 (Use Tools via MCP):
```bash
# 예시: 파일 목록 조회 (Example: List files)
mcp list_files

# 예시: 시트 데이터 조회 (Example: Get sheet data)
mcp get_sheet_data --spreadsheet-id "your_spreadsheet_id" --range "Sheet1!A1:D10"

# 예시: 차트 생성 (Example: Create chart)
mcp create_chart --chart-type "LINE" --range "A1:B10" --sheet-name "Sheet1" --title "Sales Trend"
```

## 환경 변수 (Environment Variables)

- `MCPGD_CLIENT_SECRET_PATH`: Google OAuth 2.0 클라이언트 시크릿 파일 경로 (Path to Google OAuth 2.0 client secret file)
- `MCPGD_FOLDER_ID`: Google Drive 폴더 ID (Google Drive folder ID)
- `MCPGD_TOKEN_PATH`: 토큰 저장 파일 경로 (Path to token storage file, Optional, Default: ~/.mcp_google_spreadsheet.json)

## 라이선스 (License)

MIT License
