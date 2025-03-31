# MCP Google Spreadsheet

Google Spreadsheet와 Google Drive를 조작하기 위한 MCP(Metoro Control Protocol) 도구입니다.

## 기능

### Google Drive 기능
- 파일 목록 조회 (list_files)
- 파일 복사 (copy_file)
- 파일 이름 변경 (rename_file)

### Google Sheets 기능
- 시트 목록 조회 (list_sheets)
- 시트 복사 (copy_sheet)
- 시트 이름 변경 (rename_sheet)
- 시트 데이터 조회 (get_sheet_data)
- 행 추가/삭제 (add_rows, delete_rows)
- 열 추가/삭제 (add_columns, delete_columns)
- 셀 업데이트 (update_cells, batch_update_cells)

## 설치 방법

1. 필요한 패키지 설치:
```bash
pip install -r requirements.txt
```

2. Google Cloud Console에서 프로젝트를 생성하고 OAuth 2.0 클라이언트 ID를 생성합니다.

3. 환경 변수 설정:
```bash
export MCPGS_CLIENT_SECRET_PATH="/path/to/client_secret.json"
export MCPGS_FOLDER_ID="your_folder_id"
export MCPGS_TOKEN_PATH="/path/to/token.json"  # 선택사항
```

## 사용 방법

1. 프로그램 실행:
```bash
python main.py
```

2. MCP를 통해 도구 사용:
```bash
# 예시: 파일 목록 조회
mcp list_files

# 예시: 시트 데이터 조회
mcp get_sheet_data --spreadsheet-id "your_spreadsheet_id" --range "Sheet1!A1:D10"
```

## 환경 변수

- `MCPGS_CLIENT_SECRET_PATH`: Google OAuth 2.0 클라이언트 시크릿 파일 경로
- `MCPGS_FOLDER_ID`: Google Drive 폴더 ID
- `MCPGS_TOKEN_PATH`: 토큰 저장 파일 경로 (선택사항, 기본값: ~/.mcp_google_spreadsheet.json)

## 라이선스

MIT License
