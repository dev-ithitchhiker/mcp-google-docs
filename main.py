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

# 경고 메시지 무시
warnings.filterwarnings('ignore', message='file_cache is only supported with oauth2client<4.0.0')

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# googleapiclient의 로깅 레벨을 WARNING으로 설정
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.WARNING)

# MCP 서버 초기화
mcp = FastMCP("Google Spreadsheet MCP")

# 설정 및 인증 초기화
config = Config.from_env()
auth = GoogleAuth(config)
drive = GoogleDrive(auth)
sheets = GoogleSheets(auth)

# 전역 변수로 현재 작업 중인 스프레드시트 ID 저장
current_spreadsheet_id = None

@mcp.tool()
def list_files() -> List[Dict[str, Any]]:
    """Google Drive의 파일 목록을 조회합니다."""
    return drive.list_files()

@mcp.tool()
def copy_file(file_id: str, new_name: str) -> Dict[str, Any]:
    """파일을 복사합니다."""
    return drive.copy_file(file_id, new_name)

@mcp.tool()
def rename_file(file_id: str, new_name: str) -> Dict[str, Any]:
    """파일 이름을 변경합니다."""
    return drive.rename_file(file_id, new_name)

@mcp.tool()
def create_spreadsheet(title: str) -> Dict[str, Any]:
    """빈 스프레드시트를 생성합니다."""
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
    """템플릿을 기반으로 새 스프레드시트를 생성합니다."""
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
    """기존 스프레드시트를 복사하여 새 스프레드시트를 생성합니다."""
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
    """스프레드시트의 시트 목록을 조회합니다."""
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
    """새 시트를 생성합니다."""
    global current_spreadsheet_id
    return sheets.add_sheet(spreadsheet_id, sheet_name)

@mcp.tool()
def duplicate_sheet(spreadsheet_id: str, sheet_id: int, new_name: str) -> Dict[str, Any]:
    """기존 시트를 복제하여 새 시트를 생성합니다.
    
    Args:
        values: 시트 데이터
        range_name: 데이터 범위 (예: 'A1:B5')
        sheet_name: 시트 이름
        spreadsheet_id: 스프레드시트 ID
        source_sheet_id: 복제할 원본 시트 ID
        new_name: 새 시트 이름
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
    """시트 이름을 변경합니다."""
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
    """시트의 데이터를 조회합니다."""
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
    """시트에 행을 추가합니다."""
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
    """시트에 열을 추가합니다."""
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
    """시트의 셀을 업데이트합니다. HTML 태그와 서식을 지원합니다.
        
        Args:
            values: 업데이트할 값들
            range_name: 범위 (예: 'A1:B5')
            sheet_name: 시트 이름
            spreadsheet_id: 스프레드시트 ID
            format: 서식 설정 정보. 다음 형식을 가집니다:
                {
                    'textFormat': {  # 텍스트 서식
                        'fontFamily': str,  # 글꼴
                        'fontSize': int,  # 글자 크기
                        'bold': bool,  # 굵게
                        'italic': bool,  # 기울임
                        'strikethrough': bool,  # 취소선
                        'underline': bool,  # 밑줄
                        'foregroundColor': {  # 글자 색상
                            'red': float,  # 0.0 ~ 1.0
                            'green': float,
                            'blue': float,
                            'alpha': float
                        }
                    },
                    'backgroundColor': {  # 배경 색상
                        'red': float,
                        'green': float,
                        'blue': float,
                        'alpha': float
                    },
                    'horizontalAlignment': str,  # 가로 정렬 ('LEFT', 'CENTER', 'RIGHT')
                    'verticalAlignment': str,  # 세로 정렬 ('TOP', 'MIDDLE', 'BOTTOM')
                    'padding': {  # 여백
                        'top': int,  # 위쪽 여백
                        'right': int,  # 오른쪽 여백
                        'bottom': int,  # 아래쪽 여백
                        'left': int  # 왼쪽 여백
                    },
                    'wrapText': bool,  # 자동 줄바꿈
                    'textRotation': {  # 텍스트 회전
                        'angle': int  # 회전 각도 (0 ~ 360)
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
    """시트의 여러 셀을 일괄 업데이트합니다. 
        HTML 태그를 지원합니다.
        
        Args:
            updates: 업데이트할 값들
            sheet_name: 시트 이름
            spreadsheet_id: 스프레드시트 ID
            format: 서식 설정 정보. 다음 형식을 가집니다:
                {
                    'textFormat': {  # 텍스트 서식
                        'fontFamily': str,  # 글꼴
                        'fontSize': int,  # 글자 크기
                        'bold': bool,  # 굵게
                        'italic': bool,  # 기울임
                        'strikethrough': bool,  # 취소선
                        'underline': bool,  # 밑줄
                        'foregroundColor': {  # 글자 색상
                            'red': float,  # 0.0 ~ 1.0
                            'green': float,
                            'blue': float,
                            'alpha': float
                        },
                        'backgroundColor': {  # 배경 색상
                            'red': float,
                            'green': float,
                            'blue': float,
                            'alpha': float
                        },
                        'horizontalAlignment': str,  # 가로 정렬 ('LEFT', 'CENTER', 'RIGHT')
                        'verticalAlignment': str,  # 세로 정렬 ('TOP', 'MIDDLE', 'BOTTOM')
                        'padding': {  # 여백
                            'top': int,  # 위쪽 여백
                            'right': int,  # 오른쪽 여백
                            'bottom': int,  # 아래쪽 여백
                            'left': int  # 왼쪽 여백
                        },
                        'wrapText': bool,  # 자동 줄바꿈
                        'textRotation': {  # 텍스트 회전
                            'angle': int  # 회전 각도 (0 ~ 360)
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
    """시트의 행을 삭제합니다."""
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
    """시트의 열을 삭제합니다.

    Args:
        spreadsheet_id: 스프레드시트 ID
        sheet_name: 시트 이름
        start_index: 삭제할 열의 시작 인덱스
        end_index: 삭제할 열의 끝 인덱스
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
    """차트를 생성합니다.
        
        Args:
            chart_type: 차트 유형 ('LINE', 'COLUMN', 'PIE', 'SCATTER', 'BAR')
            range_name: 데이터 범위 (예: 'A1:B10')
            sheet_name: 시트 이름
            spreadsheet_id: 스프레드시트 ID
            title: 차트 제목 (선택사항)
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
    """시그널 핸들러"""
    logger.info("Received signal to terminate")
    sys.exit(0)

if __name__ == "__main__":
    # 시그널 핸들러 등록
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    
    # MCP 서버 실행
    mcp.run() 