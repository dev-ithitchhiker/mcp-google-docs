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
        """기존 시트를 복제하여 새 시트를 생성합니다."""
        try:
            logger.info(f"Duplicating sheet {sheet_id} to new sheet with name: {new_name}")
            # 시트 복사 요청
            copy_request = {
                'destinationSpreadsheetId': spreadsheet_id
            }
            
            # 시트 복사 실행
            result = self.service.spreadsheets().sheets().copyTo(
                spreadsheetId=spreadsheet_id,
                sheetId=sheet_id,
                body=copy_request
            ).execute()
            
            # 새 시트의 ID 가져오기
            new_sheet_id = result.get('sheetId')
            if not new_sheet_id:
                logger.error(f"Failed to get new sheet ID after duplicating sheet {sheet_id}")
                return {
                    'success': False,
                    'error': 'Failed to get new sheet ID',
                    'details': result
                }
            
            # 새 시트 이름 변경
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
        """시트 이름을 변경합니다."""
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
        """시트의 데이터를 조회합니다."""
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
        """시트에 행을 추가합니다."""
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
        """시트에 열을 추가합니다."""
        try:
            logger.info(f"Adding columns to sheet '{sheet_name}'")
            # 시트 ID 가져오기
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

            # 열 추가
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
        """시트의 셀을 업데이트합니다."""
        try:
            logger.info(f"Updating cells in sheet '{sheet_name}' range '{range_name}'")
            
            # 시트 ID 가져오기
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
            
            # range_name 파싱
            range_info = self._parse_range(range_name)
            
            # 데이터 차원 검증
            expected_cols = range_info['endColumnIndex'] - range_info['startColumnIndex']
            if any(len(row) != expected_cols for row in values):
                logger.error(f"Data dimensions do not match range. Expected {expected_cols} columns, got {len(values[0])}")
                return {
                    'success': False,
                    'error': f"Data dimensions do not match range. Expected {expected_cols} columns, got {len(values[0])}"
                }
            
            # format 파라미터 처리
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
            
            # 업데이트 요청 생성
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
            
            # batch_update_cells 사용
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
        """HTML 태그를 Google Sheets 스타일로 변환합니다."""
        style = {}
        content = text
        
        # HTML 태그 제거 및 스타일 추출
        while '<' in content and '>' in content:
            start = content.find('<')
            end = content.find('>', start)
            if end == -1:
                break
                
            tag = content[start:end+1]
            content = content[:start] + content[end+1:]
            
            # 닫는 태그 처리
            if tag.startswith('</'):
                continue
                
            # 태그 내용 추출
            tag_name = tag[1:].split()[0].lower()
            tag_attrs = {}
            
            # 속성 추출
            if ' ' in tag:
                attrs = tag[tag.find(' ')+1:-1].split()
                for attr in attrs:
                    if '=' in attr:
                        key, value = attr.split('=', 1)
                        tag_attrs[key.lower()] = value.strip("'\"")
            
            # 스타일 적용
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
        """범위 문자열을 파싱하여 인덱스로 변환합니다."""
        # 예: 'A1:D1' -> {'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 3}
        if ':' in range_str:
            start_range, end_range = range_str.split(':')
            # 열 문자 추출 (첫 번째 문자만 사용)
            start_col = start_range[0].upper()
            start_row = int(start_range[1:]) - 1  # 0-based index
            end_col = end_range[0].upper()
            end_row = int(end_range[1:]) - 1  # 0-based index
            
            # 열 문자를 인덱스로 변환 (A=0, B=1, ...)
            start_col_index = ord(start_col) - ord('A')
            end_col_index = ord(end_col) - ord('A')
        else:
            # 단일 셀 범위 처리
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
        """시트의 여러 셀을 일괄 업데이트합니다. HTML 태그를 지원합니다."""
        try:
            logger.info(f"Batch updating cells in sheet '{sheet_name}'")
            # 시트 ID 가져오기
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in spreadsheet.get('sheets', []):
                if sheet['properties']['title'] == sheet_name:
                    sheet_id = sheet['properties']['sheetId']
                    break
            
            # 시트가 존재하지 않는 경우 생성
            if sheet_id is None:
                logger.info(f"Sheet '{sheet_name}' not found, creating new sheet")
                create_result = self.add_sheet(spreadsheet_id, sheet_name)
                if not create_result['success']:
                    logger.error(f"Failed to create new sheet '{sheet_name}'")
                    return create_result
                sheet_id = create_result['result']['sheet_id']

            # 업데이트 요청 생성
            requests = []
            for update in updates:
                if 'range' in update and 'values' in update:
                    # range가 문자열인 경우 파싱
                    if isinstance(update['range'], str):
                        range_info = self._parse_range(update['range'])
                    else:
                        range_info = update['range']
                    
                    values = update['values']
                    merge = update.get('merge', False)
                    format_config = update.get('format', {})

                    # 셀 업데이트 요청
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

                    # 각 행의 데이터와 스타일 추가
                    for row in values:
                        row_data = {'values': []}
                        for value in row:
                            # 빈 문자열 처리
                            if value == "":
                                cell_data = {'userEnteredValue': None}
                            else:
                                # HTML 태그 파싱
                                parsed = self._parse_html_tags(str(value))
                                cell_data = {'userEnteredValue': {'stringValue': parsed['text']}}
                                
                                # 스타일 적용
                                cell_data['userEnteredFormat'] = {}
                                
                                # 기본 스타일 적용
                                if format_config:
                                    if 'textFormat' in format_config:
                                        cell_data['userEnteredFormat']['textFormat'] = format_config['textFormat']
                                    if 'backgroundColor' in format_config:
                                        cell_data['userEnteredFormat']['backgroundColor'] = format_config['backgroundColor']
                                
                                # HTML 태그 스타일 적용
                                if parsed['style']:
                                    if 'textFormat' not in cell_data['userEnteredFormat']:
                                        cell_data['userEnteredFormat']['textFormat'] = {}
                                    
                                    # 텍스트 스타일
                                    if 'bold' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['bold'] = parsed['style']['bold']
                                    if 'italic' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['italic'] = parsed['style']['italic']
                                    if 'strikethrough' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['strikethrough'] = parsed['style']['strikethrough']
                                    if 'underline' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['underline'] = parsed['style']['underline']
                                    
                                    # 글자 크기
                                    if 'fontSize' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['fontSize'] = parsed['style']['fontSize']
                                    
                                    # 글자 색상
                                    if 'foregroundColor' in parsed['style']:
                                        cell_data['userEnteredFormat']['textFormat']['foregroundColor'] = parsed['style']['foregroundColor']
                                    
                                    # 배경 색상
                                    if 'backgroundColor' in parsed['style']:
                                        cell_data['userEnteredFormat']['backgroundColor'] = parsed['style']['backgroundColor']
                                    
                                    # 정렬
                                    if 'horizontalAlignment' in parsed['style']:
                                        cell_data['userEnteredFormat']['horizontalAlignment'] = parsed['style']['horizontalAlignment']
                            
                            row_data['values'].append(cell_data)
                        cell_request['updateCells']['rows'].append(row_data)

                    requests.append(cell_request)

                    # 셀 병합 요청
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
        """시트의 행을 삭제합니다."""
        try:
            logger.info(f"Deleting rows {start_index} to {end_index} from sheet '{sheet_name}'")
            # 시트 ID 가져오기
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
        """시트의 열을 삭제합니다."""
        try:
            logger.info(f"Deleting columns {start_index} to {end_index} from sheet '{sheet_name}'")
            # 시트 ID 가져오기
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
        """차트를 생성합니다.
        
        Args:
            chart_type: 차트 유형 ('LINE', 'COLUMN', 'PIE', 'SCATTER', 'BAR')
            range_name: 데이터 범위 (예: 'A1:B10')
            sheet_name: 시트 이름
            spreadsheet_id: 스프레드시트 ID
            title: 차트 제목 (선택사항)
        """
        try:
            # sheet_name에서 따옴표 제거
            sheet_name = sheet_name.strip('"')
            logger.info(f"Creating chart in sheet '{sheet_name}'")
            
            # 시트 ID 가져오기
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

            # 범위 파싱
            range_info = self._parse_range(range_name)

            # 차트 요청 생성
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

            # 모든 계열 추가
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
        """기존 차트를 업데이트합니다."""
        try:
            logger.info(f"Updating chart {chart_id} in sheet '{sheet_name}'")
            # 시트 ID 가져오기
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

            # 차트 업데이트 요청 생성
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

            # 차트 옵션 적용
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
        """차트를 삭제합니다."""
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
        """새로운 시트를 추가합니다."""
        try:
            logger.info(f"Adding new sheet '{sheet_name}' to spreadsheet")
            
            # 새 시트 추가 요청
            body = {
                'requests': [{
                    'addSheet': {
                        'properties': {
                            'title': sheet_name,
                            'gridProperties': {
                                'rowCount': 1000,  # 기본 행 수
                                'columnCount': 26  # 기본 열 수 (A-Z)
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
            
            # 새로 생성된 시트의 ID 가져오기
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
        """셀의 서식 스타일을 변경합니다.
        
        Args:
            spreadsheet_id: 스프레드시트 ID
            sheet_name: 시트 이름
            range_name: 범위 (예: 'A1:B5')
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
        try:
            logger.info(f"Updating cell format in sheet '{sheet_name}' range '{range_name}'")
            
            # 시트 ID 가져오기
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

            # 범위 파싱
            range_info = self._parse_range(range_name)
            
            # 서식 업데이트 요청 생성
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

            # 서식 설정 적용
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
    """차트를 생성합니다.
    
    Args:
        chart_type: 차트 유형 ('LINE', 'COLUMN', 'PIE', 'SCATTER', 'BAR')
        range_name: 데이터 범위 (예: 'A1:B10')
        sheet_name: 시트 이름
        spreadsheet_id: 스프레드시트 ID
        title: 차트 제목 (선택사항)
    """
    auth = GoogleAuth()
    sheets = GoogleSheets(auth)
    return sheets.create_chart(chart_type, range_name, sheet_name, spreadsheet_id, title)

def duplicate_sheet(values: List[List[Any]], range_name: str, sheet_name: str, spreadsheet_id: str, source_sheet_id: int, new_name: str) -> Dict[str, Any]:
    """기존 시트를 복제하여 새 시트를 생성합니다.
    
    Args:
        values: 시트 데이터 (사용되지 않음)
        range_name: 데이터 범위 (사용되지 않음)
        sheet_name: 시트 이름 (사용되지 않음)
        spreadsheet_id: 스프레드시트 ID
        source_sheet_id: 복제할 원본 시트 ID
        new_name: 새 시트 이름
    """
    auth = GoogleAuth()
    sheets = GoogleSheets(auth)
    return sheets.duplicate_sheet(spreadsheet_id, source_sheet_id, new_name)

def add_sheet(values: List[List[Any]], range_name: str, sheet_name: str, spreadsheet_id: str) -> Dict[str, Any]:
    """새로운 시트를 추가합니다.
    
    Args:
        values: 시트 데이터 (사용되지 않음)
        range_name: 데이터 범위 (사용되지 않음)
        sheet_name: 시트 이름
        spreadsheet_id: 스프레드시트 ID
    """
    auth = GoogleAuth()
    sheets = GoogleSheets(auth)
    return sheets.add_sheet(spreadsheet_id, sheet_name) 