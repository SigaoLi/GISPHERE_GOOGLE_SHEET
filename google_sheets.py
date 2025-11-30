"""
Google Sheets API 模块 - 处理Google表格的读写操作
"""
import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from config import (
    SCOPES_SHEETS, 
    SPREADSHEET_ID, 
    TOKEN_PICKLE_FILE, 
    CREDENTIALS_FILE
)


def authorize_credentials():
    """授权Google API凭据"""
    creds = None
    if os.path.exists(TOKEN_PICKLE_FILE):
        with open(TOKEN_PICKLE_FILE, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES_SHEETS)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_PICKLE_FILE, 'wb') as token:
            pickle.dump(creds, token)
    
    return creds


def fetch_data(range_name):
    """从Google表格获取数据"""
    creds = authorize_credentials()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    values = result.get('values', [])
    return values


def delete_rows_from_sheet(sheet_id, rows_to_delete):
    """从Google表格中删除行"""
    if not rows_to_delete:
        return
    
    creds = authorize_credentials()
    service = build('sheets', 'v4', credentials=creds)
    
    # 按降序排序，从后往前删除，避免索引变化
    rows_to_delete.sort(reverse=True)
    
    batch_update_body = {
        "requests": [
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": start_index + 1
                    }
                }
            } for start_index in rows_to_delete
        ]
    }
    
    request = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=batch_update_body)
    response = request.execute()
    print(f"{len(rows_to_delete)} rows deleted.")


def append_data_to_sheet(range_name, data):
    """向Google表格追加数据"""
    creds = authorize_credentials()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    body = {'values': data}
    
    result = sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body=body,
        insertDataOption='INSERT_ROWS'
    ).execute()
    
    print(f"{result.get('updates').get('updatedRows')} rows appended.")


def update_data_in_sheet(range_name, data):
    """更新Google表格中指定范围的数据"""
    creds = authorize_credentials()
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    body = {'values': data}
    
    result = sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    
    print(f"{result.get('updatedRows')} rows updated.")

