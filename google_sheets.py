"""
Google Sheets API 模块 - 处理Google表格的读写操作
"""
import os
import pickle
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.exceptions import RefreshError
from google_http import build_google_service, setup_google_proxy_env, refresh_credentials, execute_with_retry
from config import (
    SCOPES_SHEETS,
    SPREADSHEET_ID,
    TOKEN_PICKLE_FILE,
    TOKEN_SHEETS_JSON_FILE,
    CREDENTIALS_FILE
)


def _load_saved_credentials():
    """加载已保存的凭据：优先可移植的 JSON，回退兼容旧的 pickle。

    不向 from_authorized_user_file 传入 scopes，以便 has_scopes() 反映 token 实际授予的权限。
    任何加载失败都视为无凭据（触发重新授权），不会让程序崩溃。
    """
    if os.path.exists(TOKEN_SHEETS_JSON_FILE):
        try:
            return Credentials.from_authorized_user_file(TOKEN_SHEETS_JSON_FILE)
        except Exception as e:
            print(f"⚠ 读取 token_sheets.json 失败，将重新授权: {e}")

    # 兼容旧版 token.pickle（不同 google-auth 版本可能无法反序列化）
    if os.path.exists(TOKEN_PICKLE_FILE):
        try:
            with open(TOKEN_PICKLE_FILE, 'rb') as token:
                return pickle.load(token)
        except Exception as e:
            print(f"⚠ 读取 token.pickle 失败（可能为版本不兼容），将重新授权: {e}")

    return None


def _save_credentials(creds):
    """以可移植的 JSON 格式持久化凭据"""
    with open(TOKEN_SHEETS_JSON_FILE, 'w', encoding='utf-8') as token:
        token.write(creds.to_json())


# 凭据进程内缓存：避免每个表格操作都重复"读 token 文件 + 解析 + 校验"
_cached_creds = None


def authorize_credentials():
    """授权Google API凭据（进程内缓存；过期/失效时自动走原有刷新与重授权流程）"""
    global _cached_creds
    if _cached_creds is not None and _cached_creds.valid:
        return _cached_creds

    setup_google_proxy_env()
    creds = _load_saved_credentials()

    needs_reauth = not creds or not creds.has_scopes(SCOPES_SHEETS)

    if creds and not needs_reauth:
        if creds.expired and creds.refresh_token:
            try:
                refresh_credentials(creds)
                _save_credentials(creds)
            except RefreshError:
                needs_reauth = True
        if not creds.valid:
            needs_reauth = True

    if needs_reauth:
        if creds and not creds.has_scopes(SCOPES_SHEETS):
            print("需要 Gmail 发送权限，正在打开浏览器重新授权...")
        elif creds and not creds.valid:
            print("Token已失效，正在打开浏览器进行重新授权...")
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES_SHEETS)
        creds = flow.run_local_server(port=0)
        _save_credentials(creds)

    _cached_creds = creds
    return creds


def fetch_data(range_name):
    """从Google表格获取数据"""
    creds = authorize_credentials()

    def _get_values():
        service = build_google_service('sheets', 'v4', creds)
        return service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=range_name
        )

    result = execute_with_retry(_get_values)
    return result.get('values', [])


def delete_rows_from_sheet(sheet_id, rows_to_delete):
    """从Google表格中删除行"""
    if not rows_to_delete:
        return
    
    creds = authorize_credentials()
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

    def _delete_rows():
        service = build_google_service('sheets', 'v4', creds)
        return service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body=batch_update_body
        )

    execute_with_retry(_delete_rows)
    print(f"{len(rows_to_delete)} rows deleted.")


def append_data_to_sheet(range_name, data):
    """向Google表格追加数据"""
    creds = authorize_credentials()
    body = {'values': data}

    def _append():
        service = build_google_service('sheets', 'v4', creds)
        return service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body,
            insertDataOption='INSERT_ROWS'
        )

    result = execute_with_retry(_append)
    
    print(f"{result.get('updates').get('updatedRows')} rows appended.")


def batch_update_data_in_sheet(updates):
    """一次 values.batchUpdate 更新多个范围（N 行写回合并为 1 次 API 调用）

    Args:
        updates: [{'range': 'Unfilled!A4', 'values': [[...]]}, ...]
    """
    if not updates:
        return
    creds = authorize_credentials()
    body = {'valueInputOption': 'USER_ENTERED', 'data': updates}

    def _batch_update():
        service = build_google_service('sheets', 'v4', creds)
        return service.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body=body
        )

    result = execute_with_retry(_batch_update)
    print(f"{result.get('totalUpdatedRows', 0)} rows updated ({len(updates)} ranges, 1 API call).")


def update_data_in_sheet(range_name, data):
    """更新Google表格中指定范围的数据"""
    creds = authorize_credentials()
    body = {'values': data}

    def _update():
        service = build_google_service('sheets', 'v4', creds)
        return service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=range_name,
            valueInputOption='USER_ENTERED',
            body=body
        )

    result = execute_with_retry(_update)
    
    print(f"{result.get('updatedRows')} rows updated.")

