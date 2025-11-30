"""
Google Docs API 模块 - 处理Google文档的操作
"""
import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from config import SCOPES_DOCS, DOCUMENT_ID, TOKEN_JSON_FILE, CREDENTIALS_FILE


def build_docs_service():
    """构建Google Docs服务"""
    creds = None
    if os.path.exists(TOKEN_JSON_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_JSON_FILE, SCOPES_DOCS)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES_DOCS)
            creds = flow.run_local_server(port=0)
        
        with open(TOKEN_JSON_FILE, 'w', encoding='utf-8') as token:
            token.write(creds.to_json())
    
    return build('docs', 'v1', credentials=creds)


def retrieve_document_content(service, document_id):
    """获取文档内容"""
    document = service.documents().get(documentId=document_id).execute()
    doc_content = document.get('body').get('content')
    text = ""
    
    for element in doc_content:
        if 'paragraph' in element:
            for elem in element['paragraph']['elements']:
                if 'textRun' in elem and 'content' in elem['textRun']:
                    text += elem['textRun']['content']
    
    return text


def content_exists(service, document_id, content):
    """检查内容是否已存在于文档中"""
    current_content = retrieve_document_content(service, document_id)
    return content in current_content


def append_to_document(service, document_id, text, date_subtitle=""):
    """
    向文档追加内容
    Args:
        service: Google Docs服务
        document_id: 文档ID
        text: 要追加的文本
        date_subtitle: 日期副标题（可选）
    """
    # 获取文档当前内容以确定长度
    document = service.documents().get(documentId=document_id).execute()
    doc_content = document.get('body').get('content')
    end_index = doc_content[-1].get('endIndex', 1) - 1
    
    requests = []
    
    # 如果有日期副标题，先添加
    if date_subtitle:
        formatted_date_subtitle = "\n\n" + date_subtitle + "\n\n"
        requests.extend([
            {
                'insertText': {
                    'location': {'index': end_index},
                    'text': formatted_date_subtitle
                }
            },
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': end_index + 2,
                        'endIndex': end_index + 2 + len(date_subtitle),
                    },
                    'textStyle': {
                        'bold': True,
                        'fontSize': {'magnitude': 16, 'unit': 'PT'}
                    },
                    'fields': 'bold,fontSize'
                }
            },
            {
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': end_index + 2,
                        'endIndex': end_index + 2 + len(date_subtitle),
                    },
                    'paragraphStyle': {'alignment': 'CENTER'},
                    'fields': 'alignment'
                }
            }
        ])
        end_index += len(formatted_date_subtitle)
    
    # 添加主要内容
    requests.append({
        'insertText': {
            'location': {'index': end_index},
            'text': text + "\n\n"
        }
    })
    
    result = service.documents().batchUpdate(
        documentId=document_id, 
        body={'requests': requests}
    ).execute()
    
    return result


def add_wechat_content_to_doc(wechat_template_output, date_subtitle):
    """
    将微信公众号内容添加到Google文档
    Args:
        wechat_template_output: 微信格式的输出
        date_subtitle: 日期副标题
    Returns:
        str: 操作结果消息
    """
    service = build_docs_service()
    
    # 检查日期副标题是否已存在
    date_subtitle_exists = content_exists(service, DOCUMENT_ID, date_subtitle)
    
    # 检查具体内容是否已存在
    job_listing_exists = content_exists(service, DOCUMENT_ID, wechat_template_output)
    
    # 根据存在情况决定是否添加
    if not date_subtitle_exists and not job_listing_exists:
        append_to_document(service, DOCUMENT_ID, wechat_template_output, date_subtitle)
        return "Both date subtitle and job listing added to the document."
    elif date_subtitle_exists and not job_listing_exists:
        append_to_document(service, DOCUMENT_ID, wechat_template_output, "")
        return "Job listing added to the document under an existing date subtitle."
    else:
        return "No new content added; both date subtitle and job listing already exist in the document."

