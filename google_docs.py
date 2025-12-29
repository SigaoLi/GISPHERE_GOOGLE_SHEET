"""
Google Docs API 模块 - 处理Google文档的操作
"""
import os
import re
import pandas as pd
from datetime import datetime
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from config import SCOPES_DOCS, DOCUMENT_ID, TOKEN_JSON_FILE, CREDENTIALS_FILE, BASE_DIR
from data_processor import (
    get_job_category, 
    get_time_category, 
    get_sort_priority,
    parse_deadline_for_sort
)
from utils import get_pinyin_sort_key


def get_openai_key():
    """从openai_key.txt读取OpenAI API密钥"""
    key_file = os.path.join(BASE_DIR, 'openai_key.txt')
    try:
        with open(key_file, 'r', encoding='utf-8') as f:
            key = f.read().strip()
            return key if key else None
    except FileNotFoundError:
        print(f"⚠ 未找到openai_key.txt文件")
        return None
    except Exception as e:
        print(f"⚠ 读取openai_key.txt失败: {e}")
        return None


def call_llm_for_content_organization(existing_content, new_content, date_subtitle):
    """
    使用LLM决定如何组织和插入新内容到现有周期内容中
    
    Args:
        existing_content: 周期中已有的内容（不包含date_subtitle）
        new_content: 要添加的新内容
        date_subtitle: 日期副标题
    
    Returns:
        str: LLM组织后的完整内容（包含date_subtitle）
    """
    try:
        import openai
        from config import OPENAI_BASE_URL
        
        openai_key = get_openai_key()
        if not openai_key:
            print("⚠ 无法获取OpenAI密钥，使用默认规则")
            return None
        
        client = openai.OpenAI(api_key=openai_key, base_url=OPENAI_BASE_URL)
        
        system_prompt = """你是一个专业的文档编辑助手。你的任务是根据现有的文档内容和新的内容，智能地决定如何组织和插入新内容。

规则：
1. 保持文档的层次结构：使用 # 表示一级标题（如"# 博士招生"，"# 博后招聘"，"# 研究助理"，"# 学术会议"等），使用 ## 表示二级标题（如"## 尽快申请："，"## 本月及以后："）
2. 根据职位类型（博士、硕士、博后、研究助理等）和时间紧迫性（尽快申请、本月及以后）进行合理分类
3. 将新内容插入到合适的位置，保持逻辑顺序
4. 如果新内容与现有内容属于同一类别，应该放在一起
5. 输出完整的、格式化的文档内容，包括日期副标题
6. 保持Markdown格式的一致性
7. **重要：联系人信息处理规则**：
   - 如果职位内容中的联系人为空，显示为"联系人：\n- (-)"或类似格式，必须将其修改为"联系方式：\n见详情页"
   - 如果联系人信息不完整或无效，也应替换为"联系方式：\n见详情页"
   - 只有当联系人信息完整有效时，才保留原始格式

输出格式要求：
- 第一行是日期副标题（如：Week: 2025-12-14 to 2025-12-27）
- 然后是一级标题（# 类别名称）
- 然后是二级标题（## 时间类别：）
- 然后是具体职位内容（以 ### 开头）
- 每个职位内容之间用空行分隔"""
        
        user_prompt = f"""现有周期内容（日期：{date_subtitle}）：

{existing_content if existing_content.strip() else "（当前周期还没有内容）"}

新要添加的内容：

{new_content}

请根据上述内容，输出完整的、组织好的文档内容。要求：
1. 包含日期副标题：{date_subtitle}
2. 根据职位类型和时间紧迫性合理分类和组织
3. 将新内容插入到合适的位置
4. 保持Markdown格式
5. 输出完整的文档内容（从日期副标题开始到最后一个职位结束）
6. **重要**：检查所有职位内容中的联系人信息：
   - 如果发现"联系人：\n- (-)"或联系人信息为空/无效的情况，必须将其替换为"联系方式：\n见详情页"
   - 确保所有职位都有有效的联系方式信息

只输出组织后的完整内容，不要包含任何解释或说明。"""
        
        from config import OPENAI_MODEL
        
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.1,
            max_tokens=4000
        )
        
        organized_content = response.choices[0].message.content.strip()
        return organized_content
        
    except ImportError:
        print("⚠ openai库未安装，使用默认规则")
        return None
    except Exception as e:
        print(f"⚠ LLM调用失败: {e}，使用默认规则")
        return None


def build_initial_content_for_new_period(wechat_template_output, date_subtitle, job_row):
    """
    为新周期构建初始内容（使用现有规则）
    
    Args:
        wechat_template_output: 微信格式的输出
        date_subtitle: 日期副标题
        job_row: 职位数据行
    
    Returns:
        str: 格式化的初始内容
    """
    # 获取职位类别
    category = get_job_category(job_row)
    time_category = get_time_category(job_row.get('Deadline', ''))
    
    # 构建内容
    content = f"\n\n{date_subtitle}\n\n"
    content += f"# {category}\n"
    content += f"## {time_category}：\n"
    content += wechat_template_output + "\n"
    
    return content


def build_docs_service():
    """构建Google Docs服务"""
    creds = None
    if os.path.exists(TOKEN_JSON_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_JSON_FILE, SCOPES_DOCS)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                # Token刷新失败（可能被撤销），需要重新授权
                print("Token已失效，正在打开浏览器进行重新授权...")
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES_DOCS)
                creds = flow.run_local_server(port=0)
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


def parse_job_from_text(job_text):
    """
    从文本中解析职位信息
    Args:
        job_text: 职位文本
    Returns:
        dict: 包含职位信息的字典
    """
    job_info = {
        'university': '',
        'country': '',
        'direction': '',
        'job_type': '',
        'deadline': '',
        'deadline_display': '',
        'content': job_text
    }
    
    lines = job_text.strip().split('\n')
    
    for line in lines:
        line = line.strip()
        if line.startswith('### '):
            # 大学名称
            university_full = line.replace('### ', '')
            job_info['university'] = university_full
            # 尝试提取国家（假设格式为"国家+大学"）
            for country in ['美国', '英国', '加拿大', '澳大利亚', '新西兰', '德国', '法国', 
                          '日本', '新加坡', '中国香港', '中国', '瑞士', '荷兰', '瑞典',
                          '丹麦', '挪威', '芬兰', '意大利', '西班牙', '爱尔兰', '比利时',
                          '奥地利', '韩国', '印度', '巴西', '葡萄牙', '波兰']:
                if university_full.startswith(country):
                    job_info['country'] = country
                    break
        elif line.startswith('> '):
            # 方向
            direction_line = line.replace('> ', '')
            if '方向：' in direction_line or '方向:' in direction_line:
                job_info['direction'] = direction_line.replace('方向：', '').replace('方向:', '').strip()
        elif '类型：' in line or '类型:' in line:
            # 职位类型
            job_info['job_type'] = line
        elif '申请截止：' in line or '申请截止:' in line:
            # 截止日期
            deadline_line = line.replace('申请截止：', '').replace('申请截止:', '').strip()
            job_info['deadline_display'] = deadline_line
            job_info['deadline'] = deadline_line
    
    return job_info


def parse_jobs_in_period(doc_content, date_subtitle):
    """
    解析文档中指定周期内的所有职位
    Args:
        doc_content: 文档内容
        date_subtitle: 日期副标题（如 "Week: 2025-11-30 to 2025-12-13"）
    Returns:
        list: 职位信息列表
    """
    jobs = []
    
    # 找到对应周期的内容
    if date_subtitle not in doc_content:
        return jobs
    
    # 分割文档，找到当前周期的部分
    parts = doc_content.split('Week:')
    current_period_content = None
    
    for i, part in enumerate(parts):
        if date_subtitle.replace('Week: ', '') in part:
            # 找到当前周期，提取到下一个 Week: 之前的内容
            if i + 1 < len(parts):
                # 有下一个周期，截取到下一个周期之前
                current_period_content = part.split('Week:')[0]
            else:
                # 这是最后一个周期
                current_period_content = part
            break
    
    if not current_period_content:
        return jobs
    
    # 按 ### 分割职位
    job_texts = current_period_content.split('### ')
    
    for job_text in job_texts[1:]:  # 跳过第一个空元素
        if job_text.strip():
            job_info = parse_job_from_text('### ' + job_text)
            if job_info['university']:
                jobs.append(job_info)
    
    return jobs


def sort_jobs(jobs):
    """
    对职位列表进行排序
    Args:
        jobs: 职位信息列表
    Returns:
        dict: 按类别和时间分组的职位字典
    """
    # 为每个职位添加排序信息
    for job in jobs:
        # 从 job_type 推断类别
        job_type_text = job.get('job_type', '')
        
        if '硕士' in job_type_text:
            job['category'] = '硕士招生'
        elif '博士' in job_type_text:
            job['category'] = '博士招生'
        elif '博后' in job_type_text or '博士后' in job_type_text:
            job['category'] = '博后招聘'
        elif '研究助理' in job_type_text:
            job['category'] = '研究助理招聘'
        elif '暑期学校' in job_type_text:
            job['category'] = '暑期学校'
        elif '学术会议' in job_type_text or '会议' in job_type_text:
            job['category'] = '学术会议'
        elif '研讨会' in job_type_text:
            job['category'] = '研讨会'
        elif '竞赛' in job_type_text:
            job['category'] = '竞赛'
        else:
            job['category'] = '其他'
        
        # 时间类别
        deadline = job.get('deadline', '')
        if deadline in ['Soon', '尽快申请', 'soon', 'SOON']:
            job['time_category'] = '尽快申请'
        else:
            job['time_category'] = '本月及以后'
        
        # 排序键
        job['category_priority'] = get_sort_priority(job['category'])
        job['time_priority'] = 0 if job['time_category'] == '尽快申请' else 1
        job['country_pinyin'] = get_pinyin_sort_key(job.get('country', ''))
        job['direction_pinyin'] = get_pinyin_sort_key(job.get('direction', ''))
        job['deadline_date'] = parse_deadline_for_sort(job.get('deadline', ''))
    
    # 排序
    sorted_jobs = sorted(jobs, key=lambda x: (
        x['category_priority'],      # 一级：职位类型
        x['time_priority'],           # 二级：时间类别
        x['country_pinyin'],          # 三级：国家拼音
        x['direction_pinyin'],        # 四级：方向拼音
        x['deadline_date']            # 五级：截止日期
    ))
    
    # 按类别和时间分组
    grouped = {}
    for job in sorted_jobs:
        category = job['category']
        time_cat = job['time_category']
        
        if category not in grouped:
            grouped[category] = {}
        if time_cat not in grouped[category]:
            grouped[category][time_cat] = []
        
        grouped[category][time_cat].append(job)
    
    return grouped


def build_sorted_content(grouped_jobs, date_subtitle):
    """
    构建排序后的文档内容
    Args:
        grouped_jobs: 分组后的职位字典
        date_subtitle: 日期副标题
    Returns:
        str: 格式化的文档内容
    """
    content = f"\n\n{date_subtitle}\n\n"
    
    # 职位类别顺序
    category_order = ['硕士招生', '博士招生', '博后招聘', '研究助理招聘', 
                     '暑期学校', '学术会议', '研讨会', '竞赛', '其他']
    
    for category in category_order:
        if category not in grouped_jobs:
            continue
        
        # 添加类别标题
        content += f"# {category}\n"
        
        # 时间类别顺序
        time_order = ['尽快申请', '本月及以后']
        
        for time_cat in time_order:
            if time_cat not in grouped_jobs[category]:
                continue
            
            # 添加时间标题
            content += f"## {time_cat}：\n"
            
            # 添加职位
            for job in grouped_jobs[category][time_cat]:
                content += job['content']
                if not job['content'].endswith('\n'):
                    content += '\n'
    
    return content


def find_text_indices_in_document(document, target_text):
    """
    在文档结构中找到目标文本的起始和结束索引
    Args:
        document: Google Docs文档对象
        target_text: 要查找的文本
    Returns:
        tuple: (start_index, end_index) 或 (None, None) 如果未找到
    """
    body_content = document.get('body').get('content')
    accumulated_text = ""
    start_index = None
    end_index = None
    
    for element in body_content:
        if 'paragraph' in element:
            for elem in element['paragraph']['elements']:
                if 'textRun' in elem:
                    text_content = elem['textRun'].get('content', '')
                    current_start = len(accumulated_text)
                    accumulated_text += text_content
                    current_end = len(accumulated_text)
                    
                    # 检查目标文本是否在当前文本运行中
                    if target_text in accumulated_text and start_index is None:
                        # 找到起始位置
                        start_index = accumulated_text.find(target_text)
                    
                    # 如果已经找到起始位置，检查是否找到了完整的目标文本
                    if start_index is not None:
                        if target_text in accumulated_text[start_index:]:
                            end_index = start_index + len(target_text)
                            return start_index, end_index
    
    return start_index, end_index


def find_period_indices(document, doc_content, date_subtitle):
    """
    在文档中找到指定周期的起始和结束索引
    Args:
        document: Google Docs文档对象
        doc_content: 文档的纯文本内容
        date_subtitle: 日期副标题
    Returns:
        tuple: (start_index, end_index) 或 (None, None) 如果未找到
    """
    # 找到当前周期的起始位置（在纯文本中）
    period_start_text = doc_content.find(date_subtitle)
    
    if period_start_text == -1:
        return None, None
    
    # 向前查找，找到当前周期的真正开始
    # 查找前面的 "Week:" 或文档开始
    prev_week_pos = doc_content.rfind('Week:', 0, period_start_text)
    if prev_week_pos != -1:
        # 找到上一个周期的Week标题，当前周期从它之后开始
        # 找到上一个周期结束的位置（下一个换行之后，包括空行）
        search_pos = prev_week_pos
        while search_pos < period_start_text:
            next_newline = doc_content.find('\n', search_pos)
            if next_newline == -1 or next_newline >= period_start_text:
                break
            search_pos = next_newline + 1
        # 当前周期从上一个周期结束后的换行开始
        period_start_text = max(0, search_pos - 2)  # 包括前面的换行
    else:
        # 没有找到上一个周期，当前周期从文档开始或前面的换行开始
        period_start_text = max(0, period_start_text - 2)  # 包括前面的换行
    
    # 找到下一个 "Week:" 的位置，如果没有则到文档末尾
    next_period_start_text = doc_content.find('Week:', period_start_text + len(date_subtitle))
    
    if next_period_start_text == -1:
        # 这是最后一个周期，删除到文档末尾
        period_end_text = len(doc_content)
    else:
        # 删除到下一个周期之前
        period_end_text = next_period_start_text
    
    # 在文档结构中查找对应的索引
    # 通过累积文本内容来匹配位置
    body_content = document.get('body').get('content')
    accumulated_text = ""
    start_index = None
    end_index = None
    
    # 遍历文档结构，累积文本并找到对应的索引位置
    for element in body_content:
        if 'paragraph' in element:
            para = element['paragraph']
            
            for elem in para.get('elements', []):
                if 'textRun' in elem:
                    text_content = elem['textRun'].get('content', '')
                    text_start = elem.get('startIndex', 0)
                    
                    # 累积文本以匹配位置
                    text_before = len(accumulated_text)
                    accumulated_text += text_content
                    text_after = len(accumulated_text)
                    
                    # 检查是否到达了period_start_text
                    if start_index is None and text_after > period_start_text:
                        # 计算在文本运行中的偏移
                        offset = period_start_text - text_before
                        if offset >= 0:
                            start_index = text_start + offset
                    
                    # 检查是否到达了period_end_text
                    if start_index is not None and end_index is None and text_after >= period_end_text:
                        # 计算在文本运行中的偏移
                        offset = period_end_text - text_before
                        if offset >= 0:
                            end_index = text_start + offset
                            break
            
            if end_index is not None:
                break
    
    # 如果无法找到精确索引，使用文档的实际endIndex进行估算
    if start_index is None or end_index is None:
        # 获取文档的实际长度
        body_content = document.get('body').get('content')
        if body_content:
            doc_end_index = body_content[-1].get('endIndex', 1) - 1
        else:
            doc_end_index = 1
        
        # 使用比例估算（更保守的方法）
        if len(doc_content) > 0 and doc_end_index > 1:
            # 使用更保守的估算，避免删除过多内容
            start_ratio = max(0, min(1, period_start_text / len(doc_content)))
            end_ratio = max(0, min(1, period_end_text / len(doc_content)))
            start_index = max(1, int(doc_end_index * start_ratio))
            end_index = min(doc_end_index, int(doc_end_index * end_ratio))
        else:
            return None, None
    
    # 确保索引有效且合理
    if start_index is None or end_index is None:
        return None, None
    
    if start_index >= end_index:
        # 如果索引无效，返回None
        return None, None
    
    # 确保索引在文档范围内
    body_content = document.get('body').get('content')
    if body_content:
        doc_end_index = body_content[-1].get('endIndex', 1) - 1
        start_index = max(1, min(start_index, doc_end_index - 1))
        end_index = max(start_index + 1, min(end_index, doc_end_index))
    
    return start_index, end_index


def replace_period_content(service, document_id, date_subtitle, new_content):
    """
    替换文档中指定周期的内容
    Args:
        service: Google Docs服务
        document_id: 文档ID
        date_subtitle: 日期副标题
        new_content: 新的内容（包含date_subtitle）
    """
    # 获取文档内容
    document = service.documents().get(documentId=document_id).execute()
    doc_content = retrieve_document_content(service, document_id)
    
    # 找到周期的起始和结束位置
    if date_subtitle not in doc_content:
        # 周期不存在，直接追加
        append_to_document(service, document_id, new_content, "")
        return
    
    # 查找周期的索引位置
    start_index, end_index = find_period_indices(document, doc_content, date_subtitle)
    
    if start_index is None or end_index is None:
        # 如果无法找到位置，直接追加（避免错误）
        print("⚠ 无法找到周期位置，使用追加方式")
        append_to_document(service, document_id, new_content, "")
        return
    
    # 构建请求：先删除旧内容，再插入新内容
    # 找到date_subtitle在新内容中的位置（可能前面有换行）
    subtitle_in_new = new_content.find(date_subtitle)
    
    requests = [
        {
            'deleteContentRange': {
                'range': {
                    'startIndex': start_index,
                    'endIndex': end_index
                }
            }
        },
        {
            'insertText': {
                'location': {'index': start_index},
                'text': new_content
            }
        }
    ]
    
    # 在插入内容后，对日期副标题应用格式（居中、加粗、加大字号）
    if subtitle_in_new != -1:
        # 计算date_subtitle在文档中的实际索引位置
        # start_index是插入位置，subtitle_in_new是新内容中date_subtitle的偏移
        subtitle_start_index = start_index + subtitle_in_new
        subtitle_end_index = subtitle_start_index + len(date_subtitle)
        
        # 添加格式设置请求
        requests.extend([
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': subtitle_start_index,
                        'endIndex': subtitle_end_index,
                    },
                    'textStyle': {
                        'bold': True,
                        'fontSize': {'magnitude': 18, 'unit': 'PT'}  # 加大字号到18
                    },
                    'fields': 'bold,fontSize'
                }
            },
            {
                'updateParagraphStyle': {
                    'range': {
                        'startIndex': subtitle_start_index,
                        'endIndex': subtitle_end_index,
                    },
                    'paragraphStyle': {'alignment': 'CENTER'},
                    'fields': 'alignment'
                }
            }
        ])
    
    try:
        service.documents().batchUpdate(
            documentId=document_id,
            body={'requests': requests}
        ).execute()
    except Exception as e:
        # 如果替换失败，回退到追加方式
        print(f"⚠ 替换内容失败，使用追加方式: {e}")
        append_to_document(service, document_id, new_content, "")


def get_period_content_without_subtitle(doc_content, date_subtitle):
    """
    获取周期内容（不包含date_subtitle本身）
    
    Args:
        doc_content: 文档完整内容
        date_subtitle: 日期副标题（如 "Week: 2025-12-14 to 2025-12-27"）
    
    Returns:
        str: 周期内容（不包含date_subtitle）
    """
    if date_subtitle not in doc_content:
        return ""
    
    # 找到date_subtitle的位置
    subtitle_index = doc_content.find(date_subtitle)
    if subtitle_index == -1:
        return ""
    
    # 提取date_subtitle之后的内容
    after_subtitle = doc_content[subtitle_index + len(date_subtitle):]
    
    # 找到下一个Week:的位置（如果有）
    next_week_index = after_subtitle.find('Week:')
    if next_week_index != -1:
        # 有下一个周期，只提取到下一个周期之前
        after_subtitle = after_subtitle[:next_week_index]
    
    # 移除开头的换行和空白
    after_subtitle = after_subtitle.lstrip('\n\r \t')
    
    return after_subtitle


def add_wechat_content_to_doc_sorted(wechat_template_output, date_subtitle, job_row):
    """
    将微信公众号内容添加到Google文档（使用LLM智能组织）
    
    规则：
    - 如果是新周期（还没有内容），则根据现有规则先写入一条
    - 如果已有内容，则将现有内容和新内容一起给LLM，让LLM决定插入位置和标题层级
    
    Args:
        wechat_template_output: 微信格式的输出
        date_subtitle: 日期副标题
        job_row: 职位数据行（pandas Series）
    Returns:
        str: 操作结果消息
    """
    service = build_docs_service()
    doc_content = retrieve_document_content(service, DOCUMENT_ID)
    
    # 检查职位是否已存在
    university_line = None
    for line in wechat_template_output.split('\n'):
        if line.startswith('### '):
            university_line = line
            break
    
    if university_line and university_line in doc_content:
        return "Job listing already exists in the document."
    
    # 检查周期是否存在
    period_exists = date_subtitle in doc_content
    
    if not period_exists:
        # 新周期，使用现有规则先写入一条
        print("📝 新周期，使用现有规则写入初始内容...")
        initial_content = build_initial_content_for_new_period(
            wechat_template_output, 
            date_subtitle, 
            job_row
        )
        append_to_document(service, DOCUMENT_ID, initial_content.strip() + "\n", "")
        return "New period created with initial content using default rules."
    else:
        # 已有内容，使用LLM决定如何组织
        print("🤖 周期已有内容，使用LLM智能组织...")
        
        # 获取周期中现有的内容（不包含date_subtitle）
        existing_content = get_period_content_without_subtitle(doc_content, date_subtitle)
        
        # 调用LLM组织内容
        organized_content = call_llm_for_content_organization(
            existing_content,
            wechat_template_output,
            date_subtitle
        )
        
        if organized_content:
            # 使用LLM组织的内容替换
            replace_period_content(service, DOCUMENT_ID, date_subtitle, organized_content.strip() + "\n")
            return "Content organized and updated using LLM."
        else:
            # LLM调用失败，回退到原有规则
            print("⚠ LLM调用失败，回退到原有排序规则...")
            
            # 解析当前周期内的所有职位
            existing_jobs = parse_jobs_in_period(doc_content, date_subtitle)
            
            # 创建新职位信息
            deadline_value = job_row.get('Deadline', '')
            if pd.isna(deadline_value):
                deadline_str = ''
            elif hasattr(deadline_value, 'strftime'):
                deadline_str = deadline_value.strftime('%Y-%m-%d')
            else:
                deadline_str = str(deadline_value).strip()
            
            new_job = {
                'university': job_row.get('University_CN', ''),
                'country': job_row.get('Country_CN', ''),
                'direction': job_row.get('Direction', ''),
                'deadline': deadline_str,
                'content': wechat_template_output + '\n'
            }
            
            new_job['category'] = get_job_category(job_row)
            new_job['time_category'] = get_time_category(new_job['deadline'])
            
            existing_jobs.append(new_job)
            
            # 排序所有职位
            grouped_jobs = sort_jobs(existing_jobs)
            
            # 构建完整的排序内容
            sorted_content = build_sorted_content(grouped_jobs, date_subtitle)
            
            # 替换旧内容
            replace_period_content(service, DOCUMENT_ID, date_subtitle, sorted_content.strip() + "\n")
            return "Content updated using fallback sorting rules."


def add_wechat_content_to_doc(wechat_template_output, date_subtitle):
    """
    将微信公众号内容添加到Google文档（简化版，保持向后兼容）
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

