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

# 周期标题的正则表达式模式
# 匹配新格式: "海外资讯 136 | 2026.02.08 - 2026.02.21"
# 兼容旧格式: "第XXX期 - Week: YYYY-MM-DD to YYYY-MM-DD" 或 "Week: YYYY-MM-DD to YYYY-MM-DD"
PERIOD_TITLE_PATTERN = re.compile(
    r'海外资讯\s+\d+\s*\|\s*\d{4}\.\d{2}\.\d{2}\s*-\s*\d{4}\.\d{2}\.\d{2}'
    r'|(第\d+期\s*-\s*)?Week:\s*\d{4}-\d{2}-\d{2}\s+to\s+\d{4}-\d{2}-\d{2}'
)


def find_next_period_start(doc_content, start_pos):
    """
    在文档内容中从指定位置开始查找下一个周期标题的起始位置
    
    Args:
        doc_content: 文档内容字符串
        start_pos: 开始搜索的位置
    
    Returns:
        int: 下一个周期标题的起始位置，如果没找到返回 -1
    """
    # 在 start_pos 之后的内容中搜索
    search_content = doc_content[start_pos:]
    match = PERIOD_TITLE_PATTERN.search(search_content)
    
    if match:
        return start_pos + match.start()
    return -1


def split_doc_by_periods(doc_content):
    """
    将文档内容按周期标题分割
    
    Args:
        doc_content: 文档内容字符串
    
    Returns:
        list: 每个周期的内容列表，每个元素是 (title, content) 元组
    """
    periods = []
    
    # 找到所有周期标题
    matches = list(PERIOD_TITLE_PATTERN.finditer(doc_content))
    
    for i, match in enumerate(matches):
        title = match.group()
        title_start = match.start()
        title_end = match.end()
        
        # 确定内容结束位置
        if i + 1 < len(matches):
            content_end = matches[i + 1].start()
        else:
            content_end = len(doc_content)
        
        content = doc_content[title_end:content_end]
        periods.append((title, content))
    
    return periods


def clean_trailing_spaces(text):
    """
    清理文本中每行末尾的多余空格（Markdown软换行符）
    
    Args:
        text: 要清理的文本
    
    Returns:
        str: 清理后的文本
    """
    if not text:
        return text
    
    # 将每行末尾的空格去除（保留换行符）
    lines = text.split('\n')
    cleaned_lines = [line.rstrip() for line in lines]
    return '\n'.join(cleaned_lines)


def get_openai_key():
    """从openai_key.txt读取OpenAI API密钥"""
    key_file = os.path.join(BASE_DIR, 'keys', 'openai_key.txt')
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
        from config import OPENAI_BASE_URL, OPENAI_MODEL
        from logger import log_llm_conversation
        
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
- 第一行是日期副标题（如：海外资讯 134 | 2025.12.14 - 2025.12.27）
- 在每个一级标题（# 类别名称）前加上 ---
- 如果一级标题后紧跟二级标题（## 时间类别：），则不需要在它们之间加 ---
- 如果同一个一级标题下有多个二级标题，从第二个二级标题开始，在其前面加上 ---
- 在所有内容的最末尾加上 ---
- 然后是具体职位内容（以 ### 开头）
- 每个职位内容之间用空行分隔

示例结构：
---
# 博士招生
## 尽快申请：
### 大学A
（职位内容）
---
## 本月及以后：
### 大学B
（职位内容）
---
# 博后招聘
## 尽快申请：
### 大学C
（职位内容）
---"""
        
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
        
        # 清理行末多余空格（Markdown软换行符）
        organized_content = clean_trailing_spaces(organized_content)
        
        # 记录LLM对话
        log_llm_conversation(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            response=organized_content,
            model=OPENAI_MODEL,
            metadata={
                'date_subtitle': date_subtitle,
                'existing_content_length': len(existing_content),
                'new_content_length': len(new_content),
                'response_length': len(organized_content),
                'temperature': 0.1,
                'max_tokens': 4000
            }
        )
        
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
    content += "---\n"
    content += f"# {category}\n"
    content += f"## {time_category}：\n"
    content += wechat_template_output + "\n"
    content += "---\n"
    
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
    向文档末尾追加内容
    
    注意：此函数确保内容被插入到文档的最末尾位置，不会影响之前的内容。
    
    Args:
        service: Google Docs服务
        document_id: 文档ID
        text: 要追加的文本
        date_subtitle: 日期副标题（可选）
    """
    # 获取文档当前内容以确定插入位置（文档末尾）
    document = service.documents().get(documentId=document_id).execute()
    doc_content = document.get('body').get('content')
    
    # Google Docs API 的 endIndex 是排他的，减1得到最后一个有效插入位置
    # 这确保新内容被追加到文档的最末尾
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
            university_full = line.replace('### ', '')
            job_info['university'] = university_full
            for country in ['美国', '英国', '加拿大', '澳大利亚', '新西兰', '德国', '法国', 
                          '日本', '新加坡', '中国香港', '中国', '瑞士', '荷兰', '瑞典',
                          '丹麦', '挪威', '芬兰', '意大利', '西班牙', '爱尔兰', '比利时',
                          '奥地利', '韩国', '印度', '巴西', '葡萄牙', '波兰']:
                if university_full.startswith(country):
                    job_info['country'] = country
                    break
        elif line.startswith('> '):
            direction_line = line.replace('> ', '')
            if '方向：' in direction_line or '方向:' in direction_line:
                job_info['direction'] = direction_line.replace('方向：', '').replace('方向:', '').strip()
        elif '类型：' in line or '类型:' in line:
            job_info['job_type'] = line
        elif '申请截止：' in line or '申请截止:' in line:
            deadline_line = line.replace('申请截止：', '').replace('申请截止:', '').strip()
            job_info['deadline_display'] = deadline_line
            job_info['deadline'] = deadline_line
    
    return job_info


def parse_jobs_in_period(doc_content, date_subtitle):
    """
    解析文档中指定周期内的所有职位
    Args:
        doc_content: 文档内容
        date_subtitle: 日期副标题（如 "海外资讯 134 | 2025.11.30 - 2025.12.13"）
    Returns:
        list: 职位信息列表
    """
    jobs = []
    
    # 找到对应周期的内容
    if date_subtitle not in doc_content:
        return jobs
    
    # 找到当前周期标题的位置
    subtitle_pos = doc_content.find(date_subtitle)
    if subtitle_pos == -1:
        return jobs
    
    # 从标题结束后开始提取内容
    content_start = subtitle_pos + len(date_subtitle)
    
    # 找到下一个周期的位置（使用新的辅助函数）
    next_period_pos = find_next_period_start(doc_content, content_start)
    
    if next_period_pos == -1:
        # 这是最后一个周期
        current_period_content = doc_content[content_start:]
    else:
        # 截取到下一个周期之前
        current_period_content = doc_content[content_start:next_period_pos]
    
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
        
        content += "---\n"
        content += f"# {category}\n"
        
        time_order = ['尽快申请', '本月及以后']
        is_first_time_cat = True
        
        for time_cat in time_order:
            if time_cat not in grouped_jobs[category]:
                continue
            
            if not is_first_time_cat:
                content += "---\n"
            is_first_time_cat = False
            
            content += f"## {time_cat}：\n"
            
            for job in grouped_jobs[category][time_cat]:
                content += job['content']
                if not job['content'].endswith('\n'):
                    content += '\n'
    
    content += "---\n"
    
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


def find_period_content_indices(document, doc_content, date_subtitle):
    """
    在文档中找到指定周期的内容起始和结束索引（不包括标题本身）
    
    Args:
        document: Google Docs文档对象
        doc_content: 文档的纯文本内容
        date_subtitle: 日期副标题
    Returns:
        tuple: (start_index, end_index) 或 (None, None) 如果未找到
        
    注意：返回的范围是从标题后的第一个换行符之后开始，到下一个周期之前结束。
    这样可以保留标题不被删除。
    """
    # 找到当前周期标题的位置（在纯文本中）
    subtitle_pos = doc_content.find(date_subtitle)
    
    if subtitle_pos == -1:
        return None, None
    
    # 从标题结束后的第一个换行符之后开始（保留标题和标题后的换行）
    # 找到标题结束位置
    subtitle_end = subtitle_pos + len(date_subtitle)
    
    # 跳过标题后面的换行符（通常是 \n\n）
    content_start = subtitle_end
    while content_start < len(doc_content) and doc_content[content_start] in '\n\r':
        content_start += 1
    
    # 如果没有找到非换行字符，说明标题后没有内容
    if content_start >= len(doc_content):
        return None, None
    
    # 找到下一个周期标题的位置，如果没有则到文档末尾
    next_period_start = find_next_period_start(doc_content, content_start)
    
    if next_period_start == -1:
        # 这是最后一个周期，到文档末尾
        content_end = len(doc_content)
    else:
        # 到下一个周期之前（保留下一个周期前的空行）
        content_end = next_period_start
        # 向前回退，去掉末尾的空行
        while content_end > content_start and doc_content[content_end - 1] in '\n\r':
            content_end -= 1
        # 但保留一个换行符
        if content_end < next_period_start:
            content_end += 1
    
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
                    
                    # 检查是否到达了content_start
                    if start_index is None and text_after > content_start:
                        # 计算在文本运行中的偏移
                        offset = content_start - text_before
                        if offset >= 0:
                            start_index = text_start + offset
                    
                    # 检查是否到达了content_end
                    if start_index is not None and end_index is None and text_after >= content_end:
                        # 计算在文本运行中的偏移
                        offset = content_end - text_before
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
            start_ratio = max(0, min(1, content_start / len(doc_content)))
            end_ratio = max(0, min(1, content_end / len(doc_content)))
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
    替换文档中指定周期的内容（只替换标题后的内容，保留标题本身）
    
    Args:
        service: Google Docs服务
        document_id: 文档ID
        date_subtitle: 日期副标题
        new_content: 新的内容（不包含date_subtitle，只包含职位内容）
    
    注意：
    - 此函数只删除标题后的内容，保留标题及其格式
    - 新内容会在插入前自动添加前导换行符确保格式正确
    - 删除和插入操作分开执行，中间有3秒延时和验证逻辑
    """
    import time
    
    MAX_DELETE_RETRIES = 3
    DELETE_VERIFY_DELAY = 3  # 删除后等待3秒再验证
    
    # 获取文档内容
    document = service.documents().get(documentId=document_id).execute()
    doc_content = retrieve_document_content(service, document_id)
    
    # 保存原始内容用于验证删除
    original_content = doc_content
    
    # 找到周期的标题
    if date_subtitle not in doc_content:
        # 周期不存在，直接追加
        append_to_document(service, document_id, new_content, date_subtitle)
        return
    
    # 查找周期内容的索引位置（标题后的内容）
    start_index, end_index = find_period_content_indices(document, doc_content, date_subtitle)
    
    if start_index is None or end_index is None:
        # 如果无法找到位置，直接追加到周期末尾
        print("⚠ 无法找到周期内容位置，使用追加方式")
        append_content_to_period_end(service, document_id, doc_content, date_subtitle, new_content)
        return
    
    # 获取将要删除的内容（用于验证）
    content_to_delete = get_period_content_without_subtitle(doc_content, date_subtitle)
    
    # 步骤1: 删除旧内容（带重试机制）
    delete_success = False
    for attempt in range(1, MAX_DELETE_RETRIES + 1):
        print(f"   🗑️ 尝试删除旧内容 (第 {attempt}/{MAX_DELETE_RETRIES} 次)...")
        
        try:
            # 重新获取文档和索引（因为可能已经有变化）
            if attempt > 1:
                document = service.documents().get(documentId=document_id).execute()
                doc_content = retrieve_document_content(service, document_id)
                start_index, end_index = find_period_content_indices(document, doc_content, date_subtitle)
                
                if start_index is None or end_index is None:
                    print("   ⚠ 无法找到内容索引，跳过删除")
                    break
            
            # 执行删除操作
            delete_request = {
                'deleteContentRange': {
                    'range': {
                        'startIndex': start_index,
                        'endIndex': end_index
                    }
                }
            }
            
            service.documents().batchUpdate(
                documentId=document_id,
                body={'requests': [delete_request]}
            ).execute()
            
            # 等待3秒
            print(f"   ⏳ 等待 {DELETE_VERIFY_DELAY} 秒后验证删除结果...")
            time.sleep(DELETE_VERIFY_DELAY)
            
            # 验证删除是否成功
            current_content = retrieve_document_content(service, document_id)
            current_period_content = get_period_content_without_subtitle(current_content, date_subtitle)
            
            # 检查内容是否已被删除（当前周期内容应该比原来短或为空）
            if len(current_period_content.strip()) < len(content_to_delete.strip()) * 0.5:
                # 删除成功（内容减少了50%以上）
                print("   ✅ 删除验证通过，旧内容已成功删除")
                delete_success = True
                break
            else:
                print(f"   ⚠ 删除验证失败，内容似乎未被删除，将重试...")
                
        except Exception as e:
            print(f"   ⚠ 删除操作出错: {e}")
            if attempt < MAX_DELETE_RETRIES:
                print(f"   ⏳ 等待 {DELETE_VERIFY_DELAY} 秒后重试...")
                time.sleep(DELETE_VERIFY_DELAY)
    
    if not delete_success:
        print("   ❌ 删除操作失败，为避免重复内容，改用追加方式")
        # 不执行插入，直接追加新内容到末尾
        append_content_to_period_end(service, document_id, 
                                     retrieve_document_content(service, document_id), 
                                     date_subtitle, new_content)
        return
    
    # 步骤2: 插入新内容
    print("   📝 插入新内容...")
    
    # 重新获取文档以获取正确的插入位置
    document = service.documents().get(documentId=document_id).execute()
    doc_content = retrieve_document_content(service, document_id)
    
    # 找到标题后的插入位置
    subtitle_pos = doc_content.find(date_subtitle)
    if subtitle_pos == -1:
        print("   ⚠ 无法找到标题位置，使用追加方式")
        append_to_document(service, document_id, new_content, "")
        return
    
    # 计算插入索引
    subtitle_end_text = subtitle_pos + len(date_subtitle)
    
    # 跳过标题后的换行符
    insert_text_pos = subtitle_end_text
    while insert_text_pos < len(doc_content) and doc_content[insert_text_pos] in '\n\r':
        insert_text_pos += 1
    
    # 将纯文本位置转换为API索引
    body_content = document.get('body').get('content')
    accumulated_text = ""
    insert_index = None
    
    for element in body_content:
        if 'paragraph' in element:
            for elem in element['paragraph'].get('elements', []):
                if 'textRun' in elem:
                    text_content = elem['textRun'].get('content', '')
                    text_start = elem.get('startIndex', 0)
                    text_before = len(accumulated_text)
                    accumulated_text += text_content
                    text_after = len(accumulated_text)
                    
                    if insert_index is None and text_after >= insert_text_pos:
                        offset = insert_text_pos - text_before
                        if offset >= 0:
                            insert_index = text_start + offset
                            break
            if insert_index is not None:
                break
    
    if insert_index is None:
        # 回退：使用标题结束位置
        insert_index = subtitle_end_text + 1
        
    # 确保插入位置不超过文档范围
    if body_content:
        doc_end_index = body_content[-1].get('endIndex', 1)
        # Google Docs API 要求 index < endIndex
        # 如果 insert_index 大于等于 doc_end_index，说明超出了文档范围
        if insert_index >= doc_end_index:
            # 调整为文档最后一个字符的位置（通常是末尾的换行符前）
            insert_index = max(1, doc_end_index - 1)
            print(f"   ⚠️ 调整插入位置到文档末尾: {insert_index}")
    
    # 计算已有的换行符数量
    existing_newlines_count = insert_text_pos - subtitle_end_text
    # 确保标题和内容之间刚好有一个空行（两个换行符）
    prefix_newlines_count = max(0, 2 - existing_newlines_count)
    
    # 确保新内容前面有换行符
    formatted_content = ("\n" * prefix_newlines_count) + new_content.strip() + "\n"
    
    try:
        insert_request = {
            'insertText': {
                'location': {'index': insert_index},
                'text': formatted_content
            }
        }
        
        service.documents().batchUpdate(
            documentId=document_id,
            body={'requests': [insert_request]}
        ).execute()
        
        print("   ✅ 新内容插入成功")
        
    except Exception as e:
        print(f"   ⚠ 插入内容失败: {e}，使用追加方式")
        append_content_to_period_end(service, document_id, 
                                     retrieve_document_content(service, document_id),
                                     date_subtitle, new_content)


def append_content_to_period_end(service, document_id, doc_content, date_subtitle, new_content):
    """
    将新内容追加到指定周期的末尾
    
    Args:
        service: Google Docs服务
        document_id: 文档ID
        doc_content: 文档内容
        date_subtitle: 日期副标题
        new_content: 要追加的新内容
    """
    # 找到周期标题的位置
    subtitle_pos = doc_content.find(date_subtitle)
    if subtitle_pos == -1:
        # 周期不存在，追加到文档末尾
        append_to_document(service, document_id, new_content, "")
        return
    
    # 找到周期内容的结束位置（下一个周期标题之前或文档末尾）
    subtitle_end = subtitle_pos + len(date_subtitle)
    next_period_start = find_next_period_start(doc_content, subtitle_end)
    
    if next_period_start == -1:
        # 这是最后一个周期，追加到文档末尾
        document = service.documents().get(documentId=document_id).execute()
        doc_body_content = document.get('body').get('content')
        end_index = doc_body_content[-1].get('endIndex', 1) - 1
    else:
        # 找到下一个周期前的位置
        # 需要将纯文本位置转换为API索引
        document = service.documents().get(documentId=document_id).execute()
        body_content = document.get('body').get('content')
        accumulated_text = ""
        end_index = None
        
        for element in body_content:
            if 'paragraph' in element:
                for elem in element['paragraph'].get('elements', []):
                    if 'textRun' in elem:
                        text_content = elem['textRun'].get('content', '')
                        text_start = elem.get('startIndex', 0)
                        text_before = len(accumulated_text)
                        accumulated_text += text_content
                        text_after = len(accumulated_text)
                        
                        if end_index is None and text_after >= next_period_start:
                            offset = next_period_start - text_before
                            if offset >= 0:
                                end_index = text_start + offset
                                break
                if end_index is not None:
                    break
        
        if end_index is None:
            # 回退到文档末尾
            end_index = body_content[-1].get('endIndex', 1) - 1
    
    # 确保新内容前面有换行符
    formatted_content = "\n\n" + new_content.strip() + "\n"
    
    requests = [
        {
            'insertText': {
                'location': {'index': end_index},
                'text': formatted_content
            }
        }
    ]
    
    try:
        service.documents().batchUpdate(
            documentId=document_id,
            body={'requests': requests}
        ).execute()
    except Exception as e:
        print(f"⚠ 追加内容失败: {e}")


def get_period_content_without_subtitle(doc_content, date_subtitle):
    """
    获取周期内容（不包含date_subtitle本身）
    
    Args:
        doc_content: 文档完整内容
        date_subtitle: 日期副标题（如 "海外资讯 134 | 2025.12.14 - 2025.12.27"）
    
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
    
    # 找到下一个周期标题的位置（如果有）
    next_period_index = find_next_period_start(after_subtitle, 0)
    if next_period_index != -1:
        # 有下一个周期，只提取到下一个周期之前
        after_subtitle = after_subtitle[:next_period_index]
    
    # 移除开头的换行和空白
    after_subtitle = after_subtitle.lstrip('\n\r \t')
    
    return after_subtitle


def ensure_current_period_exists(date_subtitle):
    """
    确保当前周期标题存在于文档中
    
    如果文档中不存在当前周期的日期标题，则创建一个带格式的周期标题。
    
    Args:
        date_subtitle: 日期副标题（如 "海外资讯 134 | 2026.01.11 - 2026.01.24"）
    
    Returns:
        tuple: (period_created, message) - period_created 为 True 表示创建了新周期
    """
    service = build_docs_service()
    doc_content = retrieve_document_content(service, DOCUMENT_ID)
    
    # 检查周期是否已存在
    if date_subtitle in doc_content:
        return False, f"周期 '{date_subtitle}' 已存在于文档中"
    
    # 周期不存在，创建新周期标题（带格式）
    print(f"📅 创建新周期标题: {date_subtitle}")
    
    # 获取文档当前内容以确定插入位置（确保插入到文档最末尾）
    document = service.documents().get(documentId=DOCUMENT_ID).execute()
    doc_body_content = document.get('body').get('content')
    
    # Google Docs API 的 endIndex 是排他的（exclusive），所以减1得到最后一个有效位置
    # 这确保新内容被插入到文档的最末尾
    end_index = doc_body_content[-1].get('endIndex', 1) - 1
    
    # 打印调试信息，确认插入位置
    print(f"   📍 插入位置: 文档末尾 (index: {end_index})")
    
    # 构建格式化的周期标题
    formatted_date_subtitle = "\n\n" + date_subtitle + "\n\n"
    
    requests = [
        {
            'insertText': {
                'location': {'index': end_index},
                'text': formatted_date_subtitle
            }
        },
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': end_index + 2,  # 跳过前面的换行
                    'endIndex': end_index + 2 + len(date_subtitle),
                },
                'textStyle': {
                    'bold': True,
                    'fontSize': {'magnitude': 18, 'unit': 'PT'}
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
    ]
    
    try:
        service.documents().batchUpdate(
            documentId=DOCUMENT_ID,
            body={'requests': requests}
        ).execute()
        return True, f"已创建新周期标题: {date_subtitle}"
    except Exception as e:
        print(f"⚠ 创建周期标题失败: {e}")
        return False, f"创建周期标题失败: {e}"


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
    
    # 检查职位是否已存在（仅在当前周期内检查，同时检查大学名称和研究方向）
    university_line = None
    direction_line = None
    for line in wechat_template_output.split('\n'):
        if line.startswith('### '):
            university_line = line
        elif line.startswith('> ') and '方向' in line:
            direction_line = line
        # 找到两者后就停止
        if university_line and direction_line:
            break
    
    # 获取当前周期的内容（仅在当前周期内检查重复，不影响其他周期）
    current_period_content = get_period_content_without_subtitle(doc_content, date_subtitle)
    
    # 检查重复逻辑：大学名称 + 研究方向都匹配才认为是重复（仅限当前周期）
    if university_line and university_line in current_period_content:
        if direction_line:
            # 如果有研究方向，需要同时匹配大学和方向
            if direction_line in current_period_content:
                return "Job listing already exists in the document (university + direction matched in current period)."
            # 大学存在但方向不同，不是重复
        else:
            # 没有研究方向信息，只能依靠大学名称判断
            return "Job listing already exists in the document (university matched in current period)."
    
    # 检查周期是否存在
    period_exists = date_subtitle in doc_content
    
    if not period_exists:
        # 新周期，先创建格式化的周期标题
        print("📝 新周期，创建格式化的周期标题...")
        
        # 先创建格式化的周期标题
        period_created, period_message = ensure_current_period_exists(date_subtitle)
        print(f"   {period_message}")
        
        # 然后添加职位内容（不包含 date_subtitle，因为已经创建了）
        category = get_job_category(job_row)
        time_category = get_time_category(job_row.get('Deadline', ''))
        
        content_without_subtitle = "---\n"
        content_without_subtitle += f"# {category}\n"
        content_without_subtitle += f"## {time_category}：\n"
        content_without_subtitle += wechat_template_output + "\n"
        content_without_subtitle += "---\n"
        
        append_to_document(service, DOCUMENT_ID, content_without_subtitle.strip() + "\n", "")
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
            # LLM返回的内容可能包含date_subtitle，需要移除它
            content_to_replace = organized_content.strip()
            if content_to_replace.startswith(date_subtitle):
                content_to_replace = content_to_replace[len(date_subtitle):].lstrip('\n\r')
            elif date_subtitle in content_to_replace:
                # 如果date_subtitle在中间，提取之后的内容
                idx = content_to_replace.find(date_subtitle)
                content_to_replace = content_to_replace[idx + len(date_subtitle):].lstrip('\n\r')
            
            replace_period_content(service, DOCUMENT_ID, date_subtitle, content_to_replace + "\n")
            return "Content organized and updated using LLM."
        else:
            # LLM调用失败，简单追加新内容到周期末尾
            print("⚠ LLM调用失败，将新内容追加到周期末尾...")
            
            # 构建新职位内容（包含类别和时间标题）
            category = get_job_category(job_row)
            time_category = get_time_category(job_row.get('Deadline', ''))
            
            new_content = "---\n"
            new_content += f"# {category}\n"
            new_content += f"## {time_category}：\n"
            new_content += wechat_template_output + "\n"
            new_content += "---\n"
            
            # 追加到周期末尾
            append_content_to_period_end(service, DOCUMENT_ID, doc_content, date_subtitle, new_content.strip())
            return "Content appended to period end (LLM unavailable)."


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

