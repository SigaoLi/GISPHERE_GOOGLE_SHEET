"""
数据处理模块 - 处理数据转换和格式化逻辑
"""
import pandas as pd
from datetime import date
from config import (
    COUNTRY_DICTIONARY, 
    JOB_DICTIONARY, 
    SUBJECT_DICTIONARY,
    LABEL_COLUMNS,
    REQUIRED_COLUMNS
)
from utils import (
    number_to_english_words,
    number_to_chinese_words,
    safe_convert_to_int,
    convert_date_to_chinese,
    get_pinyin_sort_key
)


def create_job_title(row):
    """创建英文职位标题"""
    job_titles = []
    non_student_titles = {'PostDoc', 'Research Assistant', 'Competition', 'Summer School', 'Conference', 'Workshop'}
    
    for column in ['Master Student', 'Doctoral Student', 'PostDoc', 'Research Assistant', 
                   'Competition', 'Summer School', 'Conference', 'Workshop']:
        if row[column] in [1, '1', 1.0]:
            if column not in non_student_titles:
                job_titles.append(column.replace(" Student", " Student"))
            else:
                job_titles.append(column)
    
    return ' or '.join(job_titles)


def map_job_titles(job_en):
    """将英文职位映射为中文"""
    job_en_list = [job.strip() for job in job_en.split(' or ')]
    job_cn_list = []
    
    for job in job_en_list:
        if job in JOB_DICTIONARY:
            job_cn_list.append(JOB_DICTIONARY[job])
        else:
            job_with_student = f"{job} Student"
            job_cn_list.append(JOB_DICTIONARY.get(job_with_student, ""))
    
    if len(job_cn_list) > 1:
        job_cn_list[0] = job_cn_list[0].replace("研究生", "", 1)
    
    return '或'.join(job_cn_list)


def create_english_title(university_en, country_en, number_places_en, job_en):
    """创建英文标题"""
    hosting_jobs = {'Competition', 'Summer School', 'Conference', 'Workshop'}
    verb = " is hosting a " if job_en in hosting_jobs else " is recruiting "
    title = university_en + " in " + country_en + verb
    
    if number_places_en is not None and number_places_en != "":
        title += "for " + number_places_en + " " + job_en
        number = safe_convert_to_int(number_places_en.split()[0])
        if number and number != 1:
            title += "s"
    else:
        title += job_en
    
    return title


def create_chinese_title(country_cn, university_cn, number_places_cn, job_cn):
    """创建中文标题"""
    hosting_jobs = {'竞赛', '暑期学校', '学术会议', '研讨会'}
    recruiting_jobs = {'博士后', '研究助理'}
    
    if job_cn in hosting_jobs:
        verb = "举办"
    elif job_cn in recruiting_jobs:
        verb = "招聘"
    else:
        verb = "招生"
    
    title = ""
    if university_cn.startswith(country_cn):
        title += university_cn
    else:
        title += country_cn + university_cn
    
    if number_places_cn is not None and number_places_cn != "":
        title += verb + number_places_cn + "名" + job_cn
    else:
        title += verb + job_cn
    
    return title


def set_label_columns(selected_row, sql_table, label_columns):
    """设置标签列"""
    for column in label_columns:
        if pd.isna(selected_row.iloc[0][column]) or selected_row.iloc[0][column] in [None, '']:
            sql_table[f"Label_{column}"] = [0]
        else:
            sql_table[f"Label_{column}"] = [selected_row.iloc[0][column]]
    return sql_table


def check_required_fields(selected_row):
    """检查必填字段是否完整"""
    for col in REQUIRED_COLUMNS:
        if pd.isna(selected_row[col]).iloc[0] or selected_row[col].iloc[0] == '':
            return '1'
    return ''


def create_sql_table(selected_row, new_event_id):
    """
    创建SQL表格数据
    Args:
        selected_row: 选中的数据行
        new_event_id: 新的事件ID
    Returns:
        DataFrame: 格式化的SQL表格数据
    """
    # 格式化Deadline，只保留日期部分（YYYY-MM-DD）
    deadline_value = selected_row["Deadline"].iloc[0]
    if pd.notna(deadline_value) and deadline_value != 'Soon':
        try:
            # 如果是datetime对象，转换为日期字符串
            if hasattr(deadline_value, 'strftime'):
                formatted_deadline = deadline_value.strftime('%Y-%m-%d')
            else:
                # 如果是字符串，尝试解析后格式化
                deadline_dt = pd.to_datetime(deadline_value, errors='coerce')
                if pd.notna(deadline_dt):
                    formatted_deadline = deadline_dt.strftime('%Y-%m-%d')
                else:
                    formatted_deadline = str(deadline_value)
        except:
            formatted_deadline = str(deadline_value)
    else:
        formatted_deadline = str(deadline_value)
    
    # 创建基础表格
    sql_table = pd.DataFrame({
        'Event_ID': [new_event_id],
        'University_CN': selected_row["University_CN"],
        'University_EN': selected_row["University_EN"],
        'Country_CN': selected_row["Country_CN"],
        'Country_EN': [None],
        'Job_CN': [None],
        'Job_EN': [None],
        'Description': (
            "<p>" + selected_row["Direction"].astype(str) + 
            "; <br>Deadline: " + formatted_deadline + 
            "; <br>Contact: " + selected_row["Contact_Name"].astype(str) + 
            " (" + selected_row["Contact_Email"].astype(str) + 
            "); <br>URL: " + selected_row["Source"].astype(str) + "</p>"
        ),
        'Title_CN': [None],
        'Title_EN': [None],
        'Label_Physical_Geo': [0],
        'Label_Human_Geo': [0],
        'Label_Urban': [0],
        'Label_GIS': [0],
        'Label_RS': [0],
        'Label_GNSS': [0],
        'Date': [date.today().strftime('%Y-%m-%d')],
        'University_ID': [None],
        'IS_Public': [1],
        'IS_Deleted': [0],
        'Event_CN': [None],
        'EVENT_EN': [None]
    })
    
    # 映射英文国家
    sql_table["Country_EN"] = selected_row["Country_CN"].map(COUNTRY_DICTIONARY)
    
    # 创建英文职位
    sql_table['Job_EN'] = create_job_title(selected_row.iloc[0])
    
    # 映射中文职位
    sql_table["Job_CN"] = sql_table["Job_EN"].apply(map_job_titles)
    
    # 处理职位数量
    number_places = selected_row["Number_Places"].item()
    number_places_int = safe_convert_to_int(number_places)
    number_places_en = None if number_places_int is None else number_to_english_words(number_places_int)
    number_places_cn = None if number_places_int is None else number_to_chinese_words(number_places_int)
    
    # 创建标题
    sql_table["Title_EN"] = sql_table.apply(
        lambda x: create_english_title(x['University_EN'], x['Country_EN'], number_places_en, x['Job_EN']), 
        axis=1
    )
    sql_table["Title_CN"] = sql_table.apply(
        lambda x: create_chinese_title(x['Country_CN'], x['University_CN'], number_places_cn, x['Job_CN']), 
        axis=1
    )
    
    # 设置标签
    sql_table = set_label_columns(selected_row, sql_table, LABEL_COLUMNS)
    
    # 将NaN替换为None
    sql_table = sql_table.where(pd.notnull(sql_table), None)
    
    return sql_table


def generate_abbreviation(row):
    """生成职位缩写"""
    abbreviations = []
    
    if row['Master Student'] == '1':
        if row['Physical_Geo'] == '1' or row['GIS'] == '1' or row['RS'] == '1' or row['GNSS'] == '1':
            abbreviations.append("MSc")
        elif row['Human_Geo'] == '1' or row['Urban'] == '1':
            abbreviations.append("MA")
        else:
            abbreviations.append("MSc")
    
    if row['Doctoral Student'] == '1':
        abbreviations.append("PhD")
    if row['PostDoc'] == '1':
        abbreviations.append("PostDoc")
    if row['Research Assistant'] == '1':
        abbreviations.append("RA")
    if row['Competition'] == '1':
        abbreviations.append("Competition")
    if row['Conference'] == '1':
        abbreviations.append("Conference")
    if row['Summer School'] == '1':
        abbreviations.append("Summer School")
    if row['Workshop'] == '1':
        abbreviations.append("Workshop")
    
    return ", ".join(abbreviations) if abbreviations else ""


def create_combined_title(country_cn, university_cn):
    """创建组合标题"""
    if university_cn.startswith(country_cn):
        return university_cn
    else:
        return country_cn + university_cn


def is_contact_info_valid(contact_name, contact_email):
    """
    检查联系人信息是否有效
    Args:
        contact_name: 联系人姓名
        contact_email: 联系人邮箱
    Returns:
        bool: 如果联系人信息有效返回True，否则返回False
    """
    # 检查是否为空、NaN、None
    if pd.isna(contact_name) or pd.isna(contact_email):
        return False
    
    # 转换为字符串并去除空格
    name_str = str(contact_name).strip()
    email_str = str(contact_email).strip()
    
    # 检查是否为无效值："-"、"(-)"、空字符串等
    invalid_values = ['-', '(-)', '()', '', 'nan', 'None', 'null']
    
    if name_str in invalid_values or email_str in invalid_values:
        return False
    
    # 检查邮箱是否只包含括号和横线（如"(-)"）
    if email_str.replace('(', '').replace(')', '').replace('-', '').strip() == '':
        return False
    
    return True


def generate_wechat_group_text(row, abbreviation, new_event_id):
    """生成微信群消息文本"""
    combined_title = create_combined_title(row['Country_CN'], row['University_CN'])
    chinese_deadline = convert_date_to_chinese(row['Deadline'])
    
    abbreviations_list = abbreviation.split(", ")
    opportunity_str = "或".join(abbreviations_list)
    
    number_places = row['Number_Places']
    if number_places and number_places != '1':
        opportunity_str = f"{number_places}名{opportunity_str}"
    
    text = f"{combined_title}{row['Direction']}方向{opportunity_str}机会\n\n"
    
    # 检查联系人信息是否有效
    contact_name = row.get('Contact_Name', '')
    contact_email = row.get('Contact_Email', '')
    
    if is_contact_info_valid(contact_name, contact_email):
        text += f"{chinese_deadline}，有意者请联系{contact_name} ({contact_email})\n\n"
    else:
        text += f"{chinese_deadline}，有意者请见详情页\n\n"
    
    text += f"https://gisphere.info/post/{new_event_id}\n\n"
    
    labels = [row['Country_CN']]
    
    for abbr in abbreviations_list:
        if abbr in JOB_DICTIONARY:
            labels.append(JOB_DICTIONARY[abbr] + '机会')
    
    for subject in SUBJECT_DICTIONARY:
        if row.get(subject) == '1':
            labels.append(SUBJECT_DICTIONARY[subject])
    
    for i in range(1, 6):
        label_key = f'WX_Label{i}'
        if label_key in row and row[label_key]:
            labels.append(row[label_key])
    
    label_str = '标签：' + '；'.join(filter(None, labels))
    text += label_str
    
    return text


def convert_to_wechat_format(row, abbreviation):
    """转换为微信公众号格式"""
    template = ""
    
    university = row['University_CN'] if row['Country_CN'] in row['University_CN'] else row['Country_CN'] + row['University_CN']
    template += f"### {university}\n"
    
    direction = f"方向：{row['Direction']}"
    template += f"> {direction}\n"
    
    hosting_jobs = {'竞赛', '暑期学校', '学术会议', '研讨会'}
    recruiting_jobs = {'博士后', '研究助理'}
    
    # 处理多个类型的情况，将缩写转换为中文
    abbreviations_list = [abbr.strip() for abbr in abbreviation.split(", ")]
    job_cn_list = []
    for abbr in abbreviations_list:
        job_cn_list.append(JOB_DICTIONARY.get(abbr, abbr))
    
    # 用"或"连接多个中文职位
    job_cn = "或".join(job_cn_list)
    
    job_nb_str = row.get('Number_Places', '1')
    job_nb = int(job_nb_str) if str(job_nb_str).isdigit() else None
    
    # 根据第一个职位类型判断动词
    first_job_cn = job_cn_list[0]
    if first_job_cn in hosting_jobs:
        verb = "举办"
    elif first_job_cn in recruiting_jobs:
        verb = "招聘"
    else:
        verb = "招生"
    
    if job_nb is not None and job_nb > 1:
        template += f"{verb}类型：{job_cn}({job_nb}名)\n"
    else:
        template += f"{verb}类型：{job_cn}\n"
    
    deadline = convert_date_to_chinese(row['Deadline'])
    if "申请截止" in deadline:
        deadline = deadline.replace("申请截止", "").strip()
    template += f"申请截止：{deadline}\n"
    
    source = row['Source']
    template += "详细信息：\n" + source + "\n"
    
    # 检查联系人信息是否有效
    contact_name = row.get('Contact_Name', '')
    contact_email = row.get('Contact_Email', '')
    
    if is_contact_info_valid(contact_name, contact_email):
        contact = f"联系人：\n{contact_name} ({contact_email})\n"
    else:
        contact = "联系方式：\n见详情页\n"
    
    template += contact
    
    return template


def get_job_category(row):
    """
    获取职位类别
    Args:
        row: 数据行
    Returns:
        str: 职位类别（硕士招生、博士招生、博后招聘等）
    """
    # 按优先级检查职位类型
    if row.get('Master Student') == '1':
        return '硕士招生'
    if row.get('Doctoral Student') == '1':
        return '博士招生'
    if row.get('PostDoc') == '1':
        return '博后招聘'
    if row.get('Research Assistant') == '1':
        return '研究助理招聘'
    if row.get('Summer School') == '1':
        return '暑期学校'
    if row.get('Conference') == '1':
        return '学术会议'
    if row.get('Workshop') == '1':
        return '研讨会'
    if row.get('Competition') == '1':
        return '竞赛'
    
    return '其他'


def get_time_category(deadline):
    """
    获取时间类别
    Args:
        deadline: 截止日期
    Returns:
        str: 时间类别（尽快申请 或 本月及以后）
    """
    if deadline in ['Soon', '尽快申请', 'soon', 'SOON']:
        return '尽快申请'
    return '本月及以后'


def get_sort_priority(category):
    """
    获取职位类别的排序优先级
    Args:
        category: 职位类别
    Returns:
        int: 优先级数字（越小越靠前）
    """
    priority_map = {
        '硕士招生': 1,
        '博士招生': 2,
        '博后招聘': 3,
        '研究助理招聘': 4,
        '暑期学校': 5,
        '学术会议': 6,
        '研讨会': 7,
        '竞赛': 8,
        '其他': 9
    }
    return priority_map.get(category, 99)


def parse_deadline_for_sort(deadline_str):
    """
    解析截止日期用于排序
    Args:
        deadline_str: 截止日期字符串
    Returns:
        date: 日期对象，如果无法解析则返回一个很远的日期
    """
    if not deadline_str or deadline_str in ['Soon', '尽快申请', 'soon', 'SOON']:
        return date(9999, 12, 31)  # 尽快申请排在最前，用最大日期
    
    try:
        # 尝试解析日期
        if isinstance(deadline_str, date):
            return deadline_str
        if isinstance(deadline_str, pd.Timestamp):
            return deadline_str.date()
        
        # 尝试从中文格式解析
        if '年' in str(deadline_str) and '月' in str(deadline_str) and '日' in str(deadline_str):
            # 格式：2025年12月15日申请截止
            deadline_str = str(deadline_str).replace('申请截止', '').strip()
            parts = deadline_str.replace('年', '-').replace('月', '-').replace('日', '').split('-')
            if len(parts) == 3:
                return date(int(parts[0]), int(parts[1]), int(parts[2]))
        
        # 尝试标准日期格式
        return pd.to_datetime(deadline_str).date()
    except:
        return date(9999, 12, 31)

