"""
工具模块 - 通用辅助函数
"""
import pandas as pd
import inflect
from datetime import datetime, date, timedelta
from config import CHINA_TZ
from pypinyin import lazy_pinyin

# 初始化inflect引擎
p = inflect.engine()


def is_date(string):
    """检查字符串是否为日期格式"""
    try:
        pd.to_datetime(string)
        return True
    except (ValueError, TypeError):
        return False


def read_group_members(filename):
    """从文件中读取组员信息"""
    members = {}
    # 尝试多种编码方式
    encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']
    
    for encoding in encodings:
        try:
            with open(filename, 'r', encoding=encoding) as file:
                for line in file:
                    line = line.strip()
                    if ',' in line:
                        name, email = line.split(',', 1)
                        members[name.strip()] = email.strip()
            return members
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    # 如果所有编码都失败，抛出错误
    raise UnicodeDecodeError(
        'utf-8', b'', 0, 1, 
        f'无法读取文件 {filename}，请确保文件使用 UTF-8, GBK 或 GB2312 编码'
    )


def safe_convert_to_int(value):
    """安全地转换值为整数"""
    try:
        return int(value)
    except (ValueError, TypeError):
        return None


def number_to_english_words(number):
    """将数字转换为英文单词"""
    return p.number_to_words(number)


def number_to_chinese_words(number):
    """将数字转换为中文"""
    chinese_digits = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    if number < 10:
        return chinese_digits[number]
    elif number < 20:
        return "十" + (chinese_digits[number % 10] if number % 10 != 0 else "")
    else:
        return chinese_digits[number // 10] + "十" + (chinese_digits[number % 10] if number % 10 != 0 else "")


def convert_date_to_chinese(date_value):
    """将日期转换为中文格式"""
    if date_value == "Soon":
        return '尽快申请'
    elif pd.notnull(date_value) and not pd.isna(date_value):
        if isinstance(date_value, datetime):
            date_obj = date_value
        else:
            try:
                date_obj = datetime.strptime(str(date_value), '%Y-%m-%d')
            except ValueError:
                return "日期格式错误"
        
        year = date_obj.strftime('%Y')
        month = int(date_obj.strftime('%m'))
        day = int(date_obj.strftime('%d'))
        return f"{year}年{month}月{day}日申请截止"
    else:
        return '日期信息缺失'


def calculate_week_range():
    """计算两周的日期范围（周日到第二周周六，共14天）
    
    基于固定的两周周期，基准日期为 2025-11-30（周日）
    每个周期包含14天（两周）
    
    如果当前日期是周期的最后一天，则返回下一个周期的日期范围。
    
    Returns:
        tuple: (start_date, end_date) 格式为 'YYYY-MM-DD'
    
    Examples:
        如果今天在 2025-11-30 到 2025-12-13 之间，返回: ('2025-11-30', '2025-12-13')
        如果今天是 2025-12-27（周期的最后一天），返回: ('2025-12-28', '2026-01-10')
    """
    # 基准日期：2025-11-30（周日）- 第一个两周周期的起始日期
    base_date = date(2025, 11, 30)
    current_date = date.today()
    
    # 计算距离基准日期的天数
    days_diff = (current_date - base_date).days
    
    # 计算当前是第几个两周周期（每个周期14天）
    # 如果是负数（当前日期早于基准日期），需要向下取整
    if days_diff >= 0:
        cycle_number = days_diff // 14
    else:
        cycle_number = -(-days_diff // 14)  # 向上取整的负数
    
    # 计算当前周期的起始日期（周日）
    start_of_cycle = base_date + timedelta(days=cycle_number * 14)
    
    # 结束日期是起始日期后的第13天（共14天）
    end_of_cycle = start_of_cycle + timedelta(days=13)
    
    # 如果当前日期是周期的最后一天，则使用下一个周期
    if current_date == end_of_cycle:
        cycle_number += 1
        start_of_cycle = base_date + timedelta(days=cycle_number * 14)
        end_of_cycle = start_of_cycle + timedelta(days=13)
    
    return start_of_cycle.strftime('%Y-%m-%d'), end_of_cycle.strftime('%Y-%m-%d')


def calculate_period_number(week_start, week_end=None):
    """根据周期的起始日期计算期数（第XXX期）
    
    期数计算规则：
    - 基准日期 2025-11-30（第131期的起始日期）
    - Week: 2026-01-11 to 2026-01-24 是第134期
    - 每隔14天期数增加1
    
    Args:
        week_start: 周期起始日期，格式为 'YYYY-MM-DD' 或 date 对象
        week_end: 周期结束日期（可选，未使用）
    
    Returns:
        int: 期数
    
    Examples:
        calculate_period_number('2025-11-30') -> 131
        calculate_period_number('2026-01-11') -> 134
        calculate_period_number('2026-01-25') -> 135
    """
    # 基准日期：2025-11-30（周日）- 第131期的起始日期
    base_date = date(2025, 11, 30)
    base_period_number = 131  # 2025-11-30 是第131期
    
    # 处理输入日期格式
    if isinstance(week_start, str):
        start_date = datetime.strptime(week_start, '%Y-%m-%d').date()
    elif isinstance(week_start, datetime):
        start_date = week_start.date()
    else:
        start_date = week_start
    
    # 计算距离基准日期的天数
    days_diff = (start_date - base_date).days
    
    # 计算期数差异（每14天一期）
    period_diff = days_diff // 14
    
    return base_period_number + period_diff


def format_period_title(week_start, week_end):
    """格式化周期标题，包含期数
    
    Args:
        week_start: 周期起始日期，格式为 'YYYY-MM-DD'
        week_end: 周期结束日期，格式为 'YYYY-MM-DD'
    
    Returns:
        str: 格式化的周期标题，如 "海外资讯 134 | 2026.01.11 - 2026.01.24"
    
    Examples:
        format_period_title('2026-01-11', '2026-01-24') -> "海外资讯 134 | 2026.01.11 - 2026.01.24"
        format_period_title('2026-01-25', '2026-02-07') -> "海外资讯 135 | 2026.01.25 - 2026.02.07"
    """
    period_number = calculate_period_number(week_start)
    start_dot = week_start.replace('-', '.')
    end_dot = week_end.replace('-', '.')
    return f"海外资讯 {period_number} | {start_dot} - {end_dot}"


def adjust_data_to_columns(data, column_headers):
    """调整数据列数以匹配标题行"""
    adjusted_data = []
    for row in data:
        adjusted_row = row + [None] * (len(column_headers) - len(row))
        adjusted_data.append(adjusted_row)
    return adjusted_data


def column_index_to_letter(index):
    """
    将列索引（0-based）转换为Excel列字母
    例如: 0 -> A, 25 -> Z, 26 -> AA, 27 -> AB
    """
    result = ""
    while index >= 0:
        result = chr(65 + (index % 26)) + result
        index = index // 26 - 1
    return result


def get_pinyin_sort_key(text):
    """
    获取中文文本的拼音排序键
    Args:
        text: 中文文本
    Returns:
        str: 拼音字符串，用于排序
    """
    if not text:
        return ""
    return ''.join(lazy_pinyin(str(text)))

