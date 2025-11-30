"""
工具模块 - 通用辅助函数
"""
import pandas as pd
import inflect
from datetime import datetime, date, timedelta
from config import CHINA_TZ

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
    """计算本周周日到周六的日期范围"""
    current_date = date.today()
    start_of_week = current_date - timedelta(days=current_date.weekday() + 1)
    end_of_week = start_of_week + timedelta(days=6)
    return start_of_week.strftime('%Y-%m-%d'), end_of_week.strftime('%Y-%m-%d')


def adjust_data_to_columns(data, column_headers):
    """调整数据列数以匹配标题行"""
    adjusted_data = []
    for row in data:
        adjusted_row = row + [None] * (len(column_headers) - len(row))
        adjusted_data.append(adjusted_row)
    return adjusted_data

