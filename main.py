"""
主程序 - GISource 自动化系统
协调所有模块的运行流程
支持跨平台（Windows, macOS, Linux）
"""
# 强制使用IPv4连接（解决macOS/VPN环境下IPv6连接超时问题）
# 可通过环境变量 FORCE_IPV4=false 禁用此功能
# 默认启用，因为IPv4在所有平台上都可用，且能解决IPv6连接问题
import os
import socket

_FORCE_IPV4 = os.getenv('FORCE_IPV4', 'true').lower() == 'true'

if _FORCE_IPV4:
    _real_getaddrinfo = socket.getaddrinfo
    def ipv4_only_getaddrinfo(host, port, family=0, type=0, proto=0, flags=0):
        return _real_getaddrinfo(host, port, socket.AF_INET, type, proto, flags)
    socket.getaddrinfo = ipv4_only_getaddrinfo

import sys
import pandas as pd
import numpy as np
import warnings
from datetime import datetime

# 本地模块导入
from config import (
    CHINA_TZ, 
    UNFILLED_SHEET_ID,
    GROUP_MEMBERS_FILE,
    KEYS_DIR,
    REQUIRED_COLUMNS
)
from utils import (
    read_group_members, 
    adjust_data_to_columns,
    is_date,
    calculate_week_range,
    column_index_to_letter,
    format_period_title
)
from google_sheets import (
    fetch_data, 
    delete_rows_from_sheet,
    append_data_to_sheet,
    update_data_in_sheet
)
from database import (
    get_database_connection,
    clean_university_names,
    get_gisource_data,
    check_universities_exist,
    get_max_event_id,
    insert_event_to_database
)
from email_sender import (
    send_reminder_emails,
    send_error_notification,
    send_wechat_notification
)
from data_processor import (
    check_required_fields,
    create_sql_table,
    generate_abbreviation,
    generate_wechat_group_text,
    convert_to_wechat_format
)
from google_docs import add_wechat_content_to_doc, add_wechat_content_to_doc_sorted, ensure_current_period_exists
from logger import (
    log_program_run,
    log_program_start,
    log_program_end,
    restore_print_logging
)

# 禁止显示警告
warnings.filterwarnings('ignore')

# 配置pandas显示选项
pd.set_option('display.max_columns', None)


def print_banner():
    """打印程序横幅"""
    print("=" * 60)
    print("GISource 自动化系统".center(60))
    print("=" * 60)
    print()


def check_and_create_current_period():
    """
    检查并创建当前周期标题
    
    在程序开始时调用，确保当前周期的日期标题存在于 Google 文档中。
    如果不存在，则创建一个格式化的周期标题（居中、加粗、加大字号）。
    
    Returns:
        tuple: (week_start, week_end, period_created, message)
    """
    print("预检查: 确保当前周期标题存在...")
    log_program_run('PRE', '开始检查当前周期标题', 'info')
    
    # 计算当前周期的日期范围
    week_start, week_end = calculate_week_range()
    date_subtitle = format_period_title(week_start, week_end)
    
    print(f"   当前周期: {date_subtitle}")
    
    # 检查并创建周期标题
    try:
        period_created, message = ensure_current_period_exists(date_subtitle)
        if period_created:
            print(f"✓ {message}")
            log_program_run('PRE', message, 'success', {
                'week_start': week_start,
                'week_end': week_end,
                'period_created': True
            })
        else:
            print(f"✓ {message}")
            log_program_run('PRE', message, 'info', {
                'week_start': week_start,
                'week_end': week_end,
                'period_created': False
            })
    except Exception as e:
        print(f"⚠ 检查周期时出错: {e}")
        log_program_run('PRE', f'检查周期时出错: {e}', 'warning')
        period_created = False
        message = str(e)
    
    print()
    return week_start, week_end, period_created, message


def get_operator_name(group_members):
    """
    获取操作员姓名。
    
    逻辑：
    1. 读取缓存文件（keys/operator_cache.json）
    2. 若缓存存在且距上次运行不超过30分钟，则直接使用缓存中的姓名
    3. 否则，提示用户手动输入姓名
    4. 输入的姓名必须与 group_members 中的姓名完全匹配，否则程序退出
    
    Args:
        group_members (dict): 已读取的组员信息（姓名 -> 邮箱）
    
    Returns:
        str: 操作员姓名，若验证失败则返回 None
    """
    import json
    
    cache_file = os.path.join(KEYS_DIR, 'operator_cache.json')
    threshold_seconds = 30 * 60  # 30分钟
    
    # 尝试读取缓存
    cached_name = None
    if os.path.exists(cache_file):
        try:
            with open(cache_file, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            last_run_str = cache_data.get('last_run')
            cached_name = cache_data.get('operator_name')
            if last_run_str and cached_name:
                last_run_dt = datetime.fromisoformat(last_run_str)
                now_dt = datetime.now()
                # 兼容有时区/无时区的情况
                if last_run_dt.tzinfo is not None:
                    now_dt = datetime.now(last_run_dt.tzinfo)
                elapsed = (now_dt - last_run_dt).total_seconds()
                if elapsed <= threshold_seconds:
                    print(f"✓ 距上次运行 {int(elapsed // 60)} 分 {int(elapsed % 60)} 秒，自动使用缓存操作员: {cached_name}")
                    return cached_name
                else:
                    print(f"  缓存已过期（距上次运行 {int(elapsed // 60)} 分钟），需要重新输入。")
        except Exception:
            pass  # 缓存损坏，忽略并要求重新输入
    
    # 提示用户输入姓名
    operator_input = input("请输入您的姓名: ").strip()
    
    if operator_input not in group_members:
        print(f"\n✗ 错误：姓名 '{operator_input}' 不在组员名单中。程序退出。")
        return None
    
    # 写入缓存
    try:
        os.makedirs(KEYS_DIR, exist_ok=True)
        cache_data = {
            'operator_name': operator_input,
            'last_run': datetime.now().isoformat()
        }
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"  ⚠ 无法写入缓存文件: {e}")
    
    print(f"✓ 操作员已确认: {operator_input}")
    return operator_input


def load_and_clean_data():
    """加载并清理Google Sheets数据"""
    print("步骤 1: 从Google Sheets获取数据...")
    log_program_run('1', '开始从Google Sheets获取数据', 'info')
    
    # 获取Unfilled数据
    unfilled_range_name = 'Unfilled'
    unfilled_headers = fetch_data(unfilled_range_name)[0]
    unfilled_data_raw = fetch_data(unfilled_range_name)[1:]
    unfilled_data_adjusted = adjust_data_to_columns(unfilled_data_raw, unfilled_headers)
    unfilled_data = pd.DataFrame(unfilled_data_adjusted, columns=unfilled_headers)
    
    # 获取Filled数据（用于检查重复）
    filled_range_name = 'Filled'
    filled_headers = fetch_data(filled_range_name)[0]
    filled_data_raw = fetch_data(filled_range_name)[1:]
    filled_data_adjusted = adjust_data_to_columns(filled_data_raw, filled_headers)
    filled_data = pd.DataFrame(filled_data_adjusted, columns=filled_headers)
    
    # 保存原始Deadline值用于重复检查（在日期转换之前）
    unfilled_deadline_original = unfilled_data['Deadline'].astype(str).str.strip()
    filled_deadline_original = filled_data['Deadline'].astype(str).str.strip() if 'Deadline' in filled_data.columns else pd.Series()
    
    # ===== 条件1: 删除过期行 =====
    # 只对非"Soon"的值进行日期转换
    mask_not_soon = ~unfilled_data['Deadline'].astype(str).str.strip().isin(['Soon', 'soon', 'SOON', '尽快申请'])
    unfilled_data.loc[mask_not_soon, 'Deadline'] = pd.to_datetime(
        unfilled_data.loc[mask_not_soon, 'Deadline'], 
        errors='coerce'
    )
    
    # 获取当前日期（时区无关的date对象）
    now = datetime.now(CHINA_TZ).date()
    
    # 找出过期的行（排除"Soon"行）
    expired_rows = unfilled_data.index[
        (unfilled_data['Deadline'].notna()) & 
        (unfilled_data['Deadline'].apply(lambda x: x.date() if hasattr(x, 'date') else None) < now)
    ].tolist()
    
    # ===== 条件2: 删除与Filled重复的行 =====
    # 检查字段: Deadline, Direction, University_EN, Contact_Email
    duplicate_check_columns = ['Deadline', 'Direction', 'University_EN', 'Contact_Email']
    duplicate_rows = []
    
    # 确保所有检查字段都存在于两个表中
    if all(col in unfilled_data.columns for col in duplicate_check_columns) and \
       all(col in filled_data.columns for col in duplicate_check_columns):
        
        # 创建Filled表的比较集合（使用原始字符串值进行比较）
        filled_compare_set = set()
        for idx, row in filled_data.iterrows():
            # 统一格式化用于比较
            deadline_val = str(row['Deadline']).strip() if pd.notna(row['Deadline']) else ''
            direction_val = str(row['Direction']).strip() if pd.notna(row['Direction']) else ''
            university_val = str(row['University_EN']).strip() if pd.notna(row['University_EN']) else ''
            email_val = str(row['Contact_Email']).strip() if pd.notna(row['Contact_Email']) else ''
            
            filled_compare_set.add((deadline_val, direction_val, university_val, email_val))
        
        # 检查Unfilled中的每一行是否在Filled中存在
        for idx in unfilled_data.index:
            # 使用原始保存的Deadline值（未转换的字符串）
            deadline_val = unfilled_deadline_original.iloc[idx] if idx < len(unfilled_deadline_original) else ''
            direction_val = str(unfilled_data.at[idx, 'Direction']).strip() if pd.notna(unfilled_data.at[idx, 'Direction']) else ''
            university_val = str(unfilled_data.at[idx, 'University_EN']).strip() if pd.notna(unfilled_data.at[idx, 'University_EN']) else ''
            email_val = str(unfilled_data.at[idx, 'Contact_Email']).strip() if pd.notna(unfilled_data.at[idx, 'Contact_Email']) else ''
            
            unfilled_tuple = (deadline_val, direction_val, university_val, email_val)
            
            if unfilled_tuple in filled_compare_set:
                duplicate_rows.append(idx)
                print(f"   发现重复数据: Deadline={deadline_val}, Direction={direction_val}, University_EN={university_val}")
    
    # ===== 合并两种删除条件 =====
    all_rows_to_delete = list(set(expired_rows + duplicate_rows))
    
    # 转换为Google Sheets行号（索引+1是因为表头，再+1是因为索引从0开始）
    rows_to_delete_sheet = [x + 2 for x in all_rows_to_delete]
    
    if rows_to_delete_sheet:
        # 分别记录删除原因
        expired_count = len(expired_rows)
        duplicate_count = len(duplicate_rows)
        # 去除重叠的行
        overlap_count = len(set(expired_rows) & set(duplicate_rows))
        
        print(f"   过期行数: {expired_count}, 重复行数: {duplicate_count}, 重叠行数: {overlap_count}")
        print(f"   总共需要删除: {len(all_rows_to_delete)} 行")
        
        delete_rows_from_sheet(UNFILLED_SHEET_ID, rows_to_delete_sheet)
        log_program_run('1', f'删除了 {len(all_rows_to_delete)} 行数据（过期: {expired_count}, 重复: {duplicate_count}）', 'info', {
            'deleted_rows': rows_to_delete_sheet,
            'expired_count': expired_count,
            'duplicate_count': duplicate_count
        })
        # 重新获取数据
        unfilled_data_raw = fetch_data(unfilled_range_name)[1:]
        unfilled_data_adjusted = adjust_data_to_columns(unfilled_data_raw, unfilled_headers)
        unfilled_data = pd.DataFrame(unfilled_data_adjusted, columns=unfilled_headers)
    else:
        print("   没有过期或重复的行需要删除")
        log_program_run('1', '没有过期或重复的行需要删除', 'info')
    
    # 获取Filled数据
    filled_range_name = 'Filled'
    filled_headers = fetch_data(filled_range_name)[0]
    filled_data_raw = fetch_data(filled_range_name)[1:]
    filled_data_adjusted = adjust_data_to_columns(filled_data_raw, filled_headers)
    filled_data = pd.DataFrame(filled_data_adjusted, columns=filled_headers)
    
    print("✓ 数据加载完成\n")
    log_program_run('1', '数据加载完成', 'success', {
        'unfilled_rows': len(unfilled_data),
        'filled_rows': len(filled_data)
    })
    return unfilled_data, filled_data, unfilled_range_name, filled_range_name


def update_university_info(unfilled_data):
    """更新大学中文名称信息"""
    print("步骤 2: 更新大学中文名称...")
    log_program_run('2', '开始更新大学中文名称', 'info')
    
    conn, cursor = get_database_connection()
    if not conn:
        print("⚠ 无法连接数据库，跳过大学信息更新")
        log_program_run('2', '无法连接数据库，跳过大学信息更新', 'warning')
        return unfilled_data
    
    try:
        # 清理大学名称
        clean_university_names(cursor, conn)
        
        # 获取GISource数据
        gisource_data = get_gisource_data(cursor)
        gisource_df = pd.DataFrame(gisource_data, columns=['University_EN', 'University_CN', 'Country_CN'])
        
        # 更新unfilled_data中的大学信息
        modified_rows = []
        for index, row in unfilled_data.iterrows():
            if pd.notna(row['University_EN']) and (pd.isna(row['University_CN']) or pd.isna(row['Country_CN'])):
                matches = gisource_df[gisource_df['University_EN'] == row['University_EN']]
                if not matches.empty:
                    latest_match = matches.iloc[-1]
                    if pd.isna(row['University_CN']):
                        unfilled_data.at[index, 'University_CN'] = latest_match['University_CN']
                        modified_rows.append(index)
                    if pd.isna(row['Country_CN']):
                        unfilled_data.at[index, 'Country_CN'] = latest_match['Country_CN']
                        modified_rows.append(index)
        
        # 更新修改的行到Google Sheets
        modified_rows = list(set(modified_rows))
        for row in modified_rows:
            range_name = f'Unfilled!A{row + 2}:Z{row + 2}'
            update_data = [unfilled_data.iloc[row].tolist()]
            update_data_in_sheet(range_name, update_data)
        
        print(f"✓ 更新了 {len(modified_rows)} 行大学信息\n")
        log_program_run('2', f'更新了 {len(modified_rows)} 行大学信息', 'success', {
            'modified_rows_count': len(modified_rows)
        })
        
    finally:
        cursor.close()
        conn.close()
    
    return unfilled_data


def check_new_universities(filled_data):
    """检查并添加新大学到Universities表"""
    print("步骤 3: 检查新大学...")
    log_program_run('3', '开始检查新大学', 'info')
    
    conn, cursor = get_database_connection()
    if not conn:
        print("⚠ 无法连接数据库，跳过新大学检查")
        log_program_run('3', '无法连接数据库，跳过新大学检查', 'warning')
        return
    
    try:
        unique_universities = filled_data[['University_EN', 'University_CN', 'Country_CN']].drop_duplicates(subset=['University_EN'])
        
        if unique_universities.empty:
            print("✓ 没有新大学需要检查\n")
            return
        
        # 检查已存在的大学
        existing_universities_set = check_universities_exist(cursor, unique_universities['University_EN'].tolist())
        
        # 找出新大学
        new_universities = unique_universities[~unique_universities['University_EN'].isin(existing_universities_set)]
        
        if not new_universities.empty:
            # 获取Universities工作表数据
            universities_data = fetch_data('Universities')
            if universities_data:
                universities_headers = universities_data[0]
                universities_existing = pd.DataFrame(universities_data[1:], columns=universities_headers)
            else:
                universities_existing = pd.DataFrame(columns=['University_EN', 'University_CN', 'Country_CN'])
            
            # 找出最终的新大学
            final_new_universities = new_universities[~new_universities['University_EN'].isin(universities_existing['University_EN'])]
            
            if not final_new_universities.empty:
                final_new_universities_data = final_new_universities.values.tolist()
                append_data_to_sheet('Universities', final_new_universities_data)
                print(f"✓ 添加了 {len(final_new_universities)} 所新大学\n")
                log_program_run('3', f'添加了 {len(final_new_universities)} 所新大学', 'success', {
                    'new_universities_count': len(final_new_universities)
                })
            else:
                print("✓ 没有新大学需要添加\n")
                log_program_run('3', '没有新大学需要添加', 'info')
        else:
            print("✓ 没有新大学需要添加\n")
            log_program_run('3', '没有新大学需要添加', 'info')
    
    finally:
        cursor.close()
        conn.close()


def select_row_to_process(unfilled_data):
    """选择要处理的行"""
    print("步骤 4: 选择要处理的数据...")
    log_program_run('4', '开始选择要处理的数据', 'info')
    
    # 过滤有效数据
    filtered_data = unfilled_data[
        (unfilled_data['Error'] == 'N') &
        unfilled_data['Verifier'].notnull() &
        (unfilled_data['Verifier'] != 'LLM')
    ]
    
    if filtered_data.empty:
        print("⚠ 没有可处理的数据")
        log_program_run('4', '没有可处理的数据', 'warning')
        return None, filtered_data
    
    # 转换Deadline为日期
    now = datetime.now(CHINA_TZ).date()
    
    # 识别Soon行（在转换为datetime之前）
    deadline_str = filtered_data['Deadline'].astype(str).str.strip()
    soon_rows = filtered_data[deadline_str.isin(['Soon', 'soon', 'SOON', '尽快申请'])]
    
    # 移除Soon行进行截止日期计算
    deadline_data = filtered_data[~filtered_data.index.isin(soon_rows.index)]
    deadline_data['Deadline'] = pd.to_datetime(deadline_data['Deadline'], errors='coerce')
    
    # 找出最近的截止日期
    nearest_deadline = deadline_data[
        (deadline_data['Deadline'].dt.date >= now) & (deadline_data['Deadline'].notna())
    ].nsmallest(1, 'Deadline')
    
    index_choices = []
    weights = []
    
    # Soon行有80%的概率
    if not soon_rows.empty:
        random_soon_index = soon_rows.sample(n=1).index[0]
        index_choices.append(random_soon_index)
        weights.append(0.8)
    
    # 最近截止日期有10%的概率
    if not nearest_deadline.empty:
        index_choices.append(nearest_deadline.index[0])
        weights.append(0.1)
    
    # 随机有效行有10%的概率
    valid_rows = deadline_data[deadline_data['Deadline'].dt.date >= now]
    if not valid_rows.empty:
        random_valid_index = valid_rows.sample(n=1).index[0]
        index_choices.append(random_valid_index)
        weights.append(0.1)
    
    # 调整权重
    if soon_rows.empty:
        weights = [0.9, 0.1]
    
    # 归一化权重
    weights = [float(w) / sum(weights) for w in weights]
    
    # 选择行
    if index_choices:
        selected_index = np.random.choice(index_choices, p=weights)
        selected_row = filtered_data.loc[selected_index]
        selected_row = pd.DataFrame(selected_row).transpose()
        print("✓ 已选择数据行\n")
        log_program_run('4', '已选择数据行', 'success', {
            'selected_index': int(selected_index),
            'source': selected_row['Source'].iloc[0] if 'Source' in selected_row.columns else None,
            'direction': selected_row['Direction'].iloc[0] if 'Direction' in selected_row.columns else None
        })
        return selected_row, filtered_data
    else:
        print("⚠ 没有有效数据可选择")
        log_program_run('4', '没有有效数据可选择', 'warning')
        return None, filtered_data


def validate_selected_row(selected_row, group_members, unfilled_data):
    """验证选中的行是否有错误"""
    print("步骤 5: 验证数据完整性...")
    log_program_run('5', '开始验证数据完整性', 'info')
    
    # 检查必填字段
    error_value = check_required_fields(selected_row)
    selected_row['Error'] = error_value
    
    current_date_china = datetime.now(CHINA_TZ).strftime("%Y-%m-%d")
    
    if error_value == '1':
        # 安全地获取字段内容，如果为空则使用默认值
        source_content = selected_row["Source"].iloc[0] if pd.notna(selected_row["Source"].iloc[0]) else "未填写"
        university_content = selected_row["University_CN"].iloc[0] if pd.notna(selected_row["University_CN"].iloc[0]) else "未填写"
        direction_content = selected_row["Direction"].iloc[0] if pd.notna(selected_row["Direction"].iloc[0]) else "未填写"
        verifier_name = selected_row["Verifier"].iloc[0] if pd.notna(selected_row["Verifier"].iloc[0]) else "未知"
        
        # 更新Unfilled表中的Error列为'1'
        source_to_update = selected_row['Source'].values[0]
        direction_to_update = selected_row['Direction'].values[0]
        
        # 找到对应的行索引
        matching_rows = unfilled_data.index[
            (unfilled_data['Source'] == source_to_update) & 
            (unfilled_data['Direction'] == direction_to_update)
        ].tolist()
        
        if matching_rows:
            row_index = matching_rows[0]
            # 获取Error列的索引（假设在表头中）
            unfilled_headers = fetch_data('Unfilled')[0]
            if 'Error' in unfilled_headers:
                error_col_index = unfilled_headers.index('Error')
                error_col_letter = column_index_to_letter(error_col_index)
                # 更新Google Sheets中的Error列（行号+2是因为：+1是表头，+1是从0开始）
                range_name = f'Unfilled!{error_col_letter}{row_index + 2}'
                update_data_in_sheet(range_name, [['1']])
                print(f"✓ 已将Error列更新为'1' (位置: {range_name})")
        else:
            print(f"⚠ 警告：无法在Unfilled表中找到匹配的行来更新Error列")
        
        if verifier_name in group_members:
            receiver_email = group_members[verifier_name]
            email_sent = send_error_notification(
                receiver_email, verifier_name, source_content, 
                university_content, direction_content, current_date_china
            )
            if email_sent:
                print(f"⚠ 发现错误，已发送邮件通知 {verifier_name}")
                log_program_run('5', f'发现错误，已发送邮件通知 {verifier_name}', 'error', {
                    'verifier': verifier_name,
                    'source': source_content,
                    'university': university_content,
                    'direction': direction_content
                })
            else:
                print(f"⚠ 发现错误，但邮件发送失败！请手动通知 {verifier_name}")
                log_program_run('5', f'发现错误，邮件发送失败: {verifier_name}', 'error', {
                    'verifier': verifier_name,
                    'source': source_content,
                    'university': university_content,
                    'direction': direction_content,
                    'email_sent': False
                })
            return False
        else:
            print(f"⚠ 验证者 {verifier_name} 不在组员列表中，跳过邮件通知")
            log_program_run('5', f'验证者 {verifier_name} 不在组员列表中，跳过邮件通知', 'warning')
            return False
    
    print("✓ 数据验证通过\n")
    log_program_run('5', '数据验证通过', 'success')
    return True


def process_and_insert_to_database(selected_row):
    """处理数据并插入到数据库"""
    print("步骤 6: 插入数据到数据库...")
    log_program_run('6', '开始插入数据到数据库', 'info')
    
    conn, cursor = get_database_connection()
    if not conn:
        print("⚠ 无法连接数据库")
        log_program_run('6', '无法连接数据库', 'error')
        return None
    
    try:
        # 获取新的Event_ID
        last_event_id = get_max_event_id(cursor)
        new_event_id = last_event_id + 1
        
        # 创建SQL表格
        sql_table = create_sql_table(selected_row, new_event_id)
        
        # 插入数据
        table_name = 'GISource'  # 或 'Coding_Test' 用于测试
        success = insert_event_to_database(cursor, conn, sql_table, table_name)
        
        if success:
            print(f"✓ 成功插入数据，Event_ID: {new_event_id}\n")
            log_program_run('6', f'成功插入数据，Event_ID: {new_event_id}', 'success', {
                'event_id': new_event_id
            })
            return new_event_id
        else:
            print("⚠ 数据插入失败")
            log_program_run('6', '数据插入失败', 'error')
            return None
    
    finally:
        cursor.close()
        conn.close()


def update_google_sheets(selected_row, unfilled_range_name, filled_range_name):
    """更新Google Sheets"""
    print("步骤 7: 更新Google Sheets...")
    log_program_run('7', '开始更新Google Sheets', 'info')
    
    # 从Unfilled中删除
    source_to_delete = selected_row['Source'].values[0]
    direction_to_delete = selected_row['Direction'].values[0]
    
    # 重新获取unfilled_data以找到正确的索引
    unfilled_headers = fetch_data(unfilled_range_name)[0]
    unfilled_data_raw = fetch_data(unfilled_range_name)[1:]
    unfilled_data_adjusted = adjust_data_to_columns(unfilled_data_raw, unfilled_headers)
    unfilled_data = pd.DataFrame(unfilled_data_adjusted, columns=unfilled_headers)
    
    rows_to_delete = unfilled_data.index[
        (unfilled_data['Source'] == source_to_delete) & 
        (unfilled_data['Direction'] == direction_to_delete)
    ].tolist()
    
    rows_to_delete = [x + 1 for x in rows_to_delete]
    
    if rows_to_delete:
        delete_rows_from_sheet(UNFILLED_SHEET_ID, rows_to_delete)
    
    # 添加到Filled
    def convert_value(value):
        if isinstance(value, pd.Timestamp):
            return value.strftime('%Y-%m-%d')
        else:
            return value
    
    converted_row = selected_row.applymap(convert_value)
    data_to_append = [converted_row.iloc[0].tolist()]
    append_data_to_sheet(filled_range_name, data_to_append)
    
    print("✓ Google Sheets更新完成\n")
    log_program_run('7', 'Google Sheets更新完成', 'success', {
        'deleted_from_unfilled': len(rows_to_delete),
        'added_to_filled': True
    })


def generate_and_send_wechat_message(selected_row, new_event_id, operator, group_members):
    """生成并发送微信群消息"""
    print("步骤 8: 生成微信消息...")
    log_program_run('8', '开始生成微信消息', 'info')
    
    # 生成缩写
    abbreviation = generate_abbreviation(selected_row.iloc[0])
    
    if not abbreviation:
        print("⚠ 无法生成职位缩写")
        log_program_run('8', '无法生成职位缩写', 'error')
        return None
    
    # 生成微信群消息
    text_output = generate_wechat_group_text(selected_row.iloc[0], abbreviation, new_event_id)
    
    print("微信群消息:")
    print("-" * 60)
    print(text_output)
    print("-" * 60)
    print()
    
    # 发送邮件通知
    direction_content = selected_row['Direction'].values[0]
    current_date_china = datetime.now(CHINA_TZ).strftime("%Y-%m-%d")
    
    recipient_name = operator if operator in group_members else "GISphere"
    receiver_email = group_members.get(recipient_name, list(group_members.values())[0])
    
    email_sent = send_wechat_notification(receiver_email, recipient_name, text_output, direction_content, current_date_china)
    if email_sent:
        print(f"✓ 已发送微信消息通知到 {recipient_name}\n")
        log_program_run('8', f'已发送微信消息通知到 {recipient_name}', 'success', {
            'recipient': recipient_name,
            'event_id': new_event_id,
            'abbreviation': abbreviation
        })
    else:
        print(f"⚠ 邮件发送失败，但微信消息已生成。请手动通知 {recipient_name}\n")
        log_program_run('8', f'邮件发送失败: {recipient_name}', 'warning', {
            'recipient': recipient_name,
            'event_id': new_event_id,
            'abbreviation': abbreviation,
            'email_sent': False
        })
    
    return text_output, abbreviation


def add_to_wechat_official_account(selected_row, abbreviation):
    """添加到微信公众号文档"""
    print("步骤 9: 添加到微信公众号文档...")
    log_program_run('9', '开始添加到微信公众号文档', 'info')
    
    # 生成微信公众号格式
    wechat_template_output = convert_to_wechat_format(selected_row.iloc[0], abbreviation)
    
    # 计算周范围
    week_start, week_end = calculate_week_range()
    date_subtitle = format_period_title(week_start, week_end)
    
    # 添加到文档（使用排序功能）
    try:
        result_message = add_wechat_content_to_doc_sorted(
            wechat_template_output, 
            date_subtitle, 
            selected_row.iloc[0]
        )
    except Exception as e:
        print(f"⚠ 排序插入失败，使用简单追加: {e}")
        result_message = add_wechat_content_to_doc(wechat_template_output, date_subtitle)
    
    print(f"✓ {result_message}\n")
    log_program_run('9', f'添加到微信公众号文档完成: {result_message}', 'success', {
        'result': result_message
    })
    
    print("微信公众号内容:")
    print("-" * 60)
    print(wechat_template_output)
    print("-" * 60)
    print()


def send_wechat_email_notification(selected_row, new_event_id, operator, group_members, abbreviation):
    """发送微信群消息邮件通知（在写入文档之后执行）"""
    print("步骤 10: 发送微信群消息邮件通知...")
    log_program_run('10', '开始发送微信群消息邮件通知', 'info')
    
    # 生成微信群消息
    text_output = generate_wechat_group_text(selected_row.iloc[0], abbreviation, new_event_id)
    
    print("微信群消息:")
    print("-" * 60)
    print(text_output)
    print("-" * 60)
    print()
    
    # 发送邮件通知
    direction_content = selected_row['Direction'].values[0]
    current_date_china = datetime.now(CHINA_TZ).strftime("%Y-%m-%d")
    
    recipient_name = operator if operator in group_members else "GISphere"
    receiver_email = group_members.get(recipient_name, list(group_members.values())[0])
    
    email_sent = send_wechat_notification(receiver_email, recipient_name, text_output, direction_content, current_date_china)
    if email_sent:
        print(f"✓ 已发送微信消息通知到 {recipient_name}\n")
        log_program_run('10', f'已发送微信消息通知到 {recipient_name}', 'success', {
            'recipient': recipient_name,
            'event_id': new_event_id,
            'abbreviation': abbreviation
        })
    else:
        print(f"⚠ 邮件发送失败，但微信消息已生成。请手动通知 {recipient_name}\n")
        log_program_run('10', f'邮件发送失败: {recipient_name}', 'warning', {
            'recipient': recipient_name,
            'event_id': new_event_id,
            'abbreviation': abbreviation,
            'email_sent': False
        })


def main():
    """主函数"""
    # 设置print输出日志（必须在所有print之前）
    tee_output = log_program_start()
    
    try:
        print_banner()
        # 读取组员信息
        group_members = read_group_members(GROUP_MEMBERS_FILE)
        
        # 获取并验证操作员姓名（必须在组员名单中，且支持30分钟内免重复输入）
        operator = get_operator_name(group_members)
        if operator is None:
            log_program_end(success=False, error_message='操作员验证失败')
            return
        
        log_program_run('INIT', f'程序初始化完成，操作员: {operator}', 'info', {
            'operator': operator,
            'group_members_count': len(group_members)
        })
        
        # 预检查: 确保当前周期标题存在于 Google 文档中
        week_start, week_end, period_created, period_message = check_and_create_current_period()
        
        # 步骤1: 加载数据
        unfilled_data, filled_data, unfilled_range_name, filled_range_name = load_and_clean_data()
        
        # 步骤2: 更新大学信息
        unfilled_data = update_university_info(unfilled_data)
        
        # 步骤3: 检查新大学
        check_new_universities(filled_data)
        
        # 步骤4: 选择要处理的行
        selected_row, filtered_data = select_row_to_process(unfilled_data)
        
        if selected_row is None:
            print("\n没有可处理的数据，发送提醒邮件...")
            log_program_run('MAIN', '没有可处理的数据，发送提醒邮件', 'info')
            send_reminder_emails(group_members)
            print("程序结束")
            log_program_end(success=True)
            return
        
        # 步骤5: 验证数据
        if not validate_selected_row(selected_row, group_members, unfilled_data):
            print("程序结束")
            log_program_end(success=False, error_message='数据验证失败')
            return
        
        # 步骤6: 插入到数据库
        new_event_id = process_and_insert_to_database(selected_row)
        
        if new_event_id is None:
            print("程序结束")
            log_program_end(success=False, error_message='数据库插入失败')
            return
        
        # 步骤7: 更新Google Sheets
        update_google_sheets(selected_row, unfilled_range_name, filled_range_name)
        
        # 步骤8: 生成微信群消息内容和缩写（不发送邮件）
        abbreviation = generate_abbreviation(selected_row.iloc[0])
        
        if not abbreviation:
            print("⚠ 无法生成职位缩写")
            log_program_run('8', '无法生成职位缩写', 'error')
            print("程序结束")
            log_program_end(success=False, error_message='微信消息生成失败')
            return
        
        # 步骤9: 添加到微信公众号
        add_to_wechat_official_account(selected_row, abbreviation)
        
        # 步骤10: 发送微信群消息邮件通知（在写入文档之后）
        send_wechat_email_notification(selected_row, new_event_id, operator, group_members, abbreviation)
        
        print("=" * 60)
        print("所有步骤完成！".center(60))
        print("=" * 60)
        log_program_end(success=True)
    
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
        log_program_end(success=False, error_message='用户中断')
    
    except Exception as e:
        print(f"\n⚠ 发生错误: {e}")
        import traceback
        error_trace = traceback.format_exc()
        log_program_end(success=False, error_message=str(e))
        log_program_run('ERROR', f'程序异常: {str(e)}', 'error', {
            'traceback': error_trace
        })
        traceback.print_exc()
    
    finally:
        # 确保恢复stdout和stderr（在所有情况下都会执行）
        restore_print_logging(tee_output)


if __name__ == "__main__":
    main()

