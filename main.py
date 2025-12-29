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
    REQUIRED_COLUMNS
)
from utils import (
    read_group_members, 
    adjust_data_to_columns,
    is_date,
    calculate_week_range,
    column_index_to_letter
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
from google_docs import add_wechat_content_to_doc, add_wechat_content_to_doc_sorted

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


def get_operator_name():
    """获取操作员姓名"""
    # 可以从环境变量或配置文件读取，这里使用默认值
    # 如果需要动态输入，可以使用: return input("请输入您的姓名: ")
    return "李思高"  # 默认操作员


def load_and_clean_data():
    """加载并清理Google Sheets数据"""
    print("步骤 1: 从Google Sheets获取数据...")
    
    # 获取Unfilled数据
    unfilled_range_name = 'Unfilled'
    unfilled_headers = fetch_data(unfilled_range_name)[0]
    unfilled_data_raw = fetch_data(unfilled_range_name)[1:]
    unfilled_data_adjusted = adjust_data_to_columns(unfilled_data_raw, unfilled_headers)
    unfilled_data = pd.DataFrame(unfilled_data_adjusted, columns=unfilled_headers)
    
    # 转换截止日期并删除过期行
    # 先保存原始Deadline值（包括"Soon"字符串）
    original_deadlines = unfilled_data['Deadline'].copy()
    
    # 只对非"Soon"的值进行日期转换
    mask_not_soon = ~unfilled_data['Deadline'].astype(str).str.strip().isin(['Soon', 'soon', 'SOON', '尽快申请'])
    unfilled_data.loc[mask_not_soon, 'Deadline'] = pd.to_datetime(
        unfilled_data.loc[mask_not_soon, 'Deadline'], 
        errors='coerce'
    )
    
    # 获取当前日期（时区无关的date对象）
    now = datetime.now(CHINA_TZ).date()
    
    # 找出过期的行（排除"Soon"行）
    rows_to_delete = unfilled_data.index[
        (unfilled_data['Deadline'].notna()) & 
        (unfilled_data['Deadline'].apply(lambda x: x.date() if hasattr(x, 'date') else None) < now)
    ].tolist()
    rows_to_delete = [x + 1 for x in rows_to_delete]
    
    if rows_to_delete:
        delete_rows_from_sheet(UNFILLED_SHEET_ID, rows_to_delete)
        # 重新获取数据
        unfilled_data_raw = fetch_data(unfilled_range_name)[1:]
        unfilled_data_adjusted = adjust_data_to_columns(unfilled_data_raw, unfilled_headers)
        unfilled_data = pd.DataFrame(unfilled_data_adjusted, columns=unfilled_headers)
    else:
        print("No expired rows to delete.")
    
    # 获取Filled数据
    filled_range_name = 'Filled'
    filled_headers = fetch_data(filled_range_name)[0]
    filled_data_raw = fetch_data(filled_range_name)[1:]
    filled_data_adjusted = adjust_data_to_columns(filled_data_raw, filled_headers)
    filled_data = pd.DataFrame(filled_data_adjusted, columns=filled_headers)
    
    print("✓ 数据加载完成\n")
    return unfilled_data, filled_data, unfilled_range_name, filled_range_name


def update_university_info(unfilled_data):
    """更新大学中文名称信息"""
    print("步骤 2: 更新大学中文名称...")
    
    conn, cursor = get_database_connection()
    if not conn:
        print("⚠ 无法连接数据库，跳过大学信息更新")
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
        
    finally:
        cursor.close()
        conn.close()
    
    return unfilled_data


def check_new_universities(filled_data):
    """检查并添加新大学到Universities表"""
    print("步骤 3: 检查新大学...")
    
    conn, cursor = get_database_connection()
    if not conn:
        print("⚠ 无法连接数据库，跳过新大学检查")
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
            else:
                print("✓ 没有新大学需要添加\n")
        else:
            print("✓ 没有新大学需要添加\n")
    
    finally:
        cursor.close()
        conn.close()


def select_row_to_process(unfilled_data):
    """选择要处理的行"""
    print("步骤 4: 选择要处理的数据...")
    
    # 过滤有效数据
    filtered_data = unfilled_data[
        (unfilled_data['Error'] == 'N') &
        unfilled_data['Verifier'].notnull() &
        (unfilled_data['Verifier'] != 'LLM')
    ]
    
    if filtered_data.empty:
        print("⚠ 没有可处理的数据")
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
        return selected_row, filtered_data
    else:
        print("⚠ 没有有效数据可选择")
        return None, filtered_data


def validate_selected_row(selected_row, group_members, unfilled_data):
    """验证选中的行是否有错误"""
    print("步骤 5: 验证数据完整性...")
    
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
            send_error_notification(
                receiver_email, verifier_name, source_content, 
                university_content, direction_content, current_date_china
            )
            print(f"⚠ 发现错误，已发送邮件通知 {verifier_name}")
            return False
        else:
            print(f"⚠ 验证者 {verifier_name} 不在组员列表中，跳过邮件通知")
            return False
    
    print("✓ 数据验证通过\n")
    return True


def process_and_insert_to_database(selected_row):
    """处理数据并插入到数据库"""
    print("步骤 6: 插入数据到数据库...")
    
    conn, cursor = get_database_connection()
    if not conn:
        print("⚠ 无法连接数据库")
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
            return new_event_id
        else:
            print("⚠ 数据插入失败")
            return None
    
    finally:
        cursor.close()
        conn.close()


def update_google_sheets(selected_row, unfilled_range_name, filled_range_name):
    """更新Google Sheets"""
    print("步骤 7: 更新Google Sheets...")
    
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


def generate_and_send_wechat_message(selected_row, new_event_id, operator, group_members):
    """生成并发送微信群消息"""
    print("步骤 8: 生成微信消息...")
    
    # 生成缩写
    abbreviation = generate_abbreviation(selected_row.iloc[0])
    
    if not abbreviation:
        print("⚠ 无法生成职位缩写")
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
    
    send_wechat_notification(receiver_email, recipient_name, text_output, direction_content, current_date_china)
    print(f"✓ 已发送微信消息通知到 {recipient_name}\n")
    
    return text_output, abbreviation


def add_to_wechat_official_account(selected_row, abbreviation):
    """添加到微信公众号文档"""
    print("步骤 9: 添加到微信公众号文档...")
    
    # 生成微信公众号格式
    wechat_template_output = convert_to_wechat_format(selected_row.iloc[0], abbreviation)
    
    # 计算周范围
    week_start, week_end = calculate_week_range()
    date_subtitle = f"Week: {week_start} to {week_end}"
    
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
    
    print("微信公众号内容:")
    print("-" * 60)
    print(wechat_template_output)
    print("-" * 60)
    print()


def main():
    """主函数"""
    print_banner()
    
    try:
        # 读取组员信息
        group_members = read_group_members(GROUP_MEMBERS_FILE)
        operator = get_operator_name()
        
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
            send_reminder_emails(group_members)
            print("程序结束")
            return
        
        # 步骤5: 验证数据
        if not validate_selected_row(selected_row, group_members, unfilled_data):
            print("程序结束")
            return
        
        # 步骤6: 插入到数据库
        new_event_id = process_and_insert_to_database(selected_row)
        
        if new_event_id is None:
            print("程序结束")
            return
        
        # 步骤7: 更新Google Sheets
        update_google_sheets(selected_row, unfilled_range_name, filled_range_name)
        
        # 步骤8: 生成微信群消息
        result = generate_and_send_wechat_message(selected_row, new_event_id, operator, group_members)
        
        if result is None:
            print("程序结束")
            return
        
        text_output, abbreviation = result
        
        # 步骤9: 添加到微信公众号
        add_to_wechat_official_account(selected_row, abbreviation)
        
        print("=" * 60)
        print("所有步骤完成！".center(60))
        print("=" * 60)
    
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
        sys.exit(0)
    
    except Exception as e:
        print(f"\n⚠ 发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

