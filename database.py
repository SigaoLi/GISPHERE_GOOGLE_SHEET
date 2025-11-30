"""
数据库模块 - 处理MySQL数据库操作
"""
import configparser
import threading
import mysql.connector
from mysql.connector import Error
from config import SQL_CREDENTIALS_FILE


def connect_to_database(config, result):
    """连接到MySQL数据库的线程函数"""
    try:
        conn = mysql.connector.connect(**config)
        cursor = conn.cursor()
        result['connection'] = conn
        result['cursor'] = cursor
    except mysql.connector.Error as err:
        result['error'] = err


def get_database_connection(timeout=60):
    """
    获取数据库连接
    Args:
        timeout: 连接超时时间（秒）
    Returns:
        tuple: (connection, cursor) 或 (None, None)
    """
    # 从文件中读取凭据，尝试多种编码
    config = configparser.ConfigParser()
    encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']
    
    config_loaded = False
    for encoding in encodings:
        try:
            config.read(SQL_CREDENTIALS_FILE, encoding=encoding)
            if config.sections():  # 确保配置已加载
                config_loaded = True
                break
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    if not config_loaded:
        print(f"⚠ 无法读取配置文件 {SQL_CREDENTIALS_FILE}，请检查文件编码")
        return None, None
    
    # MySQL配置
    mysql_config = {
        'host': config['MySQL']['host'],
        'port': config['MySQL'].getint('port', 3306),
        'user': config['MySQL']['user'],
        'password': config['MySQL']['password'],
        'database': config['MySQL']['database'],
        'ssl_disabled': True  # 禁用SSL以避免版本不匹配错误
    }
    
    # 使用字典存储连接结果
    result = {}
    
    # 创建连接线程
    db_thread = threading.Thread(target=connect_to_database, args=(mysql_config, result))
    db_thread.start()
    db_thread.join(timeout=timeout)
    
    if 'error' in result:
        print(f"Error connecting to database: {result['error']}")
        return None, None
    elif 'connection' in result:
        return result['connection'], result['cursor']
    else:
        print("Connection to the database timed out.")
        return None, None


def clean_university_names(cursor, conn):
    """清除University_Name_EN列中末尾的多余空格"""
    cursor.execute("UPDATE TEST.new_Universities SET University_Name_EN = RTRIM(University_Name_EN)")
    conn.commit()


def get_gisource_data(cursor):
    """从数据库中获取GISource表的数据"""
    cursor.execute("SELECT University_EN, University_CN, Country_CN FROM TEST.GISource")
    return cursor.fetchall()


def check_universities_exist(cursor, university_list):
    """检查哪些大学已存在于数据库中"""
    if not university_list:
        return set()
    
    query = "SELECT University_Name_EN FROM TEST.new_Universities WHERE University_Name_EN IN (%s)"
    format_strings = ','.join(['%s'] * len(university_list))
    cursor.execute(query % format_strings, tuple(university_list))
    existing_universities = cursor.fetchall()
    
    return set([row[0] for row in existing_universities])


def get_max_event_id(cursor):
    """获取最近日期的最后一个Event_ID"""
    cursor.execute("""
        SELECT MAX(Event_ID) FROM GISource
        WHERE Date = (SELECT MAX(Date) FROM GISource);
    """)
    result = cursor.fetchone()
    return result[0] if result and result[0] is not None else 0


def insert_event_to_database(cursor, conn, sql_table, table_name='GISource'):
    """
    将事件数据插入到数据库
    Args:
        cursor: 数据库游标
        conn: 数据库连接
        sql_table: 包含事件数据的DataFrame
        table_name: 目标表名
    Returns:
        bool: 插入是否成功
    """
    try:
        for i, row in sql_table.iterrows():
            sql_query = f"INSERT INTO {table_name} ({', '.join(row.index)}) VALUES ({', '.join(['%s']*len(row))})"
            cursor.execute(sql_query, tuple(row))
        
        conn.commit()
        print("Data inserted successfully.")
        return True
    except mysql.connector.Error as err:
        print(f"Error: {err}")
        conn.rollback()
        return False

