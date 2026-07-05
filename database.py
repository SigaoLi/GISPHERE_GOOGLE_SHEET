"""
数据库模块 - 处理MySQL数据库操作
"""
import configparser
import os
import time
import mysql.connector
from mysql.connector import Error
from config import SQL_CREDENTIALS_FILE


def get_database_connection(timeout=15):
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
    
    host = config['MySQL']['host']
    mysql_config = {
        'host': host,
        'port': config['MySQL'].getint('port', 3306),
        'user': config['MySQL']['user'],
        'password': config['MySQL']['password'],
        'database': config['MySQL']['database'],
        'ssl_disabled': True,
        'connection_timeout': timeout,
    }

    # 确保 MySQL 直连，不受系统/全局代理环境变量影响
    saved_proxy = {k: os.environ.get(k) for k in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy', 'ALL_PROXY', 'all_proxy')}
    for key in saved_proxy:
        os.environ.pop(key, None)
    os.environ['NO_PROXY'] = host
    os.environ['no_proxy'] = host

    last_error = None
    try:
        for attempt in range(1, 4):
            try:
                conn = mysql.connector.connect(**mysql_config)
                cursor = conn.cursor()
                return conn, cursor
            except mysql.connector.Error as err:
                last_error = err
                if attempt < 3:
                    print(f"   数据库连接失败，2 秒后重试（{attempt}/3）...")
                    time.sleep(2)
    finally:
        for key, value in saved_proxy.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        os.environ.pop('NO_PROXY', None)
        os.environ.pop('no_proxy', None)

    print(f"Error connecting to database: {last_error}")
    return None, None


def clean_university_names(cursor, conn):
    """清除University_Name_EN列中末尾的多余空格"""
    # 库名统一由 sql_credentials.txt 的 database 决定，不在 SQL 里写死
    cursor.execute("UPDATE new_Universities SET University_Name_EN = RTRIM(University_Name_EN)")
    conn.commit()


def get_gisource_data(cursor):
    """从数据库中获取GISource表的数据"""
    cursor.execute("SELECT University_EN, University_CN, Country_CN FROM GISource")
    return cursor.fetchall()


def check_universities_exist(cursor, university_list):
    """检查哪些大学已存在于数据库中"""
    if not university_list:
        return set()
    
    query = "SELECT University_Name_EN FROM new_Universities WHERE University_Name_EN IN (%s)"
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

