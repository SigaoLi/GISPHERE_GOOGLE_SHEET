"""
日志记录模块 - 记录程序运行历史和LLM对话记录
所有日志输出到单一的 .txt 文件中
"""
import os
import sys
import json
from datetime import datetime
from config import BASE_DIR, CHINA_TZ

# 日志文件夹路径
LLM_LOGS_DIR = os.path.join(BASE_DIR, 'llm_logs')
LOGS_DIR = os.path.join(BASE_DIR, 'logs')

# 确保日志文件夹存在
os.makedirs(LLM_LOGS_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)

# 会话日志缓冲区（存储结构化日志条目）
_session_log_buffer = []


class TeeOutput:
    """
    同时输出到控制台和文件的类
    用于捕获所有print输出并保存到日志文件
    """
    def __init__(self, file_path):
        self.file = open(file_path, 'w', encoding='utf-8')
        self.stdout = sys.stdout
        self.stderr = sys.stderr
        
    def write(self, text):
        # 同时写入控制台和文件
        self.stdout.write(text)
        self.file.write(text)
        self.file.flush()  # 确保立即写入文件
        
    def flush(self):
        self.stdout.flush()
        self.file.flush()
        
    def close(self):
        if self.file and not self.file.closed:
            self.file.close()
            
    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()


def setup_print_logging():
    """
    设置print输出同时记录到日志文件
    返回TeeOutput对象，需要在程序结束时调用close()
    日志文件包含print输出和结构化日志
    """
    # 获取会话ID（与log_program_run使用相同的会话ID）
    session_id = getattr(log_program_run, '_session_id', None)
    if session_id is None:
        session_id = datetime.now(CHINA_TZ).strftime('%Y%m%d_%H%M%S')
        log_program_run._session_id = session_id
    
    # 创建日志文件（单一的txt文件，包含所有输出和结构化日志）
    log_filename = f"run_{session_id}.txt"
    log_filepath = os.path.join(LOGS_DIR, log_filename)
    
    # 创建TeeOutput对象
    tee = TeeOutput(log_filepath)
    
    # 重定向stdout和stderr
    sys.stdout = tee
    sys.stderr = tee
    
    return tee


def restore_print_logging(tee):
    """
    恢复原始的stdout和stderr，并将结构化日志追加到txt文件
    """
    if tee:
        # 先恢复stdout，这样后续的写入不会输出到控制台
        original_stdout = tee.stdout
        original_stderr = tee.stderr
        
        if sys.stdout == tee:
            sys.stdout = original_stdout
        if sys.stderr == tee:
            sys.stderr = original_stderr
        
        # 将结构化日志摘要写入到同一个txt文件
        log_summary = format_log_summary()
        if log_summary and tee.file and not tee.file.closed:
            tee.file.write(log_summary)
            tee.file.flush()
        
        tee.close()


def get_timestamp():
    """获取当前时间戳（中国时区）"""
    return datetime.now(CHINA_TZ).strftime('%Y-%m-%d %H:%M:%S')


def get_log_filename(prefix='run', extension='log'):
    """生成日志文件名"""
    timestamp = datetime.now(CHINA_TZ).strftime('%Y%m%d_%H%M%S')
    return f"{prefix}_{timestamp}.{extension}"


def log_llm_conversation(system_prompt, user_prompt, response, model=None, metadata=None):
    """
    记录LLM对话到llm_logs文件夹（TXT格式）
    
    Args:
        system_prompt: 系统提示词
        user_prompt: 用户提示词
        response: LLM响应内容
        model: 使用的模型名称（可选）
        metadata: 额外的元数据（可选，字典格式）
    """
    try:
        timestamp = get_timestamp()
        
        # 构建TXT格式的日志内容
        log_lines = [
            "=" * 60,
            f"时间: {timestamp}",
            f"模型: {model or '未指定'}",
            "=" * 60,
            "",
            "【系统提示词】",
            "-" * 40,
            system_prompt or "(无)",
            "",
            "【用户提示词】",
            "-" * 40,
            user_prompt or "(无)",
            "",
            "【LLM响应】",
            "-" * 40,
            response or "(无)",
            "",
        ]
        
        # 添加元数据（如果有）
        if metadata:
            log_lines.append("【元数据】")
            log_lines.append("-" * 40)
            for key, value in metadata.items():
                log_lines.append(f"{key}: {value}")
            log_lines.append("")
        
        log_lines.append("=" * 60)
        
        # 生成日志文件名
        filename = get_log_filename('llm', 'txt')
        filepath = os.path.join(LLM_LOGS_DIR, filename)
        
        # 写入TXT文件
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write('\n'.join(log_lines))
        
        print(f"✓ LLM对话已记录到: {filename}")
        return filepath
        
    except Exception as e:
        print(f"⚠ 记录LLM对话失败: {e}")


def log_program_run(step, message, status='info', data=None):
    """
    记录程序运行日志到内存缓冲区
    
    Args:
        step: 步骤名称或编号
        message: 日志消息
        status: 状态（info, success, warning, error）
        data: 额外的数据（可选，字典格式）
    """
    global _session_log_buffer
    
    try:
        timestamp = get_timestamp()
        log_entry = {
            'timestamp': timestamp,
            'step': step,
            'message': message,
            'status': status,
            'data': data or {}
        }
        
        # 获取或创建当前运行会话的ID
        session_id = getattr(log_program_run, '_session_id', None)
        if session_id is None:
            session_id = datetime.now(CHINA_TZ).strftime('%Y%m%d_%H%M%S')
            log_program_run._session_id = session_id
        
        # 将日志条目添加到内存缓冲区
        _session_log_buffer.append(log_entry)
        
        return True
        
    except Exception as e:
        print(f"⚠ 记录程序运行日志失败: {e}")


def reset_session():
    """重置会话ID和日志缓冲区（用于新的一次程序运行）"""
    global _session_log_buffer
    _session_log_buffer = []
    if hasattr(log_program_run, '_session_id'):
        delattr(log_program_run, '_session_id')


def log_program_start():
    """记录程序开始运行"""
    reset_session()
    log_program_run('START', '程序开始运行', 'info', {
        'timezone': 'Asia/Shanghai'
    })
    # 设置print输出日志
    return setup_print_logging()


def format_log_summary():
    """
    将内存中的日志条目格式化为人类可读的文本
    
    Returns:
        str: 格式化后的日志摘要文本
    """
    global _session_log_buffer
    
    if not _session_log_buffer:
        return ""
    
    # 状态符号映射
    status_symbols = {
        'info': 'ℹ',
        'success': '✓',
        'warning': '⚠',
        'error': '✗'
    }
    
    lines = [
        "",
        "",
        "=" * 60,
        "                     结构化运行日志                      ",
        "=" * 60,
        ""
    ]
    
    for entry in _session_log_buffer:
        symbol = status_symbols.get(entry['status'], '•')
        step = entry['step']
        timestamp = entry['timestamp']
        message = entry['message']
        
        # 格式化步骤标识
        if step in ['START', 'END', 'INIT', 'PRE', 'MAIN', 'ERROR']:
            step_str = f"[{step}]"
        else:
            step_str = f"[步骤 {step}]"
        
        # 主日志行
        lines.append(f"{symbol} {timestamp} {step_str} {message}")
        
        # 如果有额外数据，格式化显示
        if entry['data']:
            for key, value in entry['data'].items():
                if key == 'traceback':
                    # 特殊处理错误追踪
                    lines.append(f"     └─ 错误追踪:")
                    for trace_line in str(value).split('\n'):
                        if trace_line.strip():
                            lines.append(f"        {trace_line}")
                else:
                    lines.append(f"     └─ {key}: {value}")
    
    lines.append("")
    lines.append("=" * 60)
    
    return '\n'.join(lines)


def log_program_end(success=True, error_message=None):
    """记录程序结束"""
    status = 'success' if success else 'error'
    data = {}
    if error_message:
        data['error'] = error_message
    log_program_run('END', '程序运行结束', status, data)
