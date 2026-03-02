"""
邮件发送模块 - 处理SMTP邮件发送
"""
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from config import EMAIL_CREDENTIALS_FILE, SMTP_SERVER, SMTP_PORT

# 重试配置
MAX_RETRY_ATTEMPTS = 2  # 最大尝试次数（包括首次）
RETRY_DELAY_SECONDS = 3  # 重试前等待秒数


def read_email_credentials():
    """从文件读取邮箱凭据"""
    # 尝试多种编码方式
    encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']
    
    for encoding in encodings:
        try:
            with open(EMAIL_CREDENTIALS_FILE, "r", encoding=encoding) as file:
                username = file.readline().strip()
                app_password = file.readline().strip()
            return username, app_password
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    # 如果所有编码都失败，抛出错误
    raise UnicodeDecodeError(
        'utf-8', b'', 0, 1,
        f'无法读取文件 {EMAIL_CREDENTIALS_FILE}，请确保文件使用 UTF-8, GBK 或 GB2312 编码'
    )


def send_email(receiver_email, receiver_name, subject, custom_body):
    """
    发送邮件（带自动重试）
    Args:
        receiver_email: 收件人邮箱
        receiver_name: 收件人姓名
        subject: 邮件主题
        custom_body: 邮件正文
    Returns:
        bool: 邮件是否发送成功
    """
    last_error = None
    
    for attempt in range(1, MAX_RETRY_ATTEMPTS + 1):
        try:
            # 读取SMTP凭据
            username, app_password = read_email_credentials()
            
            # 创建邮件
            message = MIMEMultipart()
            message["From"] = username
            message["To"] = receiver_email
            message["Subject"] = subject
            message.attach(MIMEText(custom_body, "plain", "utf-8"))
            
            # 发送邮件
            server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30)
            server.starttls()
            server.login(username, app_password)
            server.send_message(message)
            server.quit()
            
            if attempt > 1:
                print(f"Email sent successfully to {receiver_name} ({receiver_email}) [重试第{attempt-1}次成功]")
            else:
                print(f"Email sent successfully to {receiver_name} ({receiver_email})")
            return True
        
        except Exception as e:
            last_error = e
            if attempt < MAX_RETRY_ATTEMPTS:
                print(f"Failed to send email to {receiver_name} (attempt {attempt}/{MAX_RETRY_ATTEMPTS}): {e}")
                print(f"Retrying in {RETRY_DELAY_SECONDS} seconds...")
                time.sleep(RETRY_DELAY_SECONDS)
            else:
                print(f"Failed to send email to {receiver_name} after {MAX_RETRY_ATTEMPTS} attempts: {last_error}")
    
    return False


def send_reminder_emails(group_members):
    """向所有组员发送提醒邮件"""
    reminder_subject = "GISource提醒：添加内容"
    reminder_body = (
        "亲爱的 GISource 团队成员，\n\n"
        "现有资讯消息已全部发送，请您尽快添加/完善内容"
        "（https://docs.google.com/spreadsheets/d/1LcfxcTCuj9ZJXXMxyFQwt-xnbAviNP8j9oDr6OG5-Go/edit#gid=0）。\n\n"
        "如果您已退出相关工作，请回复本邮件告知我们。"
    )
    
    for member_name, member_email in group_members.items():
        send_email(member_email, member_name, reminder_subject, reminder_body)


def send_error_notification(receiver_email, receiver_name, source_content, university_content, 
                           direction_content, current_date_china):
    """发送错误通知邮件
    
    Returns:
        bool: 邮件发送是否成功
    """
    subject = f"GISource信息错误提醒 - {current_date_china} - {direction_content}"
    body = (
        f"{receiver_name}同学您好，\n\n"
        f"您审核的 \"{university_content}-{direction_content}\" 消息有误，请及时更正。\n\n"
        f"消息链接：{source_content}"
    )
    
    return send_email(receiver_email, receiver_name, subject, body)


def send_wechat_notification(receiver_email, receiver_name, text_output, 
                            direction_content, current_date_china):
    """发送微信群信息通知邮件
    
    Returns:
        bool: 邮件发送是否成功
    """
    subject = f"微信群信息发送通知 - {current_date_china} - {direction_content}"
    email_body = (
        f"{receiver_name}同学您好，\n\n"
        f"请在确认信息无误后，发送以下信息至微信群。\n\n\n\n"
        f"{text_output}"
    )
    
    return send_email(receiver_email, receiver_name, subject, email_body)

