"""
邮件发送模块 - 处理SMTP邮件发送
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from config import EMAIL_CREDENTIALS_FILE, SMTP_SERVER, SMTP_PORT


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
    发送邮件
    Args:
        receiver_email: 收件人邮箱
        receiver_name: 收件人姓名
        subject: 邮件主题
        custom_body: 邮件正文
    """
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
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(username, app_password)
        server.send_message(message)
        server.quit()
        
        print(f"Email sent successfully to {receiver_name} ({receiver_email})")
        return True
    
    except Exception as e:
        print(f"Failed to send email to {receiver_name}: {e}")
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
    """发送错误通知邮件"""
    subject = f"GISource信息错误提醒 - {current_date_china} - {direction_content}"
    body = (
        f"{receiver_name}同学您好，\n\n"
        f"您填写的 \"{university_content}-{direction_content}\" 消息有误，请及时更正。\n\n"
        f"消息链接：{source_content}"
    )
    
    send_email(receiver_email, receiver_name, subject, body)


def send_wechat_notification(receiver_email, receiver_name, text_output, 
                            direction_content, current_date_china):
    """发送微信群信息通知邮件"""
    subject = f"微信群信息发送通知 - {current_date_china} - {direction_content}"
    email_body = (
        f"{receiver_name}同学您好，\n\n"
        f"请在确认信息无误后，发送以下信息至微信群。\n\n\n\n"
        f"{text_output}"
    )
    
    send_email(receiver_email, receiver_name, subject, email_body)

