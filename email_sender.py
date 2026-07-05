"""
邮件发送模块 - 处理SMTP邮件发送
Gmail 在检测到代理时优先经 Gmail API（与 Google Sheets 相同代理通路）发送
"""
import base64
import configparser
import os
import smtplib
import time
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from config import (
    EMAIL_CREDENTIALS_FILE,
    FAILED_EMAIL_LOG_FILE,
    GOOGLE_API_PROXY,
    GMAIL_SMTP_SERVER,
    GMAIL_SMTP_PORT,
    GMAIL_SMTP_USE_SSL,
    QQMAIL_SMTP_SERVER,
    QQMAIL_SMTP_PORT,
    QQMAIL_SMTP_USE_SSL,
)
from smtp_proxy import create_smtp_client

_gmail_api_notice_printed = False

# 重试配置
MAX_RETRY_ATTEMPTS = 2  # 最大尝试次数（包括首次）
RETRY_DELAY_SECONDS = 3  # 重试前等待秒数


def _append_failed_email_record(receiver_email, receiver_name, subject, custom_body, error_message):
    """将发送失败的邮件内容追加记录到日志文件（条目间空一行）"""
    try:
        os.makedirs(os.path.dirname(FAILED_EMAIL_LOG_FILE), exist_ok=True)
        record = (
            f"[FAILED_AT] {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"[RECEIVER_NAME] {receiver_name}\n"
            f"[RECEIVER_EMAIL] {receiver_email}\n"
            f"[SUBJECT] {subject}\n"
            f"[BODY]\n{custom_body}\n"
            f"[ERROR] {error_message}\n"
        )

        with open(FAILED_EMAIL_LOG_FILE, "a", encoding="utf-8") as file:
            # 文件已有内容时，先空一行再追加新的失败记录
            if os.path.exists(FAILED_EMAIL_LOG_FILE) and os.path.getsize(FAILED_EMAIL_LOG_FILE) > 0:
                file.write("\n")
            file.write(record)

        print(f"[FailureLog] 已记录发送失败邮件到: {FAILED_EMAIL_LOG_FILE}")
    except Exception as log_error:
        print(f"[FailureLog] 记录失败邮件时出错: {log_error}")


def read_email_credentials():
    """从文件读取 Gmail 和 QQmail 凭据（INI格式）"""
    # 尝试多种编码方式
    encodings = ['utf-8', 'gbk', 'gb2312', 'utf-8-sig']

    for encoding in encodings:
        try:
            parser = configparser.ConfigParser()
            with open(EMAIL_CREDENTIALS_FILE, "r", encoding=encoding) as file:
                parser.read_file(file)

            required_sections = ["Gmail", "QQmail"]
            for section in required_sections:
                if not parser.has_section(section):
                    raise ValueError(f"邮箱凭据缺少 [{section}] 段")
                if not parser.has_option(section, "email") or not parser.has_option(section, "password"):
                    raise ValueError(f"[{section}] 段缺少 email 或 password 字段")

            return {
                "Gmail": {
                    "email": parser.get("Gmail", "email").strip(),
                    "password": parser.get("Gmail", "password").strip(),
                },
                "QQmail": {
                    "email": parser.get("QQmail", "email").strip(),
                    "password": parser.get("QQmail", "password").strip(),
                }
            }
        except (UnicodeDecodeError, UnicodeError):
            continue
        except configparser.Error as e:
            raise ValueError(f"邮箱凭据格式错误: {e}") from e

    # 如果所有编码读取都失败，抛出错误
    raise UnicodeDecodeError(
        'utf-8', b'', 0, 1,
        f'无法读取文件 {EMAIL_CREDENTIALS_FILE}，请确保文件使用 UTF-8, GBK 或 GB2312 编码'
    )


def _build_message(sender_email, receiver_email, subject, custom_body):
    """构建邮件消息对象"""
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.attach(MIMEText(custom_body, "plain", "utf-8"))
    return message


def _send_gmail_via_api(sender_email, receiver_email, receiver_name, subject, custom_body):
    """经 Gmail API 发送（走与 Google Sheets 相同的代理）"""
    global _gmail_api_notice_printed
    from google_sheets import authorize_credentials
    from google_http import build_google_service, setup_google_proxy_env

    setup_google_proxy_env()
    if not _gmail_api_notice_printed:
        print(f"✓ Gmail 经代理 API 发送（{GOOGLE_API_PROXY}）")
        _gmail_api_notice_printed = True

    creds = authorize_credentials()
    service = build_google_service('gmail', 'v1', creds)

    message = _build_message(sender_email, receiver_email, subject, custom_body)
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(userId='me', body={'raw': raw}).execute()
    print(f"[Gmail-API] Email sent successfully to {receiver_name} ({receiver_email})")
    return True, None


def _send_via_provider(
    provider_name,
    smtp_server,
    smtp_port,
    use_ssl,
    username,
    app_password,
    receiver_email,
    receiver_name,
    subject,
    custom_body,
    use_proxy=False,
):
    """通过指定 SMTP 通道发送邮件（带重试）"""
    last_error = None

    for attempt in range(1, MAX_RETRY_ATTEMPTS + 1):
        server = None
        try:
            message = _build_message(username, receiver_email, subject, custom_body)

            if use_proxy:
                server = create_smtp_client(smtp_server, smtp_port, use_ssl, timeout=30)
            elif use_ssl:
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
            else:
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
                server.starttls()
            server.login(username, app_password)
            server.send_message(message)
            server.quit()

            if attempt > 1:
                print(
                    f"[{provider_name}] Email sent successfully to {receiver_name} "
                    f"({receiver_email}) [重试第{attempt - 1}次成功]"
                )
            else:
                print(f"[{provider_name}] Email sent successfully to {receiver_name} ({receiver_email})")
            return True, None

        except Exception as e:
            last_error = e
            if attempt < MAX_RETRY_ATTEMPTS:
                print(
                    f"[{provider_name}] Failed to send email to {receiver_name} "
                    f"(attempt {attempt}/{MAX_RETRY_ATTEMPTS}): {e}"
                )
                print(f"[{provider_name}] Retrying in {RETRY_DELAY_SECONDS} seconds...")
                time.sleep(RETRY_DELAY_SECONDS)
            else:
                print(
                    f"[{provider_name}] Failed to send email to {receiver_name} after "
                    f"{MAX_RETRY_ATTEMPTS} attempts: {last_error}"
                )
        finally:
            if server is not None:
                try:
                    server.quit()
                except Exception:
                    pass

    return False, last_error


def send_email(receiver_email, receiver_name, subject, custom_body):
    """
    发送邮件（Gmail 主通道失败后自动切换 QQmail）
    Args:
        receiver_email: 收件人邮箱
        receiver_name: 收件人姓名
        subject: 邮件主题
        custom_body: 邮件正文
    Returns:
        bool: 邮件是否发送成功
    """
    try:
        credentials = read_email_credentials()
        gmail_email = credentials["Gmail"]["email"]
        gmail_password = credentials["Gmail"]["password"]

        gmail_ok, gmail_error = False, None

        if GOOGLE_API_PROXY:
            for attempt in range(1, MAX_RETRY_ATTEMPTS + 1):
                try:
                    gmail_ok, gmail_error = _send_gmail_via_api(
                        gmail_email, receiver_email, receiver_name, subject, custom_body
                    )
                    break
                except Exception as e:
                    gmail_error = e
                    if attempt < MAX_RETRY_ATTEMPTS:
                        print(
                            f"[Gmail-API] Failed (attempt {attempt}/{MAX_RETRY_ATTEMPTS}): {e}"
                        )
                        print(f"[Gmail-API] Retrying in {RETRY_DELAY_SECONDS} seconds...")
                        time.sleep(RETRY_DELAY_SECONDS)
                    else:
                        print(f"[Gmail-API] Failed after {MAX_RETRY_ATTEMPTS} attempts: {e}")
                        print("[Gmail-API] 尝试经代理 SMTP 备用发送...")

        if not gmail_ok:
            gmail_ok, gmail_error = _send_via_provider(
                provider_name="Gmail",
                smtp_server=GMAIL_SMTP_SERVER,
                smtp_port=GMAIL_SMTP_PORT,
                use_ssl=GMAIL_SMTP_USE_SSL,
                username=gmail_email,
                app_password=gmail_password,
                receiver_email=receiver_email,
                receiver_name=receiver_name,
                subject=subject,
                custom_body=custom_body,
                use_proxy=bool(GOOGLE_API_PROXY),
            )
        if gmail_ok:
            return True

        print(f"[Fallback] Gmail发送失败，开始切换QQmail备用通道: {gmail_error}")

        qq_ok, qq_error = _send_via_provider(
            provider_name="QQmail",
            smtp_server=QQMAIL_SMTP_SERVER,
            smtp_port=QQMAIL_SMTP_PORT,
            use_ssl=QQMAIL_SMTP_USE_SSL,
            username=credentials["QQmail"]["email"],
            app_password=credentials["QQmail"]["password"],
            receiver_email=receiver_email,
            receiver_name=receiver_name,
            subject=subject,
            custom_body=custom_body,
        )
        if qq_ok:
            print("[Fallback] QQmail备用通道发送成功")
            return True

        final_error = f"GmailError={gmail_error}; QQmailError={qq_error}"
        print(f"[Fallback] QQmail备用通道也发送失败: {qq_error}")
        _append_failed_email_record(receiver_email, receiver_name, subject, custom_body, final_error)
        return False
    except Exception as e:
        _append_failed_email_record(receiver_email, receiver_name, subject, custom_body, str(e))
        print(f"[SendEmail] 发送过程中出现异常，已记录失败信息: {e}")
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

