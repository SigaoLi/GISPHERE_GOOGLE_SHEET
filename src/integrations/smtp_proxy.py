"""
SMTP 代理连接 - 通过本地代理连接 Gmail SMTP（备用方案）
"""
import ssl
import smtplib
from urllib.parse import urlparse

import socks

from ..core.config import GOOGLE_API_PROXY

_gmail_proxy_notice_printed = False


def _print_gmail_proxy_notice():
    global _gmail_proxy_notice_printed
    if GOOGLE_API_PROXY and not _gmail_proxy_notice_printed:
        print(f"✓ Gmail SMTP 通过代理连接: {GOOGLE_API_PROXY}")
        _gmail_proxy_notice_printed = True


def _connect_via_proxy(smtp_server, smtp_port, timeout=30):
    """经 SOCKS5/HTTP 代理建立到 SMTP 服务器的 TCP 连接"""
    parsed = urlparse(GOOGLE_API_PROXY)
    proxy_host = parsed.hostname or '127.0.0.1'
    proxy_port = parsed.port or 7890

    last_error = None
    for proxy_type, label in (
        (socks.SOCKS5, 'SOCKS5'),
        (socks.HTTP, 'HTTP'),
    ):
        sock = socks.socksocket()
        try:
            sock.set_proxy(proxy_type, proxy_host, proxy_port, rdns=True)
            sock.settimeout(timeout)
            sock.connect((smtp_server, smtp_port))
            return sock
        except Exception as e:
            last_error = e
            try:
                sock.close()
            except Exception:
                pass

    raise ConnectionError(f"无法经代理连接 {smtp_server}:{smtp_port} ({last_error})")


def create_smtp_client(smtp_server, smtp_port, use_ssl, timeout=30):
    """
    创建 SMTP 客户端。
    当配置了 GOOGLE_API_PROXY 时，强制经代理连接（仅用于 Gmail）。
    """
    if not GOOGLE_API_PROXY:
        if use_ssl:
            return smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=timeout)
        client = smtplib.SMTP(smtp_server, smtp_port, timeout=timeout)
        client.starttls(context=ssl.create_default_context())
        return client

    _print_gmail_proxy_notice()
    sock = _connect_via_proxy(smtp_server, smtp_port, timeout)

    if use_ssl:
        context = ssl.create_default_context()
        sock = context.wrap_socket(sock, server_hostname=smtp_server)
        client = smtplib.SMTP_SSL()
        client.sock = sock
        client.file = sock.makefile('rb')
        client._host = smtp_server
        client.ehlo()
        return client

    client = smtplib.SMTP()
    client.sock = sock
    client.file = sock.makefile('rb')
    client._host = smtp_server
    client.ehlo()
    client.starttls(context=ssl.create_default_context())
    client.ehlo()
    return client
