"""
Google API 网络配置 - 代理与 HTTP 客户端
"""
import http.client
import os
import socket
import ssl
import time
from urllib.parse import urlparse

import httplib2
import requests
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_httplib2 import AuthorizedHttp

from config import GOOGLE_API_PROXY

_proxy_notice_printed = False


def setup_google_proxy_env():
    """提示 Google API 代理（代理在各自请求中显式配置，不污染全局环境变量）"""
    global _proxy_notice_printed
    if GOOGLE_API_PROXY and not _proxy_notice_printed:
        print(f"✓ 使用代理访问 Google API: {GOOGLE_API_PROXY}")
        _proxy_notice_printed = True


def refresh_credentials(creds):
    """刷新 OAuth token（显式走代理）"""
    setup_google_proxy_env()
    session = requests.Session()
    if GOOGLE_API_PROXY:
        session.proxies = {'http': GOOGLE_API_PROXY, 'https': GOOGLE_API_PROXY}
    creds.refresh(Request(session=session))


def build_authorized_http(creds, timeout=60):
    """构建带代理的 AuthorizedHttp 客户端"""
    setup_google_proxy_env()
    if GOOGLE_API_PROXY:
        parsed = urlparse(GOOGLE_API_PROXY)
        host = parsed.hostname or '127.0.0.1'
        port = parsed.port or 7890
        proxy_info = httplib2.ProxyInfo(
            httplib2.socks.PROXY_TYPE_HTTP, host, port
        )
        http = httplib2.Http(proxy_info=proxy_info, timeout=timeout)
    else:
        http = httplib2.Http(timeout=timeout)
    return AuthorizedHttp(creds, http=http)


# service 进程内缓存：build() 每次都要解析 discovery 文档（约百毫秒），一次运行调用 15+ 次
# 缓存键为 (api, version)，同时校验 creds 身份，凭据更换时自动重建
_service_cache = {}


def build_google_service(api_name, api_version, creds):
    """构建 Google API 服务（自动走代理；进程内缓存复用）

    execute_with_retry 在瞬时网络错误后会清空缓存，
    保持"重试前重建连接"的原有语义。
    """
    key = (api_name, api_version)
    cached = _service_cache.get(key)
    if cached is not None and cached[1] is creds:
        return cached[0]
    service = build(api_name, api_version, http=build_authorized_http(creds))
    _service_cache[key] = (service, creds)
    return service


def invalidate_service_cache():
    """清空 service 缓存，使下一次构建使用全新的 HTTP 连接"""
    _service_cache.clear()


# 可重试的瞬时网络异常（连接被重置/超时/SSL 中断/不完整读取等）
_RETRYABLE_EXC = (
    ssl.SSLError,            # 含 SSLEOFError、UNEXPECTED_EOF 等
    socket.timeout,
    TimeoutError,
    ConnectionError,        # 含 ConnectionReset/Aborted/BrokenPipe
    http.client.HTTPException,
)
# 服务器侧可重试的 HTTP 状态码（限流与 5xx）
_RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}


def _is_retryable(exc):
    """判断异常是否属于值得重试的瞬时错误"""
    if isinstance(exc, HttpError):
        status = getattr(exc, 'status_code', None)
        if status is None:
            status = getattr(getattr(exc, 'resp', None), 'status', None)
        try:
            return int(status) in _RETRYABLE_HTTP_STATUS
        except (TypeError, ValueError):
            return False
    return isinstance(exc, _RETRYABLE_EXC)


def execute_with_retry(build_request, retries=3, delay=2):
    """执行 Google API 请求，遇到瞬时网络错误时重建连接并重试。

    仅对瞬时错误（SSL 中断、连接重置、超时、429/5xx 等）重试；
    其余错误（如 4xx 鉴权/参数错误）立即抛出，不浪费重试。
    """
    last_error = None
    for attempt in range(1, retries + 1):
        try:
            return build_request().execute()
        except Exception as exc:
            if not _is_retryable(exc):
                raise
            last_error = exc
            if attempt < retries:
                invalidate_service_cache()  # 连接可能已坏，重试时重建
                print(f"   网络请求失败，{delay} 秒后重试（{attempt}/{retries}）...")
                time.sleep(delay)
    raise last_error
