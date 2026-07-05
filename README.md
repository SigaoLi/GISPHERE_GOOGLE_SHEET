# GISource 自动化系统

自动化处理GISource学术信息的发布系统，支持Windows、macOS和Linux多平台运行。

**版本**: v2.3.0 | **最后更新**: 2026-07-05

---

## 📋 目录

- [功能特性](#功能特性)
- [系统要求](#系统要求)
- [快速开始](#快速开始)
- [项目结构](#项目结构)
- [详细安装步骤](#详细安装步骤)
- [使用方法](#使用方法)
- [配置说明](#配置说明)
- [核心模块说明](#核心模块说明)
- [工作流程](#工作流程)
- [自定义配置](#自定义配置)
- [故障排除](#故障排除)
- [开发指南](#开发指南)
- [版本历史](#版本历史)

---

## 功能特性

- 📊 **自动数据处理** - 从Google Sheets获取和处理学术信息
- 💾 **数据库同步** - 自动同步数据到MySQL数据库
- 📧 **邮件通知** - 自动发送邮件通知相关人员
- 📱 **微信集成** - 自动生成微信群消息和公众号内容
- 🌐 **国内网络友好** - 自动探测本地代理（Clash/V2Ray 等）并经代理访问 Google；探测后会验证可用性，避免误判
- ✉️ **Gmail 智能发送** - 检测到代理时优先经 Gmail API 发送（与 Sheets 同一代理通路），失败再回退代理 SMTP / QQmail
- 🤖 **LLM 模型回退链** - 内容组织优先 Claude，其次 GPT，最后 Gemini，任一不可用自动切换
- 🔁 **网络重试** - Google API 请求对瞬时错误（超时/连接重置/SSL 中断/429/5xx）自动重试
- 🌍 **跨平台支持** - 完全支持Windows/macOS/Linux
- 🔄 **模块化设计** - 易于维护和扩展
- 🔐 **安全管理** - 凭据文件独立管理，不纳入版本控制；OAuth 令牌改用可移植 JSON（跨 google-auth 版本均可加载）

---

## 系统要求

- Python 3.8 或更高版本（推荐 3.10+）
- 稳定的网络连接；**国内环境需本地代理**（Clash/V2Ray 等）以访问 Google，程序会自动探测常见端口（7897/7890/1087/1080/10809/33210）或读取 `HTTPS_PROXY`/`GOOGLE_API_PROXY` 环境变量；显式设置 `GOOGLE_API_PROXY` 会**跳过端口探测**（启动更快，也适用于非常规代理配置；海外直连环境无需任何设置）
- Google API 访问权限（Google Sheets + Google Docs + Gmail 发送）
- MySQL 数据库访问权限
- Gmail 和 QQmail 邮箱（用于主备发送通知）
- LLM 网关密钥（`keys/llm_key.txt`，用于 `newapi.gisphere.info` 网关，供内容组织调用 Claude/GPT/Gemini）
- `PySocks`（经代理发送 Gmail SMTP 时需要；已在 `requirements.txt` 中）

---

## 快速开始

### 准备清单

在开始之前，请确保您有：
- [ ] Python 3.8+ 已安装
- [ ] Google账号（用于访问Google Sheets和Docs）
- [ ] Gmail邮箱（主发送通道）
- [ ] QQmail邮箱（备用发送通道）
- [ ] MySQL数据库访问权限
- [ ] 网络连接正常

### 5步快速启动

#### 步骤 1: 安装依赖（1分钟）

**Windows:**
```bash
cd D:\申请\实习\GISphere\GISource\海外资讯\Test\Google_Sheet
pip install -r requirements.txt
```

**macOS/Linux:**
```bash
cd /path/to/Google_Sheet
pip3 install -r requirements.txt
```

#### 步骤 2: 配置Google API（2分钟）

1. 访问 [Google Cloud Console](https://console.cloud.google.com/)
2. 创建新项目或选择现有项目
3. 启用以下API：
   - Google Sheets API
   - Google Docs API
   - Gmail API（用于经代理通过 Gmail API 发信）
4. 创建OAuth 2.0凭据（类型：桌面应用）
5. 下载JSON文件，重命名为 `credentials.json`
6. 创建 `keys` 文件夹（如果不存在）
7. 将 `credentials.json` 放在 `keys/` 目录中

#### 步骤 3: 配置邮箱（1分钟）

**获取 Gmail / QQmail 发信凭据：**
1. 访问：https://myaccount.google.com/security
2. 启用"两步验证"
3. 在"两步验证"页面，找到"应用专用密码"
4. 生成新密码（选择"邮件"和"Windows计算机"或其他设备）
5. 复制生成的16位密码（类似：`xxxx xxxx xxxx xxxx`）
6. QQ 邮箱请在邮箱设置中开启 SMTP，并生成授权码（不是登录密码）
7. QQmail 推荐使用 `smtp.qq.com:465`（SSL 连接）

**创建凭据文件：**

Windows:
```bash
mkdir keys
copy keys\email_credentials.txt.example keys\email_credentials.txt
notepad keys\email_credentials.txt
```

macOS/Linux:
```bash
mkdir -p keys
cp keys/email_credentials.txt.example keys/email_credentials.txt
nano keys/email_credentials.txt
```

编辑内容为（INI 分段格式，区分 `Gmail` 与 `QQmail`）：
```
[Gmail]
email = your-gmail@gmail.com
password = xxxx xxxx xxxx xxxx

[QQmail]
email = your-qq@qq.com
password = your-qq-auth-code
```

#### 步骤 4: 配置数据库（1分钟）

Windows:
```bash
copy keys\sql_credentials.txt.example keys\sql_credentials.txt
notepad keys\sql_credentials.txt
```

macOS/Linux:
```bash
cp keys/sql_credentials.txt.example keys/sql_credentials.txt
nano keys/sql_credentials.txt
```

编辑内容：
```ini
[MySQL]
host = your-host-address
port = 3306
user = your-username
password = your-password
database = your-database-name
```

#### 步骤 5: 运行程序（< 1分钟）

**Windows:**
```bash
python run.py
```

**macOS/Linux:**
```bash
python3 run.py
```

### 首次运行说明

首次运行时：

1. **浏览器会自动打开**，要求您授权Google API
2. **选择您的Google账号**
3. **点击"允许"**授权访问
4. 授权成功后，浏览器会显示"The authentication flow has completed."
5. **返回终端**，程序将自动继续运行

授权完成后会生成可移植的 JSON 令牌：
- `keys/token_sheets.json` - Sheets + Gmail 发送令牌（首选；JSON 格式，跨 google-auth 版本均可加载）
- `keys/token.json` - Google Docs 令牌

> 旧版 `keys/token.pickle` 仍兼容读取，但 pickle 跨库版本可能无法加载；新版统一以 JSON 持久化，加载失败会自动重新授权而非崩溃。

以后运行将自动使用这些凭据，无需重新授权。首次运行需授予 `spreadsheets` 与 `gmail.send` 权限。

---

## 项目结构

```
Google_Sheet/
│
├── 🚀 启动入口
│   └── run.py                     # 唯一启动方式：python run.py
│
├── 📦 源码包 src/
│   ├── main.py                    # 主流程编排
│   ├── core/                      # 核心模块
│   │   ├── config.py              # 配置管理（常量、字典、代理探测、LLM 模型链）
│   │   ├── utils.py               # 工具函数
│   │   ├── logger.py              # 日志记录及终端输出捕获
│   │   └── data_processor.py      # 数据处理和格式转换
│   ├── integrations/              # 外部集成
│   │   ├── google_http.py         # Google API 代理/HTTP 客户端 + 请求重试
│   │   ├── google_sheets.py       # Google Sheets API（JSON 令牌 + 重试 + 进程内缓存）
│   │   ├── google_docs.py         # Google Docs API + LLM 内容组织（模型链）
│   │   ├── database.py            # MySQL数据库操作（重试 + 代理隔离）
│   │   ├── email_sender.py        # 邮件发送（Gmail API 优先 / 代理 SMTP / QQmail 回退）
│   │   └── smtp_proxy.py          # 经本地代理连接 Gmail SMTP（备用通道）
│   └── tools/
│       └── check_setup.py         # 环境检查：python -m src.tools.check_setup
│
├── 📋 配置文件（需手动创建）
│   └── keys/                   # 密钥文件夹
│       ├── credentials.json        # Google API凭据
│       ├── email_credentials.txt   # 邮箱凭据
│       ├── group_members.txt       # 组员信息
│       ├── llm_key.txt             # LLM 网关密钥（Claude/GPT/Gemini）
│       └── sql_credentials.txt     # 数据库凭据
│
├── 🔐 认证文件（自动生成）
│   ├── keys/token_sheets.json      # Sheets + Gmail 令牌（JSON，可移植）
│   ├── keys/token.json             # Google Docs 令牌
│   └── keys/token.pickle           # 旧版令牌（兼容读取）
│
└── 📚 其他
    ├── logs/                      # 运行日志归档目录（自动生成）
    ├── llm_logs/                  # LLM对话记录目录（自动生成）
    ├── requirements.txt           # Python依赖
    └── .gitignore                 # Git配置
```

---

## 详细安装步骤

### 1. 检查Python版本

```bash
python --version  # Windows
python3 --version # macOS/Linux
```

确保版本 >= 3.8

### 2. 安装依赖包

```bash
pip install -r requirements.txt
```

需要安装的包：
- pandas, numpy - 数据处理
- google-api-python-client - Google API
- mysql-connector-python - MySQL连接
- pytz, inflect, pypinyin - 工具库

### 3. 配置文件设置

> 以下每个文件在 `keys/` 下都有对应的 `.example` 模板，复制去掉 `.example` 后缀再填写即可。

#### credentials.json（Google API）
从 Google Cloud Console 下载（OAuth 桌面应用客户端）

#### email_credentials.txt
```ini
[Gmail]
email = your-gmail@gmail.com
password = your-gmail-app-password

[QQmail]
email = your-qq@qq.com
password = your-qq-auth-code
```

#### group_members.txt
```
姓名1,email1@example.com
姓名2,email2@example.com
```

#### sql_credentials.txt
```ini
[MySQL]
host = localhost
port = 3306
user = username
password = password
database = database_name
```

#### llm_key.txt
单行 LLM 网关密钥（`newapi.gisphere.info`）：
```
sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

### 4. 验证安装

运行环境检查工具（在仓库根目录执行）：
```bash
python -m src.tools.check_setup
```

这会检查：
- ✓ Python版本
- ✓ 核心文件存在性
- ✓ 配置文件完整性
- ✓ 依赖包安装情况

---

## 使用方法

### 基本运行

```bash
# 启动（唯一入口）
python run.py

# 检查环境后运行
python -m src.tools.check_setup && python run.py
```

### 运行输出示例

```
============================================================
                  GISource 自动化系统                      
============================================================

✓ 距上次运行 5 分 30 秒，自动使用缓存操作员: 张三
预检查: 确保当前周期标题存在...
   当前周期: 【2025年11月24日 - 2025年11月30日】
✓ 当前周期标题已存在

步骤 1: 从Google Sheets获取数据...
✓ 数据加载完成

步骤 2: 更新大学中文名称...
✓ 更新了 0 行大学信息

步骤 3: 检查新大学...
✓ 没有新大学需要添加

步骤 4: 选择要处理的数据...
✓ 已选择数据行

步骤 5: 验证数据完整性...
✓ 数据验证通过

步骤 6: 插入数据到数据库...
✓ 成功插入数据，Event_ID: 1234

步骤 7: 更新Google Sheets...
✓ Google Sheets更新完成

步骤 8: 生成微信消息...
✓ 已生成微信群消息内容和职位缩写

步骤 9: 添加到微信公众号文档...
✓ 成功将信息插入到文档适当位置

步骤 10: 发送微信群消息邮件通知...
✓ 已发送微信消息通知到 张三

============================================================
                    所有步骤完成！
============================================================
```

### 如果授权失效

删除以下文件后重新运行（程序会重新打开浏览器授权）：
```bash
# Windows
del keys\token_sheets.json keys\token.json keys\token.pickle

# macOS/Linux
rm keys/token_sheets.json keys/token.json keys/token.pickle
```

---

## 配置说明

### 修改操作员姓名

编辑 `main.py` 中的 `get_operator_name()` 函数：

```python
def get_operator_name():
    return "你的名字"  # 修改这里
```

### 切换测试/生产环境

在 `main.py` 的 `process_and_insert_to_database()` 函数中：

```python
table_name = 'Coding_Test'  # 测试环境
# table_name = 'GISource'   # 生产环境
```

### 修改字典配置

编辑 `src/core/config.py` 中的字典：

- **COUNTRY_DICTIONARY** - 国家名称映射（中文→英文）
- **JOB_DICTIONARY** - 职位类型映射
- **SUBJECT_DICTIONARY** - 学科分类映射

示例：
```python
# 添加新国家
COUNTRY_DICTIONARY = {
    ...
    '新国家': 'New Country',
}

# 添加新职位类型
JOB_DICTIONARY = {
    ...
    "新职位": "新职位中文",
}
```

### 修改Google Sheets/Docs ID

在 `src/core/config.py` 中：

```python
SPREADSHEET_ID = 'your-spreadsheet-id'
DOCUMENT_ID = 'your-document-id'
```

### 配置网络代理（访问 Google）

`src/core/config.py` 的 `_detect_local_proxy()` 会按顺序处理：

1. 优先读取环境变量 `GOOGLE_API_PROXY` / `HTTPS_PROXY` / `HTTP_PROXY`（显式配置，直接采用）；
2. 否则探测常见本地代理端口，并**经候选代理请求一次 Google `generate_204` 验证**，确认确实可用才采用，避免把"占用了端口但并非代理"的程序误判为代理。

如需手动指定：
```bash
# Windows (PowerShell)
$env:GOOGLE_API_PROXY = "http://127.0.0.1:7897"
# macOS/Linux
export GOOGLE_API_PROXY="http://127.0.0.1:7897"
```
若环境可直连 Google（如海外服务器），代理探测会全部失败并返回 `None`，程序直连访问。

### 配置 LLM 模型链（Claude → GPT → Gemini）

内容组织（`src/integrations/google_docs.py`）使用统一网关 `newapi.gisphere.info`，按 `src/core/config.py` 的模型链逐个回退：

```python
OPENAI_BASE_URL = "https://newapi.gisphere.info/v1"
LLM_MODEL_CHAIN = [
    "claude-sonnet-4-6",   # 优先
    "gpt-5.5",             # 其次
    "gemini-3.5-flash",    # 最后兜底
]
```

任一模型遇到鉴权失败 / 限流 / 报错 / 空响应时自动尝试下一个；密钥放在 `keys/llm_key.txt`。可用模型清单参见同仓库 `LLM_Analysis/MODELS.md`。

---

## 核心模块说明

### 1. src/core/config.py - 配置中心
**功能**: 集中管理所有配置和常量

**包含内容**:
- 系统平台检测
- Google API配置
- 文件路径（跨平台兼容）
- 国家/职位/学科字典
- 必填字段定义

### 2. src/core/utils.py - 工具函数库
**主要函数**:
- `is_date()` - 检查字符串是否为日期
- `read_group_members()` - 读取组员信息
- `number_to_chinese_words()` - 数字转中文
- `convert_date_to_chinese()` - 日期转中文格式
- `calculate_week_range()` - 计算本周日期范围

### 3. src/integrations/google_sheets.py - 表格操作
**主要函数**:
- `authorize_credentials()` - 授权Google API（JSON 令牌优先，兼容旧 pickle，加载失败自动重新授权；凭据进程内缓存）
- `fetch_data()` - 获取表格数据（经 `execute_with_retry` 重试）
- `delete_rows_from_sheet()` - 删除行
- `append_data_to_sheet()` - 追加数据
- `update_data_in_sheet()` - 更新数据
- `batch_update_data_in_sheet()` - 批量更新多个范围（合并为一次 API 调用）

### 4. src/integrations/google_http.py - Google 网络层
**主要函数**:
- `setup_google_proxy_env()` / `refresh_credentials()` - 代理提示与经代理刷新令牌
- `build_google_service()` - 构建带代理的 Google API 服务（进程内缓存；重试前自动失效重建）
- `execute_with_retry()` - 对瞬时网络错误（超时/连接重置/SSL 中断/429/5xx）自动重试；4xx 等立即抛出

### 5. src/integrations/google_docs.py - 文档操作
**主要函数**:
- `build_docs_service()` - 构建Docs服务
- `call_llm_for_content_organization()` - 用 LLM 组织内容（Claude→GPT→Gemini 模型链回退）
- `append_to_document()` - 追加内容
- `add_wechat_content_to_doc()` - 添加微信公众号内容

### 6. src/integrations/database.py - 数据库操作
**主要函数**:
- `get_database_connection()` - 获取连接（默认15秒超时，最多重试3次，连接时隔离代理环境变量确保直连；库名由 `sql_credentials.txt` 的 `database` 决定，SQL 不写死）
- `get_gisource_data()` - 获取GISource数据
- `insert_event_to_database()` - 插入事件数据

### 7. src/integrations/email_sender.py - 邮件发送
**主要函数**:
- `send_email()` - 通用邮件发送（有代理时优先 Gmail API，失败回退代理 SMTP，再回退 QQmail）
- `_send_gmail_via_api()` - 经 Gmail API 发送（与 Sheets 同一代理通路）
- `send_reminder_emails()` - 批量提醒
- `send_error_notification()` - 错误通知
- `send_wechat_notification()` - 微信消息通知
- `_append_failed_email_record()` - 邮件最终发送失败时落盘记录到 `logs/failed_email_records.txt`

### 8. src/integrations/smtp_proxy.py - 代理 SMTP
**主要函数**:
- `create_smtp_client()` - 配置了代理时经 SOCKS5/HTTP 代理连接 Gmail SMTP（备用通道）

### 9. src/core/data_processor.py - 数据处理
**主要函数**:
- `create_sql_table()` - 创建SQL表格数据
- `generate_wechat_group_text()` - 生成微信群消息
- `convert_to_wechat_format()` - 转换为公众号格式

### 10. src/core/logger.py - 日志
**主要函数**:
- `log_program_run()` - 记录结构化运行日志（内存缓冲，结束时归档到 `logs/`）
- `log_llm_conversation()` - LLM 对话记录（`llm_logs/`）

### 11. src/main.py - 主流程
**功能**: 协调所有模块，控制程序流程；`fetch_dataframe()` 统一整表拉取与空表防御

**主要流程**:
1. 加载数据（Google Sheets）
2. 更新大学信息
3. 检查新大学
4. 智能选择要处理的行
5. 验证数据完整性
6. 插入到数据库
7. 更新Google Sheets
8. 生成微信消息
9. 更新公众号文档

---

## 工作流程

```
Google Sheets (Unfilled)
        ↓
    数据验证 & 清理
        ↓
    智能选择数据行
        ↓
    验证必填字段
        ↓
    数据格式转换
        ↓
    MySQL 数据库
        ↓
    ├→ Google Sheets (Filled)
    ├→ 微信群消息（邮件通知）
    └→ Google Docs（微信公众号）
```

### 详细步骤说明

**预检查**: 确保当前周期的日期标题（按本周范围动态生成）存在于指定 Google 文档中。同时确认操作人员身份（支持30分钟内免输缓存）。

1. **数据加载**: 从Google Sheets获取Unfilled和Filled数据，检查删除过期行及与Filled表中重复的数据（基于特定列对比）。
2. **更新大学信息**: 从数据库匹配并自动填充缺失的大学中文名称以及国家中文名称。
3. **检查新大学**: 将新大学添加到Universities表。
4. **智能选择**: 优先级算法选择数据（80%为"Soon"期限数据，10%为最近期限数据，10%随机有效行）。
5. **数据验证**: 检查必填字段，若发现错误自动将表内 Error 列标为1，并发送邮件通知验证者。
6. **数据库插入**: 生成新Event_ID并插入目标MySQL数据库表。
7. **更新表格**: 将已处理的数据行从 Unfilled 移到 Filled 工作表内。
8. **生成微信内容**: 使用 LLM (如果配置允许) 或内部逻辑生成职位缩写及微信群消息。
9. **公众号文档**: 依据职位缩写和周期标题，将格式化后内容以字典序平滑插入至Google Docs。
10. **邮件群发通知**: 给指定操作人员或主群发送整理好的微信文案。

---

## 自定义配置

### 添加新功能

1. **确定功能类型**
   - 数据处理 → `src/core/data_processor.py`
   - API操作 → `src/integrations/google_sheets.py` 或 `src/integrations/google_docs.py`
   - 数据库 → `src/integrations/database.py`
   - 邮件 → `src/integrations/email_sender.py`
   - 工具函数 → `src/core/utils.py`

2. **在相应模块添加函数**

3. **在src/main.py中调用**

4. **更新README文档**

### 扩展示例

**添加新数据源：**
```python
# 创建新模块 api_client.py
def fetch_external_data():
    pass

# 在src/main.py中导入并使用
from api_client import fetch_external_data
```

**添加新通知方式：**
```python
# 在email_sender.py中添加
def send_slack_notification(message):
    pass
```

---

## 故障排除

### 常见问题

#### Q1: 找不到 credentials.json
**A:** 请确保已从Google Cloud Console下载凭据文件，并放在项目根目录。

#### Q2: 邮件发送失败
**A:** 
- 检查是否使用了"应用专用密码"而非普通密码
- 确认已启用两步验证
- 检查 `email_credentials.txt` 是否包含 `[Gmail]` 与 `[QQmail]` 两个段
- 主通道 Gmail 失败时，系统会自动回退到 QQmail 发送
- 当前 QQmail 使用 `smtp.qq.com:465`（SSL）
- 若两条通道都失败，邮件内容会自动追加记录到 `logs/failed_email_records.txt`

#### Q3: 数据库连接失败
**A:**
- 检查 `sql_credentials.txt` 配置
- 确认数据库服务器可访问
- 检查用户名和密码
- 确认端口号正确（默认3306）

#### Q4: Google API授权失败
**A:**
- 删除 `keys/token_sheets.json`、`keys/token.json`、`keys/token.pickle`
- 重新运行程序
- 确保已启用 Google Sheets API、Google Docs API 和 Gmail API
- 检查 `credentials.json` 是否正确
- 若日志提示"读取 token.pickle 失败（版本不兼容）"，属正常兼容提示，程序会自动改用 JSON 令牌或重新授权

#### Q4b: 连接 Google 超时 / 无法访问
**A:**
- 国内环境需开启本地代理（Clash/V2Ray 等）后再运行
- 程序启动会打印 `✓ 使用代理访问 Google API: http://127.0.0.1:xxxx`；若未出现说明未探测到可用代理
- 可用环境变量显式指定：`GOOGLE_API_PROXY=http://127.0.0.1:7897`
- 代理探测会验证能否访问 Google，确认你的代理规则已放行 `googleapis.com`/`gstatic.com`

#### Q4c: 邮件经代理 SMTP 发送失败
**A:**
- 部分代理对 SMTP 隧道支持不佳（连上但收不到响应）。此时建议确保 Gmail API 路径可用（优先通道）
- 确认已安装 `PySocks`（`pip install -r requirements.txt`）

#### Q4d: LLM 内容组织失败 / 模型不可用
**A:**
- 检查 `keys/llm_key.txt` 是否存在且有效
- 模型链 Claude→GPT→Gemini 会自动回退；若全部失败，程序回退到内部默认规则
- `gemini-3.5-flash` 上游可能被停用而返回 403，属预期，由前两个模型兜住

#### Q5: 中文显示乱码（Windows）
**A:** 在运行前执行：
```bash
chcp 65001
```

或使用 `run.py` 启动脚本，会自动处理编码。

#### Q6: ImportError: No module named 'xxx'
**A:** 运行：
```bash
pip install -r requirements.txt
```

#### Q7: 没有数据可处理
**A:** 程序会发送提醒邮件到所有组员，提示添加内容。

### 调试技巧

1. **查看详细输出**: 程序会打印每个步骤的状态
2. **检查日志**: 邮件发送、数据库操作都有日志
3. **逐步调试**: 可以在src/main.py中注释掉某些步骤
4. **使用环境检查**: `python -m src.tools.check_setup` 快速检查环境配置

---

## 跨平台兼容性

程序使用以下方法确保跨平台兼容：

1. **路径处理**: 使用 `os.path.join()` 而非硬编码路径分隔符
2. **系统检测**: 使用 `platform.system()` 检测操作系统
3. **文件编码**: 统一使用 UTF-8 编码
4. **命令区分**: 文档中区分Windows和Unix系统的命令

### 跨平台代码示例

```python
# 自动检测系统
SYSTEM_PLATFORM = platform.system()  # 'Windows', 'Darwin', 'Linux'

# 跨平台路径（config.py 位于 src/core/，BASE_DIR 上溯两级指向仓库根）
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
KEYS_DIR = os.path.join(BASE_DIR, 'keys')
CREDENTIALS_FILE = os.path.join(KEYS_DIR, 'credentials.json')
```

---

## 开发指南

### 代码结构

- `src/core/config.py` - 所有配置和常量（含代理探测、LLM 模型链、IPv4 强制开关定义）
- `src/core/utils.py` - 通用工具函数
- `src/core/logger.py` - 日志模块（结构化运行日志并写入文件，重定向终端流）
- `src/core/data_processor.py` - 数据处理逻辑（含缩写与微信内容生成）
- `src/integrations/google_http.py` - Google API 代理/HTTP 客户端与请求重试（service 缓存）
- `src/integrations/google_sheets.py` - Google Sheets操作（JSON 令牌 + 重试 + 凭据缓存）
- `src/integrations/google_docs.py` - Google Docs操作 + LLM 内容组织（模型链）
- `src/integrations/database.py` - 数据库操作（重试 + 代理隔离）
- `src/integrations/email_sender.py` - 邮件功能（Gmail API / 代理 SMTP / QQmail）
- `src/integrations/smtp_proxy.py` - 经代理连接 Gmail SMTP（备用通道）
- `src/tools/check_setup.py` - 环境检查（python -m src.tools.check_setup）
- `src/main.py` - 主程序流程的10步骤调度

### 最佳实践

1. **模块化**: 单一职责原则，每个模块只负责一个功能
2. **配置分离**: 配置与代码分离，便于修改
3. **DRY原则**: 不要重复代码，复用工具函数
4. **清晰导入**: 明确的导入结构，便于理解依赖关系

### 与Jupyter Notebook对比

| 方面 | Jupyter Notebook | Python模块化 |
|------|-----------------|-------------|
| 运行方式 | 单元格逐个运行 | 一键运行 `python run.py` |
| 维护性 | 代码混在一起 | 模块清晰，易维护 |
| 复用性 | 难以复用 | 函数可独立复用 |
| 自动化 | 手动运行 | 完全自动化 |
| 版本控制 | 困难 | 友好 |
| 跨平台 | 需Jupyter环境 | 原生Python即可 |
| 代码行数 | 3000+ | ~1430（结构化） |

---

## 安全提示

### 凭据管理

以下文件包含敏感信息，**务必确保已被 `.gitignore` 保护、切勿提交**：
- `keys/credentials.json` - Google API凭据
- `keys/token_sheets.json` - Sheets + Gmail 令牌（JSON）
- `keys/token.json` - Google Docs令牌
- `keys/token.pickle` - 旧版令牌
- `keys/email_credentials.txt` - 邮箱密码
- `keys/llm_key.txt` - LLM 网关密钥
- `keys/sql_credentials.txt` - 数据库密码

**请勿将这些文件提交到版本控制系统！**

### 示例文件

提供了示例配置文件（可安全提交到Git）：
- `email_credentials.txt.example`
- `sql_credentials.txt.example`

使用时复制并重命名为实际文件名。

---

## 技术支持

如有问题，请联系：
- **GISource团队**
- **Email**: gisphere@gmail.com

或查看：
1. 本 README 文档
2. 运行 `python -m src.tools.check_setup` 检查环境
3. 查看终端输出的错误信息

---

## 版本历史

### v2.3.0 (当前版本 - 2026-07-05)
- 📁 **目录重组**：源码迁入 `src/` 包（core / integrations / tools，相对导入），入口统一 `python run.py`，环境检查改为 `python -m src.tools.check_setup`
- 🐛 **修复删除行 off-by-one**：`load_and_clean_data` 行号换算 `+2` 修正为 `+1`（实测原实现会误删目标行的下一行）
- 🐛 **修复写回范围**：`update_university_info` 写死 `A:Z`（表 31 列必报错）改为锚点单元格写法，并合并为一次 `values.batchUpdate`
- 🐛 **Deadline 不再原地转 Timestamp**：过期判断改用临时解析序列，写回不再崩溃、不再冲掉 'Soon'/'rolling' 等原文
- 🛡️ 空表防御：新增 `fetch_dataframe()`，整表拉取为空时给出明确报错
- ⚡ 凭据与 service 进程内缓存（网络重试时自动重建连接）
- 🗃️ 数据库表名不再写死 `TEST.` 前缀，统一由 `sql_credentials.txt` 的 `database` 决定
- 🧹 移除全局警告压制（改定向静音）、`applymap` 改版本无关写法、裸 `except` 收紧、requirements 移除 pycountry/configparser 并加版本上界

### v2.2.1 (2026-06-28)
- 🐛 修复 `select_row_to_process`：无 Soon 行时直接选择截止日期最近的一条；并修复"候选只有 1 个时权重数量不匹配导致崩溃"的问题
- ⚡ 消除重复请求：`update_google_sheets`/`validate_selected_row` 不再重复整表取数；Google Docs 多处"先取结构再重复 GET 取文本"改为从同一份文档结构抽取文本（`extract_text_from_document`）
- 🧹 删除未被调用的死函数 `generate_and_send_wechat_message`
- 🛡️ `adjust_data_to_columns` 对超长行截断，避免 DataFrame 列数不匹配；`select_row_to_process` 显式 `.copy()` 消除 SettingWithCopy 告警；清理 `create_job_title` 的空操作
- ⏳ 待确认（暂未改动）：`load_and_clean_data` 删除行索引 `x+2` 与 `update_google_sheets` 的 `x+1` 不一致（疑似 off-by-one）；`database.py` 中 `GISource` 与 `TEST.GISource` 的 schema 前缀不一致

### v2.2.0 (2026-06-28)
- 🌐 **网络代理支持**：新增 `google_http.py`，自动探测本地代理并经代理访问 Google；代理探测后会请求 `generate_204` 验证可用性，避免误判非代理端口
- ✉️ **Gmail API 发送**：检测到代理时优先经 Gmail API 发信（与 Sheets 同一代理通路），失败回退代理 SMTP（`smtp_proxy.py`）/ QQmail
- 🔐 **令牌改用 JSON**：Sheets/Gmail 令牌存为可移植的 `token_sheets.json`，跨 google-auth 版本均可加载；加载失败优雅重新授权而非崩溃（兼容旧 `token.pickle`）
- 🔁 **请求重试增强**：`execute_with_retry` 覆盖超时/连接重置/SSL 中断/429/5xx 等瞬时错误，4xx 立即失败
- 🤖 **LLM 模型回退链**：内容组织优先 Claude，其次 GPT，最后 Gemini（`LLM_MODEL_CHAIN`，统一走 `newapi` 网关），任一失败/空响应自动切换
- 💾 **数据库连接加固**：改用 `connection_timeout` + 3 次重试，连接期间隔离代理环境变量确保直连
- ⚡ **取数优化**：`load_and_clean_data` 每张表只请求一次（原先重复请求）
- 🔑 LLM 密钥文件由 `openai_key.txt` 改为 `llm_key.txt`；新增依赖 `PySocks`

### v2.1.0 (2026-05-02)
- ✨ 新增 Gmail -> QQmail 自动回退发送链路
- 🔐 QQmail 改为官方推荐 `465 + SSL` 连接模式
- 📝 新增邮件失败落盘记录（`logs/failed_email_records.txt`，追加写入）
- 📚 更新邮箱配置文档与故障排除说明

### v2.0.0 (2025-11-29)
- 🎉 重构为模块化Python项目
- ✨ 支持跨平台运行（Windows/macOS/Linux）
- 🔧 改进代码结构和可维护性
- 📚 添加完整文档
- 🔐 增强安全性（凭据文件独立）
- ⚡ 性能优化（批量操作、超时控制）
- 🐛 完善错误处理

### v1.0.0
- 初始Jupyter Notebook版本

---

## 许可证

本项目仅供内部使用。

---

## 致谢

感谢 GISource 团队所有成员的贡献！

---

**最后更新**: 2026-06-28  
**维护者**: GISource团队  
**项目状态**: 生产就绪 ✅
