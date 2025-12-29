# GISource 自动化系统

自动化处理GISource学术信息的发布系统，支持Windows、macOS和Linux多平台运行。

**版本**: v2.0.0 | **最后更新**: 2025-11-29

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
- 🌍 **跨平台支持** - 完全支持Windows/macOS/Linux
- 🔄 **模块化设计** - 易于维护和扩展
- 🔐 **安全管理** - 凭据文件独立管理，不纳入版本控制

---

## 系统要求

- Python 3.7 或更高版本
- 稳定的网络连接
- Google API 访问权限（Google Sheets + Google Docs）
- MySQL 数据库访问权限
- Gmail 邮箱（用于发送通知）

---

## 快速开始

### 准备清单

在开始之前，请确保您有：
- [ ] Python 3.7+ 已安装
- [ ] Google账号（用于访问Google Sheets和Docs）
- [ ] Gmail邮箱（用于发送通知邮件）
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
4. 创建OAuth 2.0凭据（类型：桌面应用）
5. 下载JSON文件，重命名为 `credentials.json`
6. 将 `credentials.json` 放在项目根目录

#### 步骤 3: 配置邮箱（1分钟）

**获取Gmail应用专用密码：**
1. 访问：https://myaccount.google.com/security
2. 启用"两步验证"
3. 在"两步验证"页面，找到"应用专用密码"
4. 生成新密码（选择"邮件"和"Windows计算机"或其他设备）
5. 复制生成的16位密码（类似：`xxxx xxxx xxxx xxxx`）

**创建凭据文件：**

Windows:
```bash
copy email_credentials.txt.example email_credentials.txt
notepad email_credentials.txt
```

macOS/Linux:
```bash
cp email_credentials.txt.example email_credentials.txt
nano email_credentials.txt
```

编辑内容为：
```
your-email@gmail.com
xxxx xxxx xxxx xxxx
```

#### 步骤 4: 配置数据库（1分钟）

Windows:
```bash
copy sql_credentials.txt.example sql_credentials.txt
notepad sql_credentials.txt
```

macOS/Linux:
```bash
cp sql_credentials.txt.example sql_credentials.txt
nano sql_credentials.txt
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
python main.py
```

**macOS/Linux:**
```bash
python3 main.py
```

或使用快速启动脚本：
```bash
python run.py
```

### 首次运行说明

首次运行时：

1. **浏览器会自动打开**，要求您授权Google API
2. **选择您的Google账号**
3. **点击"允许"**授权访问
4. 授权成功后，浏览器会显示"The authentication flow has completed."
5. **返回终端**，程序将自动继续运行

授权完成后会生成两个文件：
- `token.pickle`
- `token.json`

以后运行将自动使用这些凭据，无需重新授权。

---

## 项目结构

```
Google_Sheet/
│
├── 🚀 启动文件
│   ├── main.py                    # 主程序入口
│   ├── run.py                     # 快速启动脚本（推荐）
│   └── check_setup.py             # 环境检查工具
│
├── ⚙️ 核心模块
│   ├── config.py                  # 配置管理（常量、字典）
│   ├── utils.py                   # 工具函数
│   ├── data_processor.py          # 数据处理和格式转换
│   ├── google_sheets.py           # Google Sheets API
│   ├── google_docs.py             # Google Docs API
│   ├── database.py                # MySQL数据库操作
│   └── email_sender.py            # 邮件发送功能
│
├── 📋 配置文件（需手动创建）
│   ├── credentials.json           # Google API凭据
│   ├── email_credentials.txt      # 邮箱凭据
│   ├── group_members.txt          # 组员信息
│   └── sql_credentials.txt        # 数据库凭据
│
├── 🔐 认证文件（自动生成）
│   ├── token.pickle               # Google API令牌
│   └── token.json                 # Google API令牌
│
└── 📚 其他
    ├── requirements.txt           # Python依赖
    ├── .gitignore                 # Git配置
    ├── VERSION.txt                # 版本信息
    └── Coding.ipynb               # 原始Notebook（参考）
```

---

## 详细安装步骤

### 1. 检查Python版本

```bash
python --version  # Windows
python3 --version # macOS/Linux
```

确保版本 >= 3.7

### 2. 安装依赖包

```bash
pip install -r requirements.txt
```

需要安装的包：
- pandas, numpy - 数据处理
- google-api-python-client - Google API
- mysql-connector-python - MySQL连接
- pytz, pycountry, inflect - 工具库

### 3. 配置文件设置

#### credentials.json（Google API）
从 Google Cloud Console 下载

#### email_credentials.txt
```
your-email@gmail.com
your-app-password
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

### 4. 验证安装

运行环境检查工具：
```bash
python check_setup.py
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
# 方式1：直接运行
python main.py

# 方式2：使用启动脚本（推荐）
python run.py

# 方式3：检查环境后运行
python check_setup.py && python main.py
```

### 运行输出示例

```
============================================================
                  GISource 自动化系统                      
============================================================

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
✓ 已发送微信消息通知

步骤 9: 添加到微信公众号文档...
✓ Job listing added to the document

============================================================
                    所有步骤完成！
============================================================
```

### 如果授权失效

删除以下文件后重新运行：
```bash
# Windows
del token.pickle token.json

# macOS/Linux
rm token.pickle token.json
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

编辑 `config.py` 中的字典：

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

在 `config.py` 中：

```python
SPREADSHEET_ID = 'your-spreadsheet-id'
DOCUMENT_ID = 'your-document-id'
```

---

## 核心模块说明

### 1. config.py - 配置中心
**功能**: 集中管理所有配置和常量

**包含内容**:
- 系统平台检测
- Google API配置
- 文件路径（跨平台兼容）
- 国家/职位/学科字典
- 必填字段定义

### 2. utils.py - 工具函数库
**主要函数**:
- `is_date()` - 检查字符串是否为日期
- `read_group_members()` - 读取组员信息
- `number_to_chinese_words()` - 数字转中文
- `convert_date_to_chinese()` - 日期转中文格式
- `calculate_week_range()` - 计算本周日期范围

### 3. google_sheets.py - 表格操作
**主要函数**:
- `authorize_credentials()` - 授权Google API
- `fetch_data()` - 获取表格数据
- `delete_rows_from_sheet()` - 删除行
- `append_data_to_sheet()` - 追加数据
- `update_data_in_sheet()` - 更新数据

### 4. google_docs.py - 文档操作
**主要函数**:
- `build_docs_service()` - 构建Docs服务
- `append_to_document()` - 追加内容
- `add_wechat_content_to_doc()` - 添加微信公众号内容

### 5. database.py - 数据库操作
**主要函数**:
- `get_database_connection()` - 获取连接（60秒超时）
- `get_gisource_data()` - 获取GISource数据
- `insert_event_to_database()` - 插入事件数据

### 6. email_sender.py - 邮件发送
**主要函数**:
- `send_email()` - 通用邮件发送
- `send_reminder_emails()` - 批量提醒
- `send_error_notification()` - 错误通知
- `send_wechat_notification()` - 微信消息通知

### 7. data_processor.py - 数据处理
**主要函数**:
- `create_sql_table()` - 创建SQL表格数据
- `generate_wechat_group_text()` - 生成微信群消息
- `convert_to_wechat_format()` - 转换为公众号格式

### 8. main.py - 主程序
**功能**: 协调所有模块，控制程序流程

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

1. **数据加载**: 从Google Sheets获取Unfilled和Filled数据，删除过期行
2. **更新大学信息**: 从数据库匹配并填充缺失的大学中文名称
3. **检查新大学**: 将新大学添加到Universities表
4. **智能选择**: 优先选择"Soon"截止的数据（80%概率）
5. **数据验证**: 检查必填字段，有误发送邮件通知
6. **数据库插入**: 生成新Event_ID并插入GISource表
7. **更新表格**: 从Unfilled移到Filled
8. **微信消息**: 生成微信群消息，发送邮件通知
9. **公众号文档**: 添加格式化内容到Google Docs

---

## 自定义配置

### 添加新功能

1. **确定功能类型**
   - 数据处理 → `data_processor.py`
   - API操作 → `google_sheets.py` 或 `google_docs.py`
   - 数据库 → `database.py`
   - 邮件 → `email_sender.py`
   - 工具函数 → `utils.py`

2. **在相应模块添加函数**

3. **在main.py中调用**

4. **更新README文档**

### 扩展示例

**添加新数据源：**
```python
# 创建新模块 api_client.py
def fetch_external_data():
    pass

# 在main.py中导入并使用
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
- 检查 `email_credentials.txt` 格式是否正确（两行，无多余空格）

#### Q3: 数据库连接失败
**A:**
- 检查 `sql_credentials.txt` 配置
- 确认数据库服务器可访问
- 检查用户名和密码
- 确认端口号正确（默认3306）

#### Q4: Google API授权失败
**A:**
- 删除 `token.pickle` 和 `token.json`
- 重新运行程序
- 确保已启用Google Sheets API和Google Docs API
- 检查 `credentials.json` 是否正确

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
3. **逐步调试**: 可以在main.py中注释掉某些步骤
4. **使用check_setup.py**: 快速检查环境配置

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

# 跨平台路径
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, 'credentials.json')
```

---

## 开发指南

### 代码结构

- `config.py` - 所有配置和常量
- `utils.py` - 通用工具函数
- `google_sheets.py` - Google Sheets操作
- `google_docs.py` - Google Docs操作
- `database.py` - 数据库操作
- `email_sender.py` - 邮件功能
- `data_processor.py` - 数据处理逻辑
- `main.py` - 主程序流程

### 最佳实践

1. **模块化**: 单一职责原则，每个模块只负责一个功能
2. **配置分离**: 配置与代码分离，便于修改
3. **DRY原则**: 不要重复代码，复用工具函数
4. **清晰导入**: 明确的导入结构，便于理解依赖关系

### 与Jupyter Notebook对比

| 方面 | Jupyter Notebook | Python模块化 |
|------|-----------------|-------------|
| 运行方式 | 单元格逐个运行 | 一键运行 `python main.py` |
| 维护性 | 代码混在一起 | 模块清晰，易维护 |
| 复用性 | 难以复用 | 函数可独立复用 |
| 自动化 | 手动运行 | 完全自动化 |
| 版本控制 | 困难 | 友好 |
| 跨平台 | 需Jupyter环境 | 原生Python即可 |
| 代码行数 | 3000+ | ~1430（结构化） |

---

## 安全提示

### 凭据管理

以下文件包含敏感信息，已被 `.gitignore` 保护：
- `credentials.json` - Google API凭据
- `token.pickle` - Google Sheets令牌
- `token.json` - Google Docs令牌
- `email_credentials.txt` - 邮箱密码
- `sql_credentials.txt` - 数据库密码

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
2. 运行 `python check_setup.py` 检查环境
3. 查看终端输出的错误信息

---

## 版本历史

### v2.0.0 (当前版本 - 2025-11-29)
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

**最后更新**: 2025-11-29  
**维护者**: GISource团队  
**项目状态**: 生产就绪 ✅
