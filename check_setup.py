#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
环境检查脚本
检查所有必需的文件和依赖是否正确配置
"""
import os
import sys
import platform


def print_header(text):
    """打印标题"""
    print("\n" + "=" * 60)
    print(text.center(60))
    print("=" * 60)


def print_check(item, status, message=""):
    """打印检查结果"""
    status_symbol = "✓" if status else "✗"
    status_text = "通过" if status else "失败"
    print(f"{status_symbol} {item}: {status_text}")
    if message:
        print(f"  → {message}")


def check_python_version():
    """检查Python版本"""
    version = sys.version_info
    is_valid = version.major == 3 and version.minor >= 7
    version_str = f"{version.major}.{version.minor}.{version.micro}"
    print_check(
        "Python版本", 
        is_valid, 
        f"当前版本: {version_str} (需要 3.7+)"
    )
    return is_valid


def check_system():
    """检查操作系统"""
    system = platform.system()
    systems = {"Windows": "Windows", "Darwin": "macOS", "Linux": "Linux"}
    system_name = systems.get(system, "未知")
    print_check(
        "操作系统",
        True,
        f"{system_name} ({system})"
    )
    return True


def check_file_exists(filepath, description):
    """检查文件是否存在"""
    exists = os.path.exists(filepath)
    print_check(description, exists, filepath if exists else f"缺失: {filepath}")
    return exists


def check_dependencies():
    """检查依赖包"""
    print("\n检查依赖包...")
    required_packages = {
        'pandas': 'pandas',
        'numpy': 'numpy',
        'googleapiclient': 'google-api-python-client',
        'mysql.connector': 'mysql-connector-python',
        'pytz': 'pytz',
        'pycountry': 'pycountry',
        'inflect': 'inflect'
    }
    
    all_ok = True
    for module_name, package_name in required_packages.items():
        try:
            __import__(module_name)
            print_check(f"依赖包 {package_name}", True)
        except ImportError:
            print_check(f"依赖包 {package_name}", False, f"未安装，请运行: pip install {package_name}")
            all_ok = False
    
    return all_ok


def main():
    """主函数"""
    print_header("GISource 环境检查")
    
    print("\n系统环境检查:")
    checks = []
    
    # 检查Python版本
    checks.append(check_python_version())
    
    # 检查操作系统
    checks.append(check_system())
    
    print("\n核心文件检查:")
    # 核心模块
    core_files = [
        ('main.py', '主程序'),
        ('config.py', '配置模块'),
        ('utils.py', '工具模块'),
        ('google_sheets.py', 'Google Sheets模块'),
        ('google_docs.py', 'Google Docs模块'),
        ('database.py', '数据库模块'),
        ('email_sender.py', '邮件模块'),
        ('data_processor.py', '数据处理模块'),
        ('requirements.txt', '依赖清单')
    ]
    
    for filename, description in core_files:
        checks.append(check_file_exists(filename, description))
    
    print("\n配置文件检查:")
    # 配置文件（必需）
    config_files = [
        ('group_members.txt', '组员信息'),
        ('email_credentials.txt', '邮箱凭据'),
        ('sql_credentials.txt', '数据库凭据'),
        ('credentials.json', 'Google API凭据')
    ]
    
    config_ok = []
    for filename, description in config_files:
        config_ok.append(check_file_exists(filename, description))
    
    # 依赖包检查
    deps_ok = check_dependencies()
    checks.append(deps_ok)
    
    # 总结
    print_header("检查总结")
    
    core_passed = all(checks)
    config_passed = all(config_ok)
    
    print(f"\n核心文件: {'✓ 全部通过' if core_passed else '✗ 存在问题'}")
    print(f"配置文件: {'✓ 全部存在' if config_passed else '✗ 缺少配置'}")
    print(f"依赖包: {'✓ 已安装' if deps_ok else '✗ 缺少依赖'}")
    
    if not core_passed:
        print("\n⚠️  核心文件缺失！请确保所有Python模块都存在。")
        return False
    
    if not deps_ok:
        print("\n⚠️  缺少依赖包！请运行以下命令安装：")
        print("   pip install -r requirements.txt")
        return False
    
    if not config_passed:
        print("\n⚠️  配置文件缺失！请按以下步骤配置：")
        print("\n1. 配置Google API:")
        print("   - 从Google Cloud Console下载credentials.json")
        print("   - 放在项目根目录")
        print("\n2. 配置邮箱:")
        print("   - 复制 email_credentials.txt.example")
        print("   - 重命名为 email_credentials.txt")
        print("   - 填写邮箱和应用专用密码")
        print("\n3. 配置数据库:")
        print("   - 复制 sql_credentials.txt.example")
        print("   - 重命名为 sql_credentials.txt")
        print("   - 填写数据库连接信息")
        print("\n详细说明请参考 QUICKSTART.md")
        return False
    
    print("\n" + "=" * 60)
    print("🎉 所有检查通过！系统已准备就绪。".center(60))
    print("=" * 60)
    print("\n可以运行以下命令启动系统：")
    if platform.system() == "Windows":
        print("   python main.py")
    else:
        print("   python3 main.py")
    print("\n或使用快速启动脚本：")
    if platform.system() == "Windows":
        print("   python run.py")
    else:
        print("   python3 run.py")
    print()
    
    return True


if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n检查被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n⚠️  发生错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

