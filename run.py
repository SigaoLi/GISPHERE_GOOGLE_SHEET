#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
快速启动脚本
跨平台支持（Windows/macOS/Linux）
"""
import sys
import os

# 确保使用UTF-8编码
if sys.platform.startswith('win'):
    # Windows系统设置控制台编码
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 导入并运行主程序
if __name__ == "__main__":
    from main import main
    main()

