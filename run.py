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
    # reconfigure 保留行缓冲，print 会实时输出（codecs.getwriter 会导致输出堆积到程序结束）
    sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
    sys.stderr.reconfigure(encoding='utf-8')

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 导入并运行主程序
if __name__ == "__main__":
    from src.main import main
    main()

