# -*- coding: utf-8 -*-
"""
95598会话抽检工具 - PythonAnywhere/WSGI 入口
只提供 Word 报告生成 API
"""
import sys
import os

# 添加项目路径（PythonAnywhere 环境）
# 格式: /home/你的用户名/仓库名
USERNAME = os.environ.get('PYTHONANYWHERE_USER', 'your-username')
PROJECT_DIR = f'/home/{USERNAME}/95598-Smart-Customer-Service-Session-Analysis-Tool-backend'

if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)

os.environ.setdefault('DJANGO_SETTINGS_MODULE', '')

# 导入 Flask 应用
from server_flask import app as application
