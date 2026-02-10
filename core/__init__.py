"""核心工具模块 - 数学处理、OCR客户端、PDF转换引擎"""

__version__ = "1.0.0"

import sys
import os


def get_app_dir():
    """获取应用程序目录（兼容PyInstaller打包）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    # core/ 包的上一层即项目根目录
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
