"""
PDF转换工具 — 主入口

用法:
    python pdf_converter.py

模块结构:
    core/               核心工具（数学处理、OCR客户端、转换引擎）
    converters/          转换器（每种功能一个文件，易于扩展）
    ui/                  界面（主应用、设置对话框）
    pdf_converter.py     本文件 — 入口 + 日志配置
"""

import logging
import os
import sys
import tkinter as tk


def _setup_logging():
    """配置日志输出到文件 + 控制台"""
    if getattr(sys, 'frozen', False):
        log_dir = os.path.dirname(sys.executable)
    else:
        log_dir = os.path.dirname(os.path.abspath(__file__))

    log_file = os.path.join(log_dir, 'pdf_converter.log')
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8', mode='a'),
            logging.StreamHandler(),
        ]
    )


def main():
    """主函数"""
    _setup_logging()

    from ui.app import PDFConverterApp
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
