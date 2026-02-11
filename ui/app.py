"""
PDF转换工具主应用类

负责：UI创建/布局、事件处理、进度追踪、设置加载保存、背景图片。
转换逻辑委托给 converters/ 模块。
"""

import io
import json
import logging
import os
import shutil
import sys
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox

from core import get_app_dir
from core.ocr_client import simple_encrypt, simple_decrypt, BaiduOCRClient, REQUESTS_AVAILABLE
from core.progress_converter import PDF2DOCX_AVAILABLE
from core.history import ConversionHistory
from converters.pdf_to_word import PDFToWordConverter
from converters.pdf_to_image import PDFToImageConverter
from converters.pdf_merge import PDFMergeConverter
from converters.pdf_split import PDFSplitConverter
from converters.image_to_pdf import ImageToPDFConverter, SUPPORTED_IMAGE_EXTS
from converters.pdf_watermark import PDFWatermarkConverter
from converters.pdf_encrypt import PDFEncryptConverter
from converters.pdf_compress import PDFCompressConverter, COMPRESS_PRESETS
from converters.pdf_extract import PDFExtractConverter
from converters.pdf_ocr import PDFOCRConverter
from converters.pdf_to_excel import PDFToExcelConverter, TABLE_STRATEGIES
from converters.pdf_batch_extract import PDFBatchExtractConverter
from converters.pdf_stamp_batch import PDFBatchStampConverter

try:
    from PIL import Image, ImageTk, ImageDraw
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import fitz
    FITZ_UI_AVAILABLE = True
except ImportError:
    FITZ_UI_AVAILABLE = False

# 拖拽支持（可选依赖）
try:
    import windnd
    WINDND_AVAILABLE = True
except ImportError:
    WINDND_AVAILABLE = False


# 所有支持的功能列表
ALL_FUNCTIONS = ["PDF转Word", "PDF转图片", "PDF合并", "PDF拆分", "图片转PDF", "PDF加水印", "PDF加密/解密", "PDF压缩", "PDF提取/删页", "OCR可搜索PDF", "PDF转Excel", "PDF批量文本/图片提取", "PDF批量盖章"]


class PDFConverterApp:
    """PDF转换工具主应用类"""

    def __init__(self, root):
        self.root = root
        from core import __version__
        self.root.title(f"PDF转换工具 v{__version__}")
        self.root.geometry("500x580")
        self.root.resizable(False, False)

        # 设置窗口图标（支持打包后路径）
        try:
            import sys
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            icon_path = os.path.join(base_path, 'logo.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                # 同时设置任务栏图标
                try:
                    import ctypes
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                        'PDFConverter.App')
                except Exception:
                    pass
        except Exception:
            pass

        # --- 通用变量 ---
        self.selected_file = tk.StringVar()
        self.status_message = tk.StringVar(value="就绪")
        self.total_pages = 0
        self.total_steps = 0
        self.start_time = None
        self.current_page_id = None
        self.current_page_index = None
        self.current_page_total = None
        self.current_phase = None
        self.page_start_time = None
        self.page_timeout_seconds = 60
        self.page_timer_job = None
        self.current_eta_text = ""
        self.base_status_text = ""
        self.conversion_active = False
        self._state_lock = threading.Lock()  # 保护跨线程共享状态
        self.page_start_var = tk.StringVar()
        self.page_end_var = tk.StringVar()
        self.title_text_var = tk.StringVar(value="PDF转换工具")
        self.settings_path = os.path.join(get_app_dir(), "settings.json")

        # --- 背景/面板 ---
        self.bg_image_path = None
        self.bg_image = None
        self.bg_pil = None
        self.bg_label = None
        self.panel_opacity_var = tk.DoubleVar(value=85.0)
        self.panel_padding = 20
        self.panel_image = None
        self.panel_canvas = None
        self.panel_image_id = None
        self.resize_job = None
        self.panel_resize_job = None
        self.progress_y = 290
        self.progress_text_y = 325
        self.btn_y = 370
        self.dnd_y = 410

        # --- 功能选择 ---
        self.current_function_var = tk.StringVar(value="PDF转Word")
        self.selected_files_list = []

        # --- PDF转图片选项 ---
        self.image_dpi_var = tk.StringVar(value="200")
        self.image_format_var = tk.StringVar(value="PNG")

        # --- OCR & 公式识别选项 ---
        self.ocr_enabled_var = tk.BooleanVar(value=False)
        self.formula_api_enabled_var = tk.BooleanVar(value=False)

        # --- PDF拆分选项 ---
        self.split_mode_var = tk.StringVar(value="每页一个PDF")
        self.split_param_var = tk.StringVar()

        # --- 图片转PDF选项 ---
        self.page_size_var = tk.StringVar(value="A4")

        # --- PDF加水印选项 ---
        self.watermark_text_var = tk.StringVar(value="机密文件")
        self.watermark_opacity_var = tk.StringVar(value="0.3")
        self.watermark_rotation_var = tk.StringVar(value="45")
        self.watermark_fontsize_var = tk.StringVar(value="40")
        self.watermark_position_var = tk.StringVar(value="平铺")
        self.watermark_image_path = None

        # --- PDF加密/解密选项 ---
        self.encrypt_mode_var = tk.StringVar(value="加密")
        self.user_password_var = tk.StringVar()
        self.owner_password_var = tk.StringVar()
        self.allow_print_var = tk.BooleanVar(value=True)
        self.allow_copy_var = tk.BooleanVar(value=True)
        self.allow_modify_var = tk.BooleanVar(value=False)
        self.allow_annotate_var = tk.BooleanVar(value=True)

        # --- 批量文本/图片提取选项 ---
        self.batch_text_enabled_var = tk.BooleanVar(value=True)
        self.batch_image_enabled_var = tk.BooleanVar(value=True)
        self.batch_text_format_var = tk.StringVar(value="txt")
        self.batch_text_mode_var = tk.StringVar(value="合并为一个文件")
        self.batch_preserve_layout_var = tk.BooleanVar(value=True)
        self.batch_ocr_enabled_var = tk.BooleanVar(value=False)
        self.batch_pages_var = tk.StringVar()
        self.batch_image_per_page_var = tk.BooleanVar(value=False)
        self.batch_image_dedupe_var = tk.BooleanVar(value=False)
        self.batch_image_format_var = tk.StringVar(value="原格式")
        self.batch_zip_enabled_var = tk.BooleanVar(value=False)

        # --- 批量盖章选项 ---
        self.stamp_mode_var = tk.StringVar(value="普通章")
        self.stamp_pages_var = tk.StringVar()
        self.stamp_opacity_var = tk.StringVar(value="0.85")
        self.stamp_position_var = tk.StringVar(value="右下")
        self.stamp_size_ratio_var = tk.StringVar(value="0.18")
        self.stamp_image_path = ""
        self.stamp_image_paths = []
        self.stamp_selected_image_idx = 0
        self.stamp_profiles = {}
        self.stamp_qr_text_var = tk.StringVar()
        self.stamp_seam_side_var = tk.StringVar(value="右侧")
        self.stamp_seam_align_var = tk.StringVar(value="居中")
        self.stamp_seam_overlap_var = tk.StringVar(value="0.25")
        self.stamp_template_path = ""
        self.stamp_remove_white_bg_var = tk.BooleanVar(value=False)
        self.stamp_preview_info_var = tk.StringVar(value="")
        self.stamp_preview_profile = {
            "x_ratio": 0.85,
            "y_ratio": 0.85,
            "size_ratio": 0.18,
            "opacity": 0.85,
        }
        self._stamp_preview_state = {}

        # --- API 配置 ---
        self.api_provider = "baidu"
        self.baidu_api_key = ""
        self.baidu_secret_key = ""
        self.xslt_path = None
        self._baidu_client = None

        # --- 转换历史 ---
        self.history = ConversionHistory()

        # --- 初始化 ---
        self.create_ui()
        self.load_settings()
        self.check_dependencies()

        # --- 拖拽支持 ---
        if WINDND_AVAILABLE:
            try:
                windnd.hook_dropfiles(self.root, func=self._on_drop_files)
            except Exception:
                pass

    # ==========================================================
    # UI 创建
    # ==========================================================

    def create_ui(self):
        """创建用户界面 - Canvas直绘实现透明面板"""
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.panel_canvas = tk.Canvas(self.root, highlightthickness=0, bd=0)
        self.panel_canvas.grid(
            row=0, column=0, sticky="nsew",
            padx=self.panel_padding, pady=self.panel_padding
        )

        # 设置按钮
        self.settings_btn = tk.Button(
            self.panel_canvas, text="⚙", font=("Microsoft YaHei", 12),
            relief=tk.FLAT, padx=4, cursor='hand2',
            command=self.open_settings_window
        )
        self.cv_settings = self.panel_canvas.create_window(
            5, 5, window=self.settings_btn, anchor="nw")

        # 历史记录按钮
        self.history_btn = tk.Button(
            self.panel_canvas, text="📋", font=("Microsoft YaHei", 12),
            relief=tk.FLAT, padx=4, cursor='hand2',
            command=self.open_history_window
        )
        self.cv_history = self.panel_canvas.create_window(
            40, 5, window=self.history_btn, anchor="nw")

        # 标题
        self.cv_title = self.panel_canvas.create_text(
            0, 35, text=self.title_text_var.get(),
            font=("Microsoft YaHei", 26, "bold"), anchor="n"
        )
        self.title_text_var.trace_add("write", self._on_title_var_changed)

        # 功能选择器
        func_frame = tk.Frame(self.panel_canvas)
        tk.Label(func_frame, text="功能:", font=("Microsoft YaHei", 10, "bold")).pack(side=tk.LEFT)
        self.func_combo = ttk.Combobox(
            func_frame, textvariable=self.current_function_var,
            values=ALL_FUNCTIONS,
            state='readonly', font=("Microsoft YaHei", 10), width=14
        )
        self.func_combo.pack(side=tk.LEFT, padx=(8, 0))
        self.func_combo.bind("<<ComboboxSelected>>", self._on_function_changed)
        self.cv_subtitle = self.panel_canvas.create_window(
            0, 75, window=func_frame, anchor="n"
        )

        # 文件选择区
        self.cv_section1 = self.panel_canvas.create_text(
            15, 105, text="选择PDF文件（可多选）",
            font=("Microsoft YaHei", 11, "bold"), anchor="nw"
        )
        file_frame = tk.Frame(self.panel_canvas)
        self.file_entry = tk.Entry(
            file_frame, textvariable=self.selected_file,
            font=("Microsoft YaHei", 10), state='readonly'
        )
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        self.order_btn = tk.Button(
            file_frame, text="排序", command=self._open_file_order_dialog,
            font=("Microsoft YaHei", 9), padx=6, cursor='hand2'
        )
        # 排序按钮默认隐藏，多文件时显示
        tk.Button(
            file_frame, text="浏览...", command=self.browse_file,
            font=("Microsoft YaHei", 10), padx=20, cursor='hand2'
        ).pack(side=tk.LEFT, padx=(10, 0), ipady=6)
        self.cv_file_frame = self.panel_canvas.create_window(
            15, 130, window=file_frame, anchor="nw", width=1
        )

        # 页范围（PDF转Word / PDF转图片 使用）
        self.cv_section2 = self.panel_canvas.create_text(
            15, 185, text="页范围（可选）",
            font=("Microsoft YaHei", 11, "bold"), anchor="nw"
        )
        range_frame = tk.Frame(self.panel_canvas)
        tk.Label(range_frame, text="起始页:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        tk.Entry(range_frame, textvariable=self.page_start_var, width=6,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=(6, 20))
        tk.Label(range_frame, text="结束页:", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)
        tk.Entry(range_frame, textvariable=self.page_end_var, width=6,
                 font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=(6, 20))
        tk.Label(range_frame, text="留空表示全部页（页码从1开始）",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.cv_range_frame = self.panel_canvas.create_window(
            15, 210, window=range_frame, anchor="nw"
        )

        # 转换选项区（Word模式）
        self.word_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.word_options_frame, text="转换选项:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        self.ocr_cb = tk.Checkbutton(
            self.word_options_frame, text="OCR识别(扫描件)",
            variable=self.ocr_enabled_var, font=("Microsoft YaHei", 9),
            command=self._on_option_changed
        )
        self.ocr_cb.pack(side=tk.LEFT, padx=(8, 0))
        self.formula_cb = tk.Checkbutton(
            self.word_options_frame, text="公式智能识别",
            variable=self.formula_api_enabled_var, font=("Microsoft YaHei", 9),
            command=self._on_option_changed
        )
        self.formula_cb.pack(side=tk.LEFT, padx=(8, 0))
        self.cv_formula_frame = self.panel_canvas.create_window(
            15, 245, window=self.word_options_frame, anchor="nw"
        )

        # 转换选项区（图片模式）
        self.image_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.image_options_frame, text="输出设置:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        tk.Label(self.image_options_frame, text="DPI:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Combobox(
            self.image_options_frame, textvariable=self.image_dpi_var,
            values=["72", "150", "200", "300", "600"],
            width=5, font=("Microsoft YaHei", 9), state='readonly'
        ).pack(side=tk.LEFT, padx=(4, 0))
        tk.Label(self.image_options_frame, text="格式:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(14, 0))
        ttk.Combobox(
            self.image_options_frame, textvariable=self.image_format_var,
            values=["PNG", "JPEG"],
            state='readonly', width=6, font=("Microsoft YaHei", 9)
        ).pack(side=tk.LEFT, padx=(4, 0))
        self.cv_image_options = self.panel_canvas.create_window(
            15, 245, window=self.image_options_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_image_options, state='hidden')

        # 合并信息区 (y=210, 与 range_frame 同位)
        self.merge_info_frame = tk.Frame(self.panel_canvas)
        self.merge_info_label = tk.Label(
            self.merge_info_frame, text="请选择至少2个PDF文件，将按选择顺序合并",
            font=("Microsoft YaHei", 9), fg="#666"
        )
        self.merge_info_label.pack(side=tk.LEFT)
        self.cv_merge_info = self.panel_canvas.create_window(
            15, 210, window=self.merge_info_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_merge_info, state='hidden')

        # 拆分选项区 (y=210)
        self.split_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.split_options_frame, text="模式:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.split_combo = ttk.Combobox(
            self.split_options_frame, textvariable=self.split_mode_var,
            values=["每页一个PDF", "每N页一个PDF", "按范围拆分"],
            state='readonly', font=("Microsoft YaHei", 9), width=12
        )
        self.split_combo.pack(side=tk.LEFT, padx=(6, 0))
        self.split_combo.bind("<<ComboboxSelected>>", self._on_split_mode_changed)
        self.split_param_label = tk.Label(
            self.split_options_frame, text="", font=("Microsoft YaHei", 9))
        self.split_param_label.pack(side=tk.LEFT, padx=(14, 0))
        self.split_param_entry = tk.Entry(
            self.split_options_frame, textvariable=self.split_param_var,
            width=18, font=("Microsoft YaHei", 9), state='disabled'
        )
        self.split_param_entry.pack(side=tk.LEFT, padx=(6, 0))
        self.split_param_hint = tk.Label(
            self.split_options_frame, text="", font=("Microsoft YaHei", 8), fg="#888"
        )
        self.split_param_hint.pack(side=tk.LEFT, padx=(6, 0))
        self.cv_split_options = self.panel_canvas.create_window(
            15, 210, window=self.split_options_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_split_options, state='hidden')

        # 图片转PDF选项区 (y=210)
        self.img2pdf_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.img2pdf_options_frame, text="页面尺寸:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        ttk.Combobox(
            self.img2pdf_options_frame, textvariable=self.page_size_var,
            values=["A4", "A3", "Letter", "Legal", "自适应"],
            state='readonly', font=("Microsoft YaHei", 9), width=8
        ).pack(side=tk.LEFT, padx=(8, 0))
        tk.Label(self.img2pdf_options_frame,
                 text="（自适应 = 页面大小匹配图片）",
                 font=("Microsoft YaHei", 8), fg="#888"
                 ).pack(side=tk.LEFT, padx=(10, 0))
        self.cv_img2pdf_options = self.panel_canvas.create_window(
            15, 210, window=self.img2pdf_options_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_img2pdf_options, state='hidden')

        # 水印选项区 (y=210)
        self.watermark_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.watermark_options_frame, text="文字:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        tk.Entry(self.watermark_options_frame, textvariable=self.watermark_text_var,
                 width=10, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(4, 0))
        tk.Button(self.watermark_options_frame, text="选图片",
                  font=("Microsoft YaHei", 8), command=self._choose_watermark_image,
                  cursor='hand2').pack(side=tk.LEFT, padx=(8, 0))
        self.watermark_img_label = tk.Label(self.watermark_options_frame, text="",
                 font=("Microsoft YaHei", 8), fg="#666")
        self.watermark_img_label.pack(side=tk.LEFT, padx=(4, 0))
        self.cv_watermark_options = self.panel_canvas.create_window(
            15, 210, window=self.watermark_options_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_watermark_options, state='hidden')

        # 水印详细选项区 (y=245)
        self.watermark_detail_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.watermark_detail_frame, text="透明度:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        ttk.Combobox(
            self.watermark_detail_frame, textvariable=self.watermark_opacity_var,
            values=["0.1", "0.2", "0.3", "0.5", "0.7"],
            width=4, font=("Microsoft YaHei", 9), state='readonly'
        ).pack(side=tk.LEFT, padx=(4, 0))
        tk.Label(self.watermark_detail_frame, text="位置:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Combobox(
            self.watermark_detail_frame, textvariable=self.watermark_position_var,
            values=["平铺", "居中", "左上角", "右上角", "左下角", "右下角"],
            width=6, font=("Microsoft YaHei", 9), state='readonly'
        ).pack(side=tk.LEFT, padx=(4, 0))
        tk.Label(self.watermark_detail_frame, text="（文字和图片二选一，都填则用图片）",
                 font=("Microsoft YaHei", 8), fg="#888").pack(side=tk.LEFT, padx=(8, 0))
        self.cv_watermark_detail = self.panel_canvas.create_window(
            15, 245, window=self.watermark_detail_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_watermark_detail, state='hidden')

        # 加密/解密选项区 (y=210)
        self.encrypt_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.encrypt_options_frame, text="模式:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.encrypt_mode_combo = ttk.Combobox(
            self.encrypt_options_frame, textvariable=self.encrypt_mode_var,
            values=["加密", "解密"], state='readonly',
            width=5, font=("Microsoft YaHei", 9)
        )
        self.encrypt_mode_combo.pack(side=tk.LEFT, padx=(4, 0))
        self.encrypt_mode_combo.bind("<<ComboboxSelected>>", self._on_encrypt_mode_changed)
        self.encrypt_pw_label = tk.Label(self.encrypt_options_frame, text="打开密码:",
                 font=("Microsoft YaHei", 9))
        self.encrypt_pw_label.pack(side=tk.LEFT, padx=(8, 0))
        self.encrypt_pw_entry = tk.Entry(self.encrypt_options_frame,
                 textvariable=self.user_password_var,
                 width=10, font=("Microsoft YaHei", 9), show="*")
        self.encrypt_pw_entry.pack(side=tk.LEFT, padx=(4, 0))
        self.encrypt_owner_label = tk.Label(self.encrypt_options_frame, text="权限密码:",
                 font=("Microsoft YaHei", 9))
        self.encrypt_owner_label.pack(side=tk.LEFT, padx=(8, 0))
        self.encrypt_owner_entry = tk.Entry(self.encrypt_options_frame,
                 textvariable=self.owner_password_var,
                 width=10, font=("Microsoft YaHei", 9), show="*")
        self.encrypt_owner_entry.pack(side=tk.LEFT, padx=(4, 0))
        self.cv_encrypt_options = self.panel_canvas.create_window(
            15, 210, window=self.encrypt_options_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_encrypt_options, state='hidden')

        # 加密权限选项区 (y=245)
        self.encrypt_perm_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.encrypt_perm_frame, text="允许操作:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        tk.Checkbutton(self.encrypt_perm_frame, text="打印",
                       variable=self.allow_print_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(8, 0))
        tk.Checkbutton(self.encrypt_perm_frame, text="复制",
                       variable=self.allow_copy_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(8, 0))
        tk.Checkbutton(self.encrypt_perm_frame, text="修改",
                       variable=self.allow_modify_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(8, 0))
        tk.Checkbutton(self.encrypt_perm_frame, text="注释",
                       variable=self.allow_annotate_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(8, 0))
        self.cv_encrypt_perm = self.panel_canvas.create_window(
            15, 245, window=self.encrypt_perm_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_encrypt_perm, state='hidden')

        # PDF压缩选项区 (y=210)
        self.compress_level_var = tk.StringVar(value='标准压缩')
        self.compress_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.compress_options_frame, text="压缩级别:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        for level in COMPRESS_PRESETS:
            tk.Radiobutton(
                self.compress_options_frame, text=level,
                variable=self.compress_level_var, value=level,
                font=("Microsoft YaHei", 9),
                command=self._on_compress_level_changed,
            ).pack(side=tk.LEFT, padx=(6, 0))
        self.cv_compress_options = self.panel_canvas.create_window(
            15, 210, window=self.compress_options_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_compress_options, state='hidden')

        # 压缩级别说明 (y=245)
        self.compress_hint_var = tk.StringVar(
            value=COMPRESS_PRESETS['标准压缩']['description'])
        self.compress_hint_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.compress_hint_frame, textvariable=self.compress_hint_var,
                 font=("Microsoft YaHei", 8), fg="#888888").pack(anchor=tk.W)
        self.cv_compress_hint = self.panel_canvas.create_window(
            15, 245, window=self.compress_hint_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_compress_hint, state='hidden')

        # PDF提取/删页选项区 (y=210)
        self.extract_mode_var = tk.StringVar(value='提取')
        self.extract_pages_var = tk.StringVar()
        self.extract_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.extract_options_frame, text="模式:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        tk.Radiobutton(self.extract_options_frame, text="提取指定页",
                       variable=self.extract_mode_var, value='提取',
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(6, 0))
        tk.Radiobutton(self.extract_options_frame, text="删除指定页",
                       variable=self.extract_mode_var, value='删除',
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(6, 0))
        tk.Label(self.extract_options_frame, text="页码:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(10, 0))
        tk.Entry(self.extract_options_frame, textvariable=self.extract_pages_var,
                 width=18, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(4, 0))
        self.cv_extract_options = self.panel_canvas.create_window(
            15, 210, window=self.extract_options_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_extract_options, state='hidden')

        # 提取/删页说明 (y=245)
        self.extract_hint_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.extract_hint_frame, text="格式示例：1,3,5-10  支持单页、范围、混合",
                 font=("Microsoft YaHei", 8), fg="#888888").pack(anchor=tk.W)
        self.cv_extract_hint = self.panel_canvas.create_window(
            15, 245, window=self.extract_hint_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_extract_hint, state='hidden')

        # PDF转Excel选项区 (y=210)
        self.excel_strategy_var = tk.StringVar(value='自动检测')
        self.excel_merge_var = tk.BooleanVar(value=False)
        self.excel_extract_mode_var = tk.StringVar(value='结构提取')
        self.excel_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.excel_options_frame, text="提取策略:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        for strategy in TABLE_STRATEGIES:
            tk.Radiobutton(
                self.excel_options_frame, text=strategy,
                variable=self.excel_strategy_var, value=strategy,
                font=("Microsoft YaHei", 9),
                command=self._on_excel_strategy_changed,
            ).pack(side=tk.LEFT, padx=(6, 0))
        tk.Checkbutton(self.excel_options_frame, text="合并到一个Sheet",
                       variable=self.excel_merge_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(12, 0))
        self.cv_excel_options = self.panel_canvas.create_window(
            15, 210, window=self.excel_options_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_excel_options, state='hidden')

        # Excel提取方式 (y=245)
        self.excel_mode_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.excel_mode_frame, text="提取方式:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        tk.Radiobutton(
            self.excel_mode_frame, text="结构提取",
            variable=self.excel_extract_mode_var, value="结构提取",
            font=("Microsoft YaHei", 9)
        ).pack(side=tk.LEFT, padx=(6, 0))
        tk.Radiobutton(
            self.excel_mode_frame, text="OCR提取",
            variable=self.excel_extract_mode_var, value="OCR提取",
            font=("Microsoft YaHei", 9)
        ).pack(side=tk.LEFT, padx=(6, 0))
        self.cv_excel_mode = self.panel_canvas.create_window(
            15, 245, window=self.excel_mode_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_excel_mode, state='hidden')

        # Excel策略说明 (y=270)
        self.excel_hint_var = tk.StringVar(
            value=TABLE_STRATEGIES['自动检测']['description'])
        self.excel_hint_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.excel_hint_frame, textvariable=self.excel_hint_var,
                 font=("Microsoft YaHei", 8), fg="#888888").pack(anchor=tk.W)
        self.cv_excel_hint = self.panel_canvas.create_window(
            15, 270, window=self.excel_hint_frame, anchor="nw"
        )
        self.panel_canvas.itemconfigure(self.cv_excel_hint, state='hidden')

        # PDF批量文本/图片提取选项（分3行，避免固定窗口遮挡）
        self.batch_options_frame = tk.Frame(self.panel_canvas)
        tk.Checkbutton(self.batch_options_frame, text="文本",
                       variable=self.batch_text_enabled_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 8))
        tk.Checkbutton(self.batch_options_frame, text="图片",
                       variable=self.batch_image_enabled_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 12))
        tk.Label(self.batch_options_frame, text="格式:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        ttk.Combobox(
            self.batch_options_frame, textvariable=self.batch_text_format_var,
            values=["txt", "json", "csv"],
            state='readonly', width=5, font=("Microsoft YaHei", 9)
        ).pack(side=tk.LEFT, padx=(4, 10))
        tk.Label(self.batch_options_frame, text="模式:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        ttk.Combobox(
            self.batch_options_frame, textvariable=self.batch_text_mode_var,
            values=["合并为一个文件", "每页一个文件"],
            state='readonly', width=8, font=("Microsoft YaHei", 9)
        ).pack(side=tk.LEFT, padx=(4, 0))
        self.cv_batch_options = self.panel_canvas.create_window(15, 210, window=self.batch_options_frame, anchor="nw")
        self.panel_canvas.itemconfigure(self.cv_batch_options, state='hidden')

        self.batch_options2_frame = tk.Frame(self.panel_canvas)
        tk.Checkbutton(self.batch_options2_frame, text="保留换行",
                       variable=self.batch_preserve_layout_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 8))
        tk.Checkbutton(self.batch_options2_frame, text="无文本时OCR",
                       variable=self.batch_ocr_enabled_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(self.batch_options2_frame, text="页码:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        tk.Entry(self.batch_options2_frame, textvariable=self.batch_pages_var,
                 width=16, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(4, 0))
        self.cv_batch_options2 = self.panel_canvas.create_window(15, 245, window=self.batch_options2_frame, anchor="nw")
        self.panel_canvas.itemconfigure(self.cv_batch_options2, state='hidden')

        self.batch_options3_frame = tk.Frame(self.panel_canvas)
        tk.Checkbutton(self.batch_options3_frame, text="按页文件夹",
                       variable=self.batch_image_per_page_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 8))
        tk.Checkbutton(self.batch_options3_frame, text="图片去重",
                       variable=self.batch_image_dedupe_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 10))
        tk.Label(self.batch_options3_frame, text="图片格式:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        ttk.Combobox(
            self.batch_options3_frame, textvariable=self.batch_image_format_var,
            values=["原格式", "PNG", "JPEG"],
            state='readonly', width=6, font=("Microsoft YaHei", 9)
        ).pack(side=tk.LEFT, padx=(4, 10))
        tk.Checkbutton(self.batch_options3_frame, text="打包ZIP",
                       variable=self.batch_zip_enabled_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 0))
        self.cv_batch_options3 = self.panel_canvas.create_window(15, 280, window=self.batch_options3_frame, anchor="nw")
        self.panel_canvas.itemconfigure(self.cv_batch_options3, state='hidden')

        self.cv_batch_hint = self.panel_canvas.create_text(
            15, 305, text="页码格式示例：1,3,5-10 或 1，3，5-10（留空表示全部页）",
            font=("Microsoft YaHei", 8), anchor="nw", fill="#888888"
        )
        self.panel_canvas.itemconfigure(self.cv_batch_hint, state='hidden')

        # PDF批量盖章选项（分4行，避免固定窗口遮挡）
        self.stamp_options_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.stamp_options_frame, text="模式:",
                 font=("Microsoft YaHei", 9, "bold")).pack(side=tk.LEFT)
        self.stamp_mode_combo = ttk.Combobox(
            self.stamp_options_frame, textvariable=self.stamp_mode_var,
            values=["普通章", "二维码", "骑缝章", "模板"],
            state='readonly', width=7, font=("Microsoft YaHei", 9)
        )
        self.stamp_mode_combo.pack(side=tk.LEFT, padx=(6, 8))
        self.stamp_mode_combo.bind("<<ComboboxSelected>>", self._on_stamp_mode_changed)
        tk.Button(self.stamp_options_frame, text="章图(多选)...",
                  font=("Microsoft YaHei", 8), command=self._choose_stamp_image,
                  cursor='hand2').pack(side=tk.LEFT, padx=(0, 6))
        self.stamp_image_label = tk.Label(self.stamp_options_frame, text="",
                                          font=("Microsoft YaHei", 8), fg="#666")
        self.stamp_image_label.pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(self.stamp_options_frame, text="模板...",
                  font=("Microsoft YaHei", 8), command=self._choose_stamp_template,
                  cursor='hand2').pack(side=tk.LEFT, padx=(0, 6))
        self.stamp_template_label = tk.Label(self.stamp_options_frame, text="",
                                             font=("Microsoft YaHei", 8), fg="#666")
        self.stamp_template_label.pack(side=tk.LEFT)
        self.cv_stamp_options = self.panel_canvas.create_window(15, 210, window=self.stamp_options_frame, anchor="nw")
        self.panel_canvas.itemconfigure(self.cv_stamp_options, state='hidden')

        self.stamp_options2_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.stamp_options2_frame, text="二维码内容:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.stamp_qr_entry = tk.Entry(self.stamp_options2_frame, textvariable=self.stamp_qr_text_var,
                                       width=14, font=("Microsoft YaHei", 9))
        self.stamp_qr_entry.pack(side=tk.LEFT, padx=(4, 10))
        tk.Label(self.stamp_options2_frame, text="页码:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        tk.Entry(self.stamp_options2_frame, textvariable=self.stamp_pages_var,
                 width=14, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(4, 10))
        tk.Checkbutton(self.stamp_options2_frame, text="去白底",
                       variable=self.stamp_remove_white_bg_var,
                       font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(0, 0))
        self.cv_stamp_options2 = self.panel_canvas.create_window(15, 245, window=self.stamp_options2_frame, anchor="nw")
        self.panel_canvas.itemconfigure(self.cv_stamp_options2, state='hidden')

        self.stamp_options3_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.stamp_options3_frame, text="骑缝边:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.stamp_seam_side_combo = ttk.Combobox(
            self.stamp_options3_frame, textvariable=self.stamp_seam_side_var,
            values=["右侧", "左侧", "顶部", "底部"],
            state='readonly', width=5, font=("Microsoft YaHei", 9)
        )
        self.stamp_seam_side_combo.pack(side=tk.LEFT, padx=(4, 8))
        tk.Label(self.stamp_options3_frame, text="对齐:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.stamp_seam_align_combo = ttk.Combobox(
            self.stamp_options3_frame, textvariable=self.stamp_seam_align_var,
            values=["居中", "顶部", "底部"],
            state='readonly', width=5, font=("Microsoft YaHei", 9)
        )
        self.stamp_seam_align_combo.pack(side=tk.LEFT, padx=(4, 8))
        tk.Label(self.stamp_options3_frame, text="压边比例:",
                 font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        self.stamp_seam_overlap_entry = tk.Entry(
            self.stamp_options3_frame, textvariable=self.stamp_seam_overlap_var,
            width=6, font=("Microsoft YaHei", 9)
        )
        self.stamp_seam_overlap_entry.pack(side=tk.LEFT, padx=(4, 0))
        self.cv_stamp_options3 = self.panel_canvas.create_window(15, 280, window=self.stamp_options3_frame, anchor="nw")
        self.panel_canvas.itemconfigure(self.cv_stamp_options3, state='hidden')

        self.stamp_options4_frame = tk.Frame(self.panel_canvas)
        self.stamp_preview_btn = tk.Button(
            self.stamp_options4_frame, text="预览设置...",
            font=("Microsoft YaHei", 8), command=self._open_stamp_preview,
            cursor='hand2'
        )
        self.stamp_preview_btn.pack(side=tk.LEFT, padx=(0, 8))
        self.stamp_export_template_btn = tk.Button(
            self.stamp_options4_frame, text="导出模板...",
            font=("Microsoft YaHei", 8), command=self._export_stamp_template_from_current,
            cursor='hand2'
        )
        self.stamp_export_template_btn.pack(side=tk.LEFT, padx=(0, 8))
        tk.Label(self.stamp_options4_frame, textvariable=self.stamp_preview_info_var,
                 font=("Microsoft YaHei", 8), fg="#666").pack(side=tk.LEFT)
        self.cv_stamp_options4 = self.panel_canvas.create_window(15, 315, window=self.stamp_options4_frame, anchor="nw")
        self.panel_canvas.itemconfigure(self.cv_stamp_options4, state='hidden')

        self.stamp_hint_var = tk.StringVar(value="")
        self.stamp_hint_frame = tk.Frame(self.panel_canvas)
        tk.Label(self.stamp_hint_frame, textvariable=self.stamp_hint_var,
                 font=("Microsoft YaHei", 8), fg="#888888").pack(anchor=tk.W)
        self.cv_stamp_hint = self.panel_canvas.create_window(15, 340, window=self.stamp_hint_frame, anchor="nw")
        self.panel_canvas.itemconfigure(self.cv_stamp_hint, state='hidden')

        # API状态提示
        self.cv_api_hint = self.panel_canvas.create_text(
            15, 270, text="", font=("Microsoft YaHei", 8), anchor="nw", fill="#888888"
        )

        # 进度条
        self.progress_bar = ttk.Progressbar(self.panel_canvas, mode='determinate')
        self.cv_progress_bar = self.panel_canvas.create_window(
            20, self.progress_y, window=self.progress_bar, anchor="nw", width=1, height=25
        )

        # 进度文本
        self.cv_progress_text = self.panel_canvas.create_text(
            0, self.progress_text_y, text="", font=("Microsoft YaHei", 9), anchor="n"
        )

        # 按钮
        btn_frame = tk.Frame(self.panel_canvas)
        self.convert_btn = tk.Button(
            btn_frame, text="开始转换", command=self.start_conversion,
            font=("Microsoft YaHei", 12, "bold"), padx=40, pady=12, cursor='hand2'
        )
        self.convert_btn.pack(side=tk.LEFT, expand=True, padx=5)
        tk.Button(
            btn_frame, text="清除", command=self.clear_selection,
            font=("Microsoft YaHei", 12), padx=40, pady=12, cursor='hand2'
        ).pack(side=tk.LEFT, expand=True, padx=5)
        self.cv_btn_frame = self.panel_canvas.create_window(
            0, self.btn_y, window=btn_frame, anchor="n"
        )

        # 拖拽提示
        dnd_text = "支持拖拽文件到窗口" if WINDND_AVAILABLE else ""
        self.cv_dnd_hint = self.panel_canvas.create_text(
            0, self.dnd_y, text=dnd_text, font=("Microsoft YaHei", 8),
            anchor="n", fill="#aaaaaa"
        )

        # 状态栏
        self.cv_status_text = self.panel_canvas.create_text(
            15, 0, text=self.status_message.get(),
            font=("Microsoft YaHei", 9), anchor="sw"
        )
        self.status_message.trace_add("write", self._on_status_var_changed)
        self._update_stamp_preview_info()

        # 事件绑定
        self.root.bind("<Configure>", self.on_root_resize)
        self.panel_canvas.bind("<Configure>", self.on_panel_resize)
        self.root.after(50, self.refresh_layout)

    # ==========================================================
    # Canvas文字/布局
    # ==========================================================

    def _on_title_var_changed(self, *args):
        if self.panel_canvas:
            self.panel_canvas.itemconfigure(self.cv_title, text=self.title_text_var.get())

    def _on_status_var_changed(self, *args):
        if self.panel_canvas:
            self.panel_canvas.itemconfigure(self.cv_status_text, text=self.status_message.get())

    def set_progress_text(self, text):
        if self.panel_canvas:
            self.panel_canvas.itemconfigure(self.cv_progress_text, text=text)

    def layout_canvas(self):
        """根据Canvas尺寸重新布局所有元素"""
        w = self.panel_canvas.winfo_width()
        h = self.panel_canvas.winfo_height()
        if w <= 1 or h <= 1:
            return
        cx = w // 2
        self.panel_canvas.coords(self.cv_title, cx, 35)
        self.panel_canvas.coords(self.cv_subtitle, cx, 75)
        self.panel_canvas.coords(self.cv_section1, 15, 105)
        self.panel_canvas.coords(self.cv_file_frame, 15, 130)
        self.panel_canvas.itemconfigure(self.cv_file_frame, width=w - 30)
        self.panel_canvas.coords(self.cv_section2, 15, 185)
        self.panel_canvas.coords(self.cv_range_frame, 15, 210)
        self.panel_canvas.coords(self.cv_formula_frame, 15, 245)
        self.panel_canvas.coords(self.cv_image_options, 15, 245)
        self.panel_canvas.coords(self.cv_merge_info, 15, 210)
        self.panel_canvas.coords(self.cv_split_options, 15, 210)
        self.panel_canvas.coords(self.cv_img2pdf_options, 15, 210)
        self.panel_canvas.coords(self.cv_watermark_options, 15, 210)
        self.panel_canvas.coords(self.cv_watermark_detail, 15, 245)
        self.panel_canvas.coords(self.cv_encrypt_options, 15, 210)
        self.panel_canvas.coords(self.cv_encrypt_perm, 15, 245)
        self.panel_canvas.coords(self.cv_compress_options, 15, 210)
        self.panel_canvas.coords(self.cv_compress_hint, 15, 245)
        self.panel_canvas.coords(self.cv_extract_options, 15, 210)
        self.panel_canvas.coords(self.cv_extract_hint, 15, 245)
        self.panel_canvas.coords(self.cv_excel_options, 15, 210)
        self.panel_canvas.coords(self.cv_excel_mode, 15, 245)
        self.panel_canvas.coords(self.cv_excel_hint, 15, 270)
        self.panel_canvas.coords(self.cv_batch_options, 15, 210)
        self.panel_canvas.coords(self.cv_batch_options2, 15, 245)
        self.panel_canvas.coords(self.cv_batch_options3, 15, 280)
        self.panel_canvas.coords(self.cv_batch_hint, 15, 305)
        self.panel_canvas.coords(self.cv_stamp_options, 15, 210)
        self.panel_canvas.coords(self.cv_stamp_options2, 15, 245)
        self.panel_canvas.coords(self.cv_stamp_options3, 15, 280)
        self.panel_canvas.coords(self.cv_stamp_options4, 15, 315)
        self.panel_canvas.coords(self.cv_stamp_hint, 15, 340)
        self.panel_canvas.coords(self.cv_api_hint, 15, 270)
        self.panel_canvas.coords(self.cv_progress_bar, 20, self.progress_y)
        self.panel_canvas.itemconfigure(self.cv_progress_bar, width=w - 40)
        self.panel_canvas.coords(self.cv_progress_text, cx, self.progress_text_y)
        self.panel_canvas.coords(self.cv_btn_frame, cx, self.btn_y)
        self.panel_canvas.coords(self.cv_dnd_hint, cx, self.dnd_y)
        self.panel_canvas.coords(self.cv_status_text, 15, h - 10)

    # ==========================================================
    # 功能切换 / 选项变化
    # ==========================================================

    def _on_function_changed(self, event=None):
        func = self.current_function_var.get()
        self.progress_y = 290
        self.progress_text_y = 325
        self.btn_y = 370
        self.dnd_y = 410

        # 先隐藏所有可选区域
        for cv_item in [self.cv_formula_frame, self.cv_api_hint,
                        self.cv_image_options, self.cv_merge_info,
                        self.cv_split_options, self.cv_img2pdf_options,
                        self.cv_watermark_options, self.cv_watermark_detail,
                        self.cv_encrypt_options, self.cv_encrypt_perm,
                        self.cv_compress_options, self.cv_compress_hint,
                        self.cv_extract_options, self.cv_extract_hint,
                        self.cv_excel_options, self.cv_excel_mode, self.cv_excel_hint,
                        self.cv_batch_options, self.cv_batch_options2, self.cv_batch_options3, self.cv_batch_hint,
                        self.cv_stamp_options, self.cv_stamp_options2, self.cv_stamp_options3, self.cv_stamp_options4, self.cv_stamp_hint]:
            self.panel_canvas.itemconfigure(cv_item, state='hidden')

        title_prefix = self.title_text_var.get().split(' - ')[0] if ' - ' in self.title_text_var.get() else self.title_text_var.get()

        if func == "PDF转Word":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='normal')
            self.panel_canvas.itemconfigure(self.cv_formula_frame, state='normal')
            self.panel_canvas.itemconfigure(self.cv_api_hint, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件（可多选）")
            self.panel_canvas.itemconfigure(self.cv_section2, text="页范围（可选）")
            self.root.title(f"{title_prefix} - PDF转Word")

        elif func == "PDF转图片":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='normal')
            self.panel_canvas.itemconfigure(self.cv_image_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件（可多选）")
            self.panel_canvas.itemconfigure(self.cv_section2, text="页范围（可选）")
            self.root.title(f"{title_prefix} - PDF转图片")

        elif func == "PDF合并":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_merge_info, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件（至少2个）")
            self.panel_canvas.itemconfigure(self.cv_section2, text="文件信息")
            self.merge_info_label.config(text="请选择至少2个PDF文件，将按选择顺序合并")
            self.root.title(f"{title_prefix} - PDF合并")

        elif func == "PDF拆分":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_split_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件")
            self.panel_canvas.itemconfigure(self.cv_section2, text="拆分选项")
            self._on_split_mode_changed()
            self.root.title(f"{title_prefix} - PDF拆分")

        elif func == "图片转PDF":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_img2pdf_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择图片文件（可多选）")
            self.panel_canvas.itemconfigure(self.cv_section2, text="输出选项")
            self.root.title(f"{title_prefix} - 图片转PDF")

        elif func == "PDF加水印":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_watermark_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_watermark_detail, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件")
            self.panel_canvas.itemconfigure(self.cv_section2, text="水印选项")
            self.root.title(f"{title_prefix} - PDF加水印")

        elif func == "PDF加密/解密":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_encrypt_options, state='normal')
            self._on_encrypt_mode_changed()
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件")
            self.panel_canvas.itemconfigure(self.cv_section2, text="加密/解密选项")
            self.root.title(f"{title_prefix} - PDF加密/解密")

        elif func == "PDF压缩":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_compress_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_compress_hint, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件")
            self.panel_canvas.itemconfigure(self.cv_section2, text="压缩选项")
            self.root.title(f"{title_prefix} - PDF压缩")

        elif func == "PDF提取/删页":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_extract_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_extract_hint, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件")
            self.panel_canvas.itemconfigure(self.cv_section2, text="提取/删页选项")
            self.root.title(f"{title_prefix} - PDF提取/删页")

        elif func == "OCR可搜索PDF":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='normal')
            self.panel_canvas.itemconfigure(self.cv_api_hint, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择扫描版PDF文件")
            self.panel_canvas.itemconfigure(self.cv_section2, text="页范围（可选）")
            self.root.title(f"{title_prefix} - OCR可搜索PDF")

        elif func == "PDF转Excel":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='normal')
            self.panel_canvas.itemconfigure(self.cv_excel_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_excel_mode, state='normal')
            self.panel_canvas.itemconfigure(self.cv_excel_hint, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择包含表格的PDF文件")
            self.panel_canvas.itemconfigure(self.cv_section2, text="页范围（可选）")
            self.root.title(f"{title_prefix} - PDF转Excel")

        if func == "PDF批量文本/图片提取":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_batch_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_batch_options2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_batch_options3, state='normal')
            self.panel_canvas.itemconfigure(self.cv_batch_hint, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件（可多选）")
            self.panel_canvas.itemconfigure(self.cv_section2, text="批量提取选项")
            self.progress_y = 335
            self.progress_text_y = 370
            self.btn_y = 415
            self.dnd_y = 455
            self.root.title(f"{title_prefix} - PDF批量文本/图片提取")

        if func == "PDF批量盖章":
            self.panel_canvas.itemconfigure(self.cv_section2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_range_frame, state='hidden')
            self.panel_canvas.itemconfigure(self.cv_stamp_options, state='normal')
            self.panel_canvas.itemconfigure(self.cv_stamp_options2, state='normal')
            self.panel_canvas.itemconfigure(self.cv_stamp_options3, state='normal')
            self.panel_canvas.itemconfigure(self.cv_stamp_options4, state='normal')
            self.panel_canvas.itemconfigure(self.cv_stamp_hint, state='normal')
            self.panel_canvas.itemconfigure(self.cv_section1, text="选择PDF文件（可多选）")
            self.panel_canvas.itemconfigure(self.cv_section2, text="批量盖章选项")
            self.progress_y = 370
            self.progress_text_y = 405
            self.btn_y = 450
            self.dnd_y = 490
            self._on_stamp_mode_changed()
            self.root.title(f"{title_prefix} - PDF批量盖章")

        self.layout_canvas()

        self.selected_file.set("")
        self.selected_files_list = []
        self._update_order_btn()
        self.status_message.set("就绪")
        self.save_settings()

    def _on_option_changed(self):
        self._update_api_hint()
        self.save_settings()

    def _default_stamp_profile(self):
        base = self.stamp_preview_profile or {}
        return {
            "enabled": True,
            "x_ratio": self._clamp_value(base.get("x_ratio", 0.85), 0.0, 1.0, 0.85),
            "y_ratio": self._clamp_value(base.get("y_ratio", 0.85), 0.0, 1.0, 0.85),
            "size_ratio": self._clamp_value(base.get("size_ratio", 0.18), 0.03, 0.7, 0.18),
            "opacity": self._clamp_value(base.get("opacity", 0.85), 0.05, 1.0, 0.85),
        }

    def _normalize_stamp_profile(self, profile):
        data = dict(profile or {})
        default = self._default_stamp_profile()
        return {
            "enabled": bool(data.get("enabled", default["enabled"])),
            "x_ratio": self._clamp_value(data.get("x_ratio", default["x_ratio"]), 0.0, 1.0, default["x_ratio"]),
            "y_ratio": self._clamp_value(data.get("y_ratio", default["y_ratio"]), 0.0, 1.0, default["y_ratio"]),
            "size_ratio": self._clamp_value(data.get("size_ratio", default["size_ratio"]), 0.03, 0.7, default["size_ratio"]),
            "opacity": self._clamp_value(data.get("opacity", default["opacity"]), 0.05, 1.0, default["opacity"]),
        }

    def _set_stamp_images(self, paths, selected_idx=None):
        cleaned = []
        seen = set()
        for p in paths or []:
            if not p:
                continue
            full = os.path.abspath(str(p))
            if not os.path.exists(full):
                continue
            ext = os.path.splitext(full)[1].lower()
            if ext not in (".png", ".jpg", ".jpeg", ".bmp"):
                continue
            if full in seen:
                continue
            seen.add(full)
            cleaned.append(full)

        old_profiles = dict(self.stamp_profiles or {})
        self.stamp_profiles = {}
        for full in cleaned:
            self.stamp_profiles[full] = self._normalize_stamp_profile(old_profiles.get(full))

        self.stamp_image_paths = cleaned
        if not cleaned:
            self.stamp_selected_image_idx = 0
            self.stamp_image_path = ""
            self._update_stamp_image_label()
            return

        if selected_idx is None:
            prev = self.stamp_image_path if self.stamp_image_path in cleaned else None
            if prev:
                idx = cleaned.index(prev)
            else:
                idx = min(max(int(self.stamp_selected_image_idx or 0), 0), len(cleaned) - 1)
        else:
            idx = min(max(int(selected_idx), 0), len(cleaned) - 1)

        self.stamp_selected_image_idx = idx
        self.stamp_image_path = cleaned[idx]
        self._update_stamp_image_label()

    def _get_stamp_profile_for_path(self, image_path):
        if not image_path:
            return self._default_stamp_profile()
        existing = self.stamp_profiles.get(image_path)
        if existing is None:
            existing = self._default_stamp_profile()
            self.stamp_profiles[image_path] = existing
        else:
            existing = self._normalize_stamp_profile(existing)
            self.stamp_profiles[image_path] = existing
        return existing

    def _get_enabled_stamp_profiles(self):
        profiles = []
        for path in self.stamp_image_paths or []:
            if not path or not os.path.exists(path):
                continue
            prof = self._get_stamp_profile_for_path(path)
            if not prof.get("enabled", True):
                continue
            profiles.append({
                "image_path": path,
                "enabled": True,
                "x_ratio": prof["x_ratio"],
                "y_ratio": prof["y_ratio"],
                "size_ratio": prof["size_ratio"],
                "opacity": prof["opacity"],
            })
        return profiles

    def _update_stamp_image_label(self):
        count = len(self.stamp_image_paths or [])
        if count <= 0:
            self.stamp_image_label.config(text="")
            return
        active = self._get_active_stamp_image_path()
        if not active:
            self.stamp_image_label.config(text=f"{count}个章图")
            return
        name = os.path.basename(active)
        short = name if len(name) <= 12 else name[:9] + "..."
        if count == 1:
            self.stamp_image_label.config(text=short)
        else:
            self.stamp_image_label.config(text=f"{count}个章图 | {short}")

    def _get_active_stamp_image_path(self):
        if self.stamp_image_paths:
            valid = [p for p in self.stamp_image_paths if os.path.exists(p)]
            if valid != self.stamp_image_paths:
                self.stamp_image_paths = valid
            if self.stamp_image_paths:
                idx = min(max(int(self.stamp_selected_image_idx or 0), 0), len(self.stamp_image_paths) - 1)
                self.stamp_selected_image_idx = idx
                self.stamp_image_path = self.stamp_image_paths[idx]
                return self.stamp_image_path
        if self.stamp_image_path and os.path.exists(self.stamp_image_path):
            self.stamp_image_paths = [self.stamp_image_path]
            self.stamp_selected_image_idx = 0
            return self.stamp_image_path
        self.stamp_image_path = ""
        return ""

    def _choose_stamp_image(self):
        filenames = filedialog.askopenfilenames(
            title="选择章图（可多选）",
            filetypes=[("图片文件", "*.png;*.jpg;*.jpeg;*.bmp"), ("所有文件", "*.*")]
        )
        if filenames:
            self._set_stamp_images(list(filenames), selected_idx=0)
            self._update_stamp_preview_info()
            self.save_settings()

    def _choose_stamp_template(self):
        filename = filedialog.askopenfilename(
            title="选择模板JSON",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if filename:
            self.stamp_template_path = filename
            name = os.path.basename(filename)
            self.stamp_template_label.config(text=name if len(name) <= 16 else name[:13] + "...")
            self._update_stamp_preview_info()
            self.save_settings()

    def _parse_template_pages_scope(self, pages_text):
        parsed = PDFBatchStampConverter._parse_pages_str((pages_text or "").strip())
        if parsed is None:
            return None
        if not parsed:
            return []
        return [p + 1 for p in sorted(set(parsed))]

    def _build_stamp_template_data(
        self,
        mode_key,
        profiles=None,
        single_profile=None,
        qr_text="",
        pages_text="",
    ):
        scope = self._parse_template_pages_scope(pages_text)
        if scope is None:
            raise ValueError("页码格式不正确，请使用 1,3,5-10 这种格式。")

        template_data = {
            "version": 1,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "mode": mode_key,
            "remove_white_bg": bool(self.stamp_remove_white_bg_var.get()),
            "elements": [],
        }

        def with_scope(elem):
            if scope:
                elem["pages"] = list(scope)
            return elem

        if mode_key == "template":
            if not (self.stamp_template_path and os.path.exists(self.stamp_template_path)):
                raise ValueError("当前未选择模板 JSON。")
            with open(self.stamp_template_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                loaded_data = dict(loaded)
                loaded_data.setdefault("version", 1)
                loaded_data.setdefault("created_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                loaded_data["exported_from"] = os.path.abspath(self.stamp_template_path)
                return loaded_data
            if isinstance(loaded, list):
                return {
                    "version": 1,
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "mode": "template",
                    "elements": loaded,
                }
            raise ValueError("模板 JSON 结构无效。")

        if mode_key in ("seal", "seam"):
            items = []
            for item in profiles or []:
                if not isinstance(item, dict):
                    continue
                image_path = os.path.abspath(str(item.get("image_path", "")).strip())
                if not image_path or not os.path.exists(image_path):
                    continue
                prof = self._normalize_stamp_profile(item)
                items.append({
                    "image_path": image_path,
                    "x_ratio": prof["x_ratio"],
                    "y_ratio": prof["y_ratio"],
                    "size_ratio": prof["size_ratio"],
                    "opacity": prof["opacity"],
                })
            if not items:
                raise ValueError("没有可导出的章图，请先勾选可用章图。")

            if mode_key == "seal":
                for it in items:
                    template_data["elements"].append(with_scope({
                        "type": "seal",
                        "image_path": it["image_path"],
                        "x_ratio": it["x_ratio"],
                        "y_ratio": it["y_ratio"],
                        "w_ratio": it["size_ratio"],
                        "opacity": it["opacity"],
                    }))
            else:
                seam_side_map = {"右侧": "right", "左侧": "left", "顶部": "top", "底部": "bottom"}
                seam_align_map = {"居中": "center", "顶部": "top", "底部": "bottom"}
                overlap = self._clamp_value(self.stamp_seam_overlap_var.get(), 0.05, 0.95, 0.25)
                side = seam_side_map.get(self.stamp_seam_side_var.get(), "right")
                align = seam_align_map.get(self.stamp_seam_align_var.get(), "center")
                for it in items:
                    template_data["elements"].append(with_scope({
                        "type": "seam",
                        "image_path": it["image_path"],
                        "size_ratio": it["size_ratio"],
                        "opacity": it["opacity"],
                        "side": side,
                        "align": align,
                        "overlap_ratio": overlap,
                    }))
            return template_data

        if mode_key == "qr":
            txt = (qr_text or "").strip()
            if not txt:
                raise ValueError("二维码内容为空，无法导出模板。")
            prof = self._normalize_stamp_profile(single_profile or self.stamp_preview_profile)
            template_data["elements"].append(with_scope({
                "type": "qr",
                "text": txt,
                "x_ratio": prof["x_ratio"],
                "y_ratio": prof["y_ratio"],
                "w_ratio": prof["size_ratio"],
                "opacity": prof["opacity"],
            }))
            return template_data

        raise ValueError(f"不支持的模式：{mode_key}")

    def _export_stamp_template(
        self,
        mode_key,
        profiles=None,
        single_profile=None,
        qr_text="",
        pages_text="",
        parent=None,
    ):
        try:
            template_data = self._build_stamp_template_data(
                mode_key=mode_key,
                profiles=profiles,
                single_profile=single_profile,
                qr_text=qr_text,
                pages_text=pages_text,
            )
        except ValueError as exc:
            messagebox.showwarning("导出模板", str(exc), parent=parent)
            return ""
        except Exception as exc:
            messagebox.showerror("导出失败", f"生成模板数据失败：\n{exc}", parent=parent)
            return ""

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"stamp_template_{mode_key}_{ts}.json"
        filename = filedialog.asksaveasfilename(
            title="导出模板JSON",
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
            parent=parent,
        )
        if not filename:
            return ""
        try:
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(template_data, f, ensure_ascii=False, indent=2)
            self.status_message.set("模板JSON已导出")
            return filename
        except Exception as exc:
            messagebox.showerror("导出失败", f"写入模板文件失败：\n{exc}", parent=parent)
            return ""

    def _export_stamp_template_from_current(self):
        mode_key = self._get_stamp_mode_key()
        profiles = self._get_enabled_stamp_profiles() if mode_key in ("seal", "seam") else None
        single_profile = self.stamp_preview_profile if mode_key in ("qr", "template") else None
        path = self._export_stamp_template(
            mode_key=mode_key,
            profiles=profiles,
            single_profile=single_profile,
            qr_text=self.stamp_qr_text_var.get().strip(),
            pages_text=self.stamp_pages_var.get().strip(),
            parent=self.root,
        )
        if path:
            messagebox.showinfo("导出成功", f"模板已导出：\n{path}")

    @staticmethod
    def _clamp_value(value, min_value, max_value, default):
        try:
            numeric = float(value)
        except Exception:
            numeric = float(default)
        if numeric < min_value:
            return min_value
        if numeric > max_value:
            return max_value
        return numeric

    def _get_stamp_mode_key(self):
        mode_map = {
            "普通章": "seal",
            "二维码": "qr",
            "骑缝章": "seam",
            "模板": "template",
        }
        return mode_map.get(self.stamp_mode_var.get(), "seal")

    def _update_stamp_preview_info(self):
        mode_key = self._get_stamp_mode_key()
        profile = self.stamp_preview_profile or {}
        x_ratio = self._clamp_value(profile.get("x_ratio", 0.85), 0.0, 1.0, 0.85)
        y_ratio = self._clamp_value(profile.get("y_ratio", 0.85), 0.0, 1.0, 0.85)
        size_ratio = self._clamp_value(profile.get("size_ratio", 0.18), 0.03, 0.7, 0.18)
        opacity = self._clamp_value(profile.get("opacity", 0.85), 0.05, 1.0, 0.85)
        active_path = self._get_active_stamp_image_path()
        image_name = os.path.basename(active_path) if active_path else ""
        image_suffix = f" | {image_name}" if image_name else ""
        enabled_count = len(self._get_enabled_stamp_profiles()) if mode_key in ("seal", "seam") else 0
        enabled_suffix = f" | 已启用{enabled_count}" if enabled_count > 0 else ""
        if mode_key == "seam":
            self.stamp_preview_info_var.set(
                f"骑缝预览 透明度 {opacity:.2f} 尺寸 {size_ratio:.2f}{image_suffix}{enabled_suffix}"
            )
        elif mode_key == "template":
            self.stamp_preview_info_var.set(f"模板预览 透明度 {opacity:.2f}")
        else:
            self.stamp_preview_info_var.set(
                f"位置({x_ratio:.2f},{y_ratio:.2f}) 透明度 {opacity:.2f} 尺寸 {size_ratio:.2f}{image_suffix}{enabled_suffix}"
            )

    def _resolve_preview_pdf(self):
        if self.selected_files_list:
            first = self.selected_files_list[0]
            if os.path.exists(first) and first.lower().endswith(".pdf"):
                return first
        selected = (self.selected_file.get() or "").strip()
        if selected and os.path.exists(selected) and selected.lower().endswith(".pdf"):
            return selected
        return None

    def _build_template_preview_image(self, opacity):
        if not self.stamp_template_path or not os.path.exists(self.stamp_template_path):
            return None
        try:
            with open(self.stamp_template_path, "r", encoding="utf-8") as f:
                template = json.load(f)
        except Exception:
            return None

        elements = template.get("elements", [])
        for elem in elements:
            elem_type = str(elem.get("type", "")).strip().lower()
            if elem_type == "seal":
                image_path = str(elem.get("image_path", "")).strip()
                if image_path and os.path.exists(image_path):
                    image = Image.open(image_path).convert("RGBA")
                    if self.stamp_remove_white_bg_var.get():
                        image = PDFBatchStampConverter._remove_white_background(image)
                    return PDFBatchStampConverter._apply_alpha(image, opacity)
            elif elem_type == "qr":
                text = str(elem.get("text", "")).strip()
                if text:
                    try:
                        qr_bytes = PDFBatchStampConverter._make_qr_png_bytes(
                            text,
                            opacity=opacity,
                            remove_white_bg=bool(self.stamp_remove_white_bg_var.get()),
                        )
                        return Image.open(io.BytesIO(qr_bytes)).convert("RGBA")
                    except Exception:
                        return None
            elif elem_type == "text":
                text = str(elem.get("text", "")).strip()
                if text:
                    image = Image.new("RGBA", (520, 120), (255, 255, 255, 0))
                    draw = ImageDraw.Draw(image)
                    draw.text((10, 40), text, fill=(220, 0, 0, 255))
                    return PDFBatchStampConverter._apply_alpha(image, opacity)
        return None

    def _open_stamp_preview(self):
        if not PIL_AVAILABLE:
            messagebox.showwarning("提示", "预览需要 Pillow 依赖。")
            return
        if not FITZ_UI_AVAILABLE:
            messagebox.showwarning("提示", "预览需要 PyMuPDF 依赖。")
            return

        mode_key = self._get_stamp_mode_key()
        active_stamp_path = self._get_active_stamp_image_path()
        if mode_key in ("seal", "seam") and not active_stamp_path:
            messagebox.showwarning("提示", "请先选择章图。")
            return
        if mode_key == "qr" and not self.stamp_qr_text_var.get().strip():
            messagebox.showwarning("提示", "请先填写二维码内容。")
            return
        if mode_key == "template" and not (self.stamp_template_path and os.path.exists(self.stamp_template_path)):
            messagebox.showwarning("提示", "请先选择模板 JSON。")
            return

        source_pdf = self._resolve_preview_pdf()
        if not source_pdf:
            messagebox.showwarning("提示", "请先选择至少一个 PDF 文件，再打开预览。")
            return

        try:
            doc = fitz.open(source_pdf)
            if len(doc) == 0:
                doc.close()
                messagebox.showwarning("提示", "该PDF没有可预览的页面。")
                return
            first_page = doc[0]
            page_count = len(doc)
            pix = first_page.get_pixmap(matrix=fitz.Matrix(1.1, 1.1), alpha=False)
            page_image = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
            doc.close()
        except Exception as exc:
            messagebox.showerror("预览失败", f"读取PDF预览失败：\n{exc}")
            return

        max_w, max_h = 760, 500
        scale = min(max_w / page_image.width, max_h / page_image.height, 1.0)
        disp_w = max(1, int(page_image.width * scale))
        disp_h = max(1, int(page_image.height * scale))
        page_display = page_image.resize((disp_w, disp_h), Image.LANCZOS) if scale < 0.999 else page_image.copy()
        pad = 24

        preview_paths = []
        if mode_key in ("seal", "seam"):
            preview_paths = [p for p in (self.stamp_image_paths or []) if p and os.path.exists(p)]
            if active_stamp_path and active_stamp_path not in preview_paths:
                preview_paths.append(active_stamp_path)
            if not preview_paths:
                messagebox.showwarning("提示", "没有可用的章图。")
                return

        preview_profiles = {}
        if mode_key in ("seal", "seam"):
            for p in preview_paths:
                preview_profiles[p] = self._normalize_stamp_profile(self._get_stamp_profile_for_path(p))
        else:
            preview_profiles["__single__"] = self._normalize_stamp_profile(self.stamp_preview_profile)

        preview_win = tk.Toplevel(self.root)
        preview_win.title("盖章预览")
        win_w = max(960, disp_w + 160)
        win_h = max(760, disp_h + 300)
        screen_w = max(1000, self.root.winfo_screenwidth())
        screen_h = max(800, self.root.winfo_screenheight())
        win_w = min(win_w, screen_w - 80)
        win_h = min(win_h, screen_h - 80)
        preview_win.geometry(f"{int(win_w)}x{int(win_h)}")
        preview_win.minsize(920, 700)
        preview_win.resizable(False, False)
        preview_win.transient(self.root)
        preview_win.grab_set()

        info_text = f"预览文件：{os.path.basename(source_pdf)}  第1页 / 共{page_count}页"
        tk.Label(preview_win, text=info_text, font=("Microsoft YaHei", 9), fg="#666").pack(anchor="w", padx=12, pady=(10, 4))

        control_frame = tk.Frame(preview_win)
        control_frame.pack(fill=tk.X, padx=12, pady=(0, 6))
        tk.Label(control_frame, text="透明度:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        opacity_var = tk.DoubleVar(value=85)
        opacity_scale = tk.Scale(control_frame, from_=5, to=100, orient=tk.HORIZONTAL, resolution=1, showvalue=True, variable=opacity_var, length=220)
        opacity_scale.pack(side=tk.LEFT, padx=(6, 12))
        tk.Label(control_frame, text="缩放:", font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
        size_var = tk.DoubleVar(value=18)
        size_scale = tk.Scale(control_frame, from_=3, to=70, orient=tk.HORIZONTAL, resolution=1, showvalue=True, variable=size_var, length=200)
        size_scale.pack(side=tk.LEFT, padx=(6, 12))
        if mode_key == "seam":
            size_scale.config(state=tk.DISABLED)
        tk.Label(control_frame, text="勾选=参与输出，单击图章=当前编辑", font=("Microsoft YaHei", 9), fg="#666").pack(side=tk.LEFT)

        active_path_var = tk.StringVar(value=(active_stamp_path if active_stamp_path in preview_paths else (preview_paths[0] if preview_paths else "")))
        enabled_vars = {}
        if mode_key in ("seal", "seam"):
            list_frame = tk.LabelFrame(preview_win, text="章图列表", font=("Microsoft YaHei", 9))
            list_frame.pack(fill=tk.X, padx=12, pady=(0, 6))
            for p in preview_paths:
                row = tk.Frame(list_frame)
                row.pack(fill=tk.X, pady=1)
                ev = tk.BooleanVar(value=bool(preview_profiles[p].get("enabled", True)))
                enabled_vars[p] = ev
                tk.Checkbutton(row, variable=ev, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT)
                tk.Radiobutton(row, variable=active_path_var, value=p, font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(2, 4))
                tk.Label(row, text=os.path.basename(p), font=("Microsoft YaHei", 9), anchor="w").pack(side=tk.LEFT)

        canvas_frame = tk.Frame(preview_win, bg="#f5f5f5")
        canvas_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=12, pady=8)
        canvas = tk.Canvas(
            canvas_frame,
            width=disp_w + pad * 2,
            height=disp_h + pad * 2,
            bg="#f5f5f5",
            highlightthickness=1,
            highlightbackground="#cccccc",
        )
        canvas.pack(fill=tk.BOTH, expand=True)

        page_tk = ImageTk.PhotoImage(page_display)
        canvas.create_image(pad, pad, anchor="nw", image=page_tk)
        canvas.page_tk = page_tk
        canvas.create_rectangle(pad, pad, pad + disp_w, pad + disp_h, outline="#bbbbbb")

        state = {
            "render_job": None,
            "suspend_slider": False,
            "drag_path": None,
            "drag_offset_x": 0.0,
            "drag_offset_y": 0.0,
            "stamp_items": {},
            "selection_rect": canvas.create_rectangle(0, 0, 0, 0, outline="#1e88e5", width=2, dash=(4, 2), state="hidden"),
        }
        image_cache = {}
        render_cache = {}

        def get_active_key():
            if mode_key in ("seal", "seam"):
                key = active_path_var.get()
                if key not in preview_paths and preview_paths:
                    key = preview_paths[0]
                    active_path_var.set(key)
                return key
            return "__single__"

        def get_profile(key):
            if key not in preview_profiles:
                preview_profiles[key] = self._default_stamp_profile()
            return preview_profiles[key]

        def sync_sliders_from_active():
            key = get_active_key()
            profile = get_profile(key)
            state["suspend_slider"] = True
            opacity_var.set(profile["opacity"] * 100.0)
            size_var.set(profile["size_ratio"] * 100.0)
            state["suspend_slider"] = False

        def get_base_image(path):
            cache_key = (path, bool(self.stamp_remove_white_bg_var.get()))
            if cache_key in image_cache:
                return image_cache[cache_key]
            img = Image.open(path).convert("RGBA")
            if self.stamp_remove_white_bg_var.get():
                img = PDFBatchStampConverter._remove_white_background(img)
            image_cache[cache_key] = img
            return img

        def get_render_image(path, profile, mode):
            op_key = int(self._clamp_value(profile["opacity"], 0.05, 1.0, 0.85) * 1000)
            size_key = int(self._clamp_value(profile["size_ratio"], 0.03, 0.7, 0.18) * 1000)
            side_key = self.stamp_seam_side_var.get()
            align_key = self.stamp_seam_align_var.get()
            overlap_key = int(self._clamp_value(self.stamp_seam_overlap_var.get(), 0.05, 0.95, 0.25) * 1000)
            cache_key = (path, mode, op_key, size_key, side_key, align_key, overlap_key, page_count, disp_w, disp_h, bool(self.stamp_remove_white_bg_var.get()))
            if cache_key in render_cache:
                return render_cache[cache_key]

            if mode == "seal":
                base = get_base_image(path).copy()
                base = PDFBatchStampConverter._apply_alpha(base, profile["opacity"])
                tw = max(16, int(disp_w * profile["size_ratio"]))
                th = max(16, int(tw * base.height / max(1, base.width)))
                out = base.resize((tw, th), Image.LANCZOS)
                render_cache[cache_key] = out
                return out

            if mode == "seam":
                base = get_base_image(path).copy()
                base = PDFBatchStampConverter._apply_alpha(base, profile["opacity"])
                side = {"右侧": "right", "左侧": "left", "顶部": "top", "底部": "bottom"}.get(self.stamp_seam_side_var.get(), "right")
                n_pages = max(1, page_count)
                if side in ("left", "right"):
                    step = base.width / n_pages
                    x1, x2 = 0, max(1, int(round(step)))
                    piece = base.crop((x1, 0, x2, base.height))
                    base_h = max(10, int(disp_h / n_pages))
                    sr = self._clamp_value(profile["size_ratio"] / 0.18, 0.6, 2.2, 1.0)
                    th = max(10, int(base_h * sr))
                    tw = max(10, int(th * piece.width / max(1, piece.height)))
                else:
                    step = base.height / n_pages
                    y1, y2 = 0, max(1, int(round(step)))
                    piece = base.crop((0, y1, base.width, y2))
                    base_w = max(10, int(disp_w / n_pages))
                    sr = self._clamp_value(profile["size_ratio"] / 0.18, 0.6, 2.2, 1.0)
                    tw = max(10, int(base_w * sr))
                    th = max(10, int(tw * piece.height / max(1, piece.width)))
                out = piece.resize((tw, th), Image.LANCZOS)
                render_cache[cache_key] = out
                return out

            return None

        def schedule_redraw(delay_ms=16):
            if state["render_job"] is not None:
                try:
                    preview_win.after_cancel(state["render_job"])
                except Exception:
                    pass
            state["render_job"] = preview_win.after(delay_ms, redraw)

        def redraw():
            state["render_job"] = None
            active_key = get_active_key()
            if mode_key in ("seal", "seam"):
                enabled_paths = [p for p in preview_paths if enabled_vars[p].get()]
                for path in preview_paths:
                    profile = get_profile(path)
                    profile["enabled"] = bool(enabled_vars[path].get())
                    profile["opacity"] = self._clamp_value(profile.get("opacity", 0.85), 0.05, 1.0, 0.85)
                    profile["size_ratio"] = self._clamp_value(profile.get("size_ratio", 0.18), 0.03, 0.7, 0.18)
                    item = state["stamp_items"].get(path)
                    if not profile["enabled"]:
                        if item:
                            canvas.itemconfigure(item["id"], state="hidden")
                            item["bbox"] = None
                        continue

                    rendered = get_render_image(path, profile, mode_key)
                    if rendered is None:
                        continue
                    rw, rh = rendered.size
                    if mode_key == "seal":
                        cx = pad + profile["x_ratio"] * disp_w
                        cy = pad + profile["y_ratio"] * disp_h
                        x = max(pad, min(cx - rw / 2, pad + disp_w - rw))
                        y = max(pad, min(cy - rh / 2, pad + disp_h - rh))
                        profile["x_ratio"] = self._clamp_value((x + rw / 2 - pad) / max(1, disp_w), 0.0, 1.0, 0.85)
                        profile["y_ratio"] = self._clamp_value((y + rh / 2 - pad) / max(1, disp_h), 0.0, 1.0, 0.85)
                    else:
                        side = {"右侧": "right", "左侧": "left", "顶部": "top", "底部": "bottom"}.get(self.stamp_seam_side_var.get(), "right")
                        align = {"居中": "center", "顶部": "top", "底部": "bottom"}.get(self.stamp_seam_align_var.get(), "center")
                        overlap = self._clamp_value(self.stamp_seam_overlap_var.get(), 0.05, 0.95, 0.25)
                        vis_idx = enabled_paths.index(path) if path in enabled_paths else 0
                        stack_off = vis_idx * 6
                        if side in ("left", "right"):
                            y = pad if align == "top" else (pad + disp_h - rh if align == "bottom" else pad + (disp_h - rh) / 2)
                            x = pad + disp_w - rw * (1.0 - overlap) if side == "right" else pad - rw * overlap
                            y += stack_off
                        else:
                            x = pad if align == "top" else (pad + disp_w - rw if align == "bottom" else pad + (disp_w - rw) / 2)
                            y = pad - rh * overlap if side == "top" else pad + disp_h - rh * (1.0 - overlap)
                            x += stack_off

                    tk_img = ImageTk.PhotoImage(rendered)
                    if item is None:
                        cid = canvas.create_image(int(x), int(y), anchor="nw", image=tk_img)
                        state["stamp_items"][path] = {"id": cid, "photo": tk_img, "bbox": (x, y, x + rw, y + rh), "size": (rw, rh)}
                    else:
                        canvas.itemconfigure(item["id"], image=tk_img, state="normal")
                        canvas.coords(item["id"], int(x), int(y))
                        item["photo"] = tk_img
                        item["bbox"] = (x, y, x + rw, y + rh)
                        item["size"] = (rw, rh)

                active_item = state["stamp_items"].get(active_key)
                if active_item and active_item.get("bbox"):
                    x1, y1, x2, y2 = active_item["bbox"]
                    canvas.coords(state["selection_rect"], x1 - 2, y1 - 2, x2 + 2, y2 + 2)
                    canvas.itemconfigure(state["selection_rect"], state="normal")
                else:
                    canvas.itemconfigure(state["selection_rect"], state="hidden")
            else:
                single_key = "__single__"
                profile = get_profile(single_key)
                profile["opacity"] = self._clamp_value(profile.get("opacity", 0.85), 0.05, 1.0, 0.85)
                profile["size_ratio"] = self._clamp_value(profile.get("size_ratio", 0.18), 0.03, 0.7, 0.18)
                if mode_key == "qr":
                    try:
                        qr_bytes = PDFBatchStampConverter._make_qr_png_bytes(
                            self.stamp_qr_text_var.get().strip(),
                            opacity=profile["opacity"],
                            remove_white_bg=bool(self.stamp_remove_white_bg_var.get()),
                        )
                        src = Image.open(io.BytesIO(qr_bytes)).convert("RGBA")
                    except Exception:
                        src = None
                else:
                    src = self._build_template_preview_image(profile["opacity"])
                item = state["stamp_items"].get(single_key)
                if src is None:
                    if item:
                        canvas.itemconfigure(item["id"], state="hidden")
                    return
                tw = max(16, int(disp_w * profile["size_ratio"]))
                th = max(16, int(tw * src.height / max(1, src.width)))
                img = src.resize((tw, th), Image.LANCZOS)
                x = max(pad, min(pad + profile["x_ratio"] * disp_w - tw / 2, pad + disp_w - tw))
                y = max(pad, min(pad + profile["y_ratio"] * disp_h - th / 2, pad + disp_h - th))
                profile["x_ratio"] = self._clamp_value((x + tw / 2 - pad) / max(1, disp_w), 0.0, 1.0, 0.85)
                profile["y_ratio"] = self._clamp_value((y + th / 2 - pad) / max(1, disp_h), 0.0, 1.0, 0.85)
                tk_img = ImageTk.PhotoImage(img)
                if item is None:
                    cid = canvas.create_image(int(x), int(y), anchor="nw", image=tk_img)
                    state["stamp_items"][single_key] = {"id": cid, "photo": tk_img, "bbox": (x, y, x + tw, y + th), "size": (tw, th)}
                else:
                    canvas.itemconfigure(item["id"], image=tk_img, state="normal")
                    canvas.coords(item["id"], int(x), int(y))
                    item["photo"] = tk_img
                    item["bbox"] = (x, y, x + tw, y + th)
                    item["size"] = (tw, th)
                canvas.coords(state["selection_rect"], x - 2, y - 2, x + tw + 2, y + th + 2)
                canvas.itemconfigure(state["selection_rect"], state="normal")

        def on_active_change(*_args):
            sync_sliders_from_active()
            schedule_redraw(1)

        def on_enabled_change(*_args):
            schedule_redraw(1)

        def on_slider_change(_value=None):
            if state["suspend_slider"]:
                return
            key = get_active_key()
            if not key:
                return
            profile = get_profile(key)
            profile["opacity"] = self._clamp_value(opacity_var.get() / 100.0, 0.05, 1.0, 0.85)
            profile["size_ratio"] = self._clamp_value(size_var.get() / 100.0, 0.03, 0.7, 0.18)
            schedule_redraw(8)

        def hit_test_path(x, y):
            for p in reversed(preview_paths):
                item = state["stamp_items"].get(p)
                if not item or not item.get("bbox"):
                    continue
                x1, y1, x2, y2 = item["bbox"]
                if x1 <= x <= x2 and y1 <= y <= y2:
                    return p
            return ""

        def on_press(event):
            if mode_key == "seal":
                p = hit_test_path(event.x, event.y)
                if not p:
                    state["drag_path"] = None
                    return
                active_path_var.set(p)
                on_active_change()
                item = state["stamp_items"].get(p)
                if not item:
                    return
                x1, y1, x2, y2 = item["bbox"]
                state["drag_path"] = p
                state["drag_offset_x"] = event.x - (x1 + x2) / 2
                state["drag_offset_y"] = event.y - (y1 + y2) / 2
                return
            if mode_key == "qr":
                item = state["stamp_items"].get("__single__")
                if not item or not item.get("bbox"):
                    return
                x1, y1, x2, y2 = item["bbox"]
                if not (x1 <= event.x <= x2 and y1 <= event.y <= y2):
                    return
                state["drag_path"] = "__single__"
                state["drag_offset_x"] = event.x - (x1 + x2) / 2
                state["drag_offset_y"] = event.y - (y1 + y2) / 2

        def on_drag(event):
            if mode_key not in ("seal", "qr"):
                return
            p = state.get("drag_path")
            if not p:
                return
            item = state["stamp_items"].get(p)
            if not item:
                return
            w, h = item.get("size", (0, 0))
            if w <= 0 or h <= 0:
                return
            cx = event.x - state["drag_offset_x"]
            cy = event.y - state["drag_offset_y"]
            cx = max(pad + w / 2, min(cx, pad + disp_w - w / 2))
            cy = max(pad + h / 2, min(cy, pad + disp_h - h / 2))
            x = cx - w / 2
            y = cy - h / 2
            canvas.coords(item["id"], int(x), int(y))
            item["bbox"] = (x, y, x + w, y + h)
            prof = get_profile(p)
            prof["x_ratio"] = self._clamp_value((cx - pad) / max(1, disp_w), 0.0, 1.0, 0.85)
            prof["y_ratio"] = self._clamp_value((cy - pad) / max(1, disp_h), 0.0, 1.0, 0.85)
            canvas.coords(state["selection_rect"], x - 2, y - 2, x + w + 2, y + h + 2)

        def on_release(_event):
            state["drag_path"] = None

        if mode_key in ("seal", "seam"):
            for p, ev in enabled_vars.items():
                ev.trace_add("write", on_enabled_change)
            active_path_var.trace_add("write", on_active_change)

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        opacity_scale.configure(command=on_slider_change)
        size_scale.configure(command=on_slider_change)

        action_frame = tk.Frame(preview_win)
        action_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=12, pady=(0, 12))

        def export_preview_template():
            if mode_key in ("seal", "seam"):
                export_profiles = []
                for p in preview_paths:
                    prof = self._normalize_stamp_profile(preview_profiles.get(p))
                    prof["enabled"] = bool(enabled_vars[p].get())
                    if not prof["enabled"]:
                        continue
                    export_profiles.append({
                        "image_path": p,
                        "enabled": True,
                        "x_ratio": prof["x_ratio"],
                        "y_ratio": prof["y_ratio"],
                        "size_ratio": prof["size_ratio"],
                        "opacity": prof["opacity"],
                    })
                single_profile = None
            else:
                export_profiles = None
                single_profile = self._normalize_stamp_profile(preview_profiles.get("__single__"))

            path = self._export_stamp_template(
                mode_key=mode_key,
                profiles=export_profiles,
                single_profile=single_profile,
                qr_text=self.stamp_qr_text_var.get().strip(),
                pages_text=self.stamp_pages_var.get().strip(),
                parent=preview_win,
            )
            if path:
                messagebox.showinfo("导出成功", f"模板已导出：\n{path}", parent=preview_win)

        def apply_preview():
            if mode_key in ("seal", "seam"):
                for p in preview_paths:
                    profile = self._normalize_stamp_profile(preview_profiles.get(p))
                    profile["enabled"] = bool(enabled_vars[p].get())
                    self.stamp_profiles[p] = profile
                active_path = active_path_var.get()
                if active_path in self.stamp_image_paths:
                    self.stamp_selected_image_idx = self.stamp_image_paths.index(active_path)
                    self.stamp_image_path = active_path
                active_profile = self._normalize_stamp_profile(preview_profiles.get(active_path))
                self.stamp_preview_profile = {
                    "x_ratio": active_profile["x_ratio"],
                    "y_ratio": active_profile["y_ratio"],
                    "size_ratio": active_profile["size_ratio"],
                    "opacity": active_profile["opacity"],
                }
            else:
                profile = self._normalize_stamp_profile(preview_profiles.get("__single__"))
                self.stamp_preview_profile = {
                    "x_ratio": profile["x_ratio"],
                    "y_ratio": profile["y_ratio"],
                    "size_ratio": profile["size_ratio"],
                    "opacity": profile["opacity"],
                }

            self.stamp_opacity_var.set(f"{self.stamp_preview_profile['opacity']:.2f}")
            self._update_stamp_image_label()
            self._update_stamp_preview_info()
            self.save_settings()
            preview_win.destroy()

        tk.Button(action_frame, text="导出模板JSON...", command=export_preview_template,
                  font=("Microsoft YaHei", 9), width=14).pack(side=tk.LEFT)
        tk.Button(action_frame, text="取消", command=preview_win.destroy,
                  font=("Microsoft YaHei", 9), width=12).pack(side=tk.RIGHT, padx=(8, 0))
        tk.Button(action_frame, text="应用到批量盖章", command=apply_preview,
                  font=("Microsoft YaHei", 9, "bold"), width=14).pack(side=tk.RIGHT)

        sync_sliders_from_active()
        schedule_redraw(1)

    def _on_stamp_mode_changed(self, event=None):
        mode_key = self._get_stamp_mode_key()
        if mode_key == "seal":
            self.stamp_hint_var.set("普通章：支持多个章图，预览中勾选章图并拖拽定位，可调透明度和缩放。")
        elif mode_key == "qr":
            self.stamp_hint_var.set("二维码：输入内容后可在预览中拖拽位置并调整透明度。")
        elif mode_key == "seam":
            self.stamp_hint_var.set("骑缝章：支持多个章图，预览中勾选章图并查看切片效果。")
        else:
            self.stamp_hint_var.set("模板：按 JSON 模板批量盖章；可预览模板元素效果。")

        self.stamp_qr_entry.config(state=('normal' if mode_key == "qr" else 'disabled'))
        seam_state = 'readonly' if mode_key == "seam" else 'disabled'
        self.stamp_seam_side_combo.config(state=seam_state)
        self.stamp_seam_align_combo.config(state=seam_state)
        self.stamp_seam_overlap_entry.config(state=('normal' if mode_key == "seam" else 'disabled'))

        self._update_stamp_preview_info()
        self.save_settings()

    def _on_split_mode_changed(self, event=None):
        mode = self.split_mode_var.get()
        if mode == "每页一个PDF":
            self.split_param_entry.config(state='disabled')
            self.split_param_label.config(text="")
            self.split_param_hint.config(text="每页将生成一个独立PDF文件")
            self.split_param_var.set("")
        elif mode == "每N页一个PDF":
            self.split_param_entry.config(state='normal')
            self.split_param_label.config(text="N =")
            self.split_param_hint.config(text="页/文件")
            if not self.split_param_var.get():
                self.split_param_var.set("5")
        elif mode == "按范围拆分":
            self.split_param_entry.config(state='normal')
            self.split_param_label.config(text="范围:")
            self.split_param_hint.config(text="如: 1-3,4-6,7-10")
            self.split_param_var.set("")

    def _on_encrypt_mode_changed(self, event=None):
        mode = self.encrypt_mode_var.get()
        if mode == "加密":
            self.encrypt_pw_label.config(text="打开密码:")
            self.encrypt_pw_entry.config(state='normal')
            self.encrypt_owner_label.pack(side=tk.LEFT, padx=(8, 0))
            self.encrypt_owner_entry.pack(side=tk.LEFT, padx=(4, 0))
            self.panel_canvas.itemconfigure(self.cv_encrypt_perm, state='normal')
        else:
            self.encrypt_pw_label.config(text="密码:")
            self.encrypt_pw_entry.config(state='normal')
            self.encrypt_owner_label.pack_forget()
            self.encrypt_owner_entry.pack_forget()
            self.panel_canvas.itemconfigure(self.cv_encrypt_perm, state='hidden')

    def _choose_watermark_image(self):
        """选择水印图片"""
        filename = filedialog.askopenfilename(
            title="选择水印图片",
            filetypes=[("图片文件", "*.png;*.jpg;*.jpeg;*.bmp"), ("所有文件", "*.*")]
        )
        if filename:
            self.watermark_image_path = filename
            name = os.path.basename(filename)
            self.watermark_img_label.config(text=name if len(name) <= 15 else name[:12] + "...")

    def _open_file_order_dialog(self):
        """打开文件排序对话框，让用户调整文件顺序"""
        if not self.selected_files_list:
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("调整文件顺序")
        dialog.geometry("480x360")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        tk.Label(dialog, text="拖拽或使用按钮调整文件顺序（上方文件在前）",
                 font=("Microsoft YaHei", 9), fg="#666").pack(pady=(8, 4))

        list_frame = tk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(
            list_frame, font=("Microsoft YaHei", 9),
            selectmode=tk.SINGLE, yscrollcommand=scrollbar.set
        )
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        # 填充列表
        for i, f in enumerate(self.selected_files_list):
            listbox.insert(tk.END, f"{i+1}. {os.path.basename(f)}")
        if self.selected_files_list:
            listbox.selection_set(0)

        # 按钮区域
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 8))

        def move_up():
            sel = listbox.curselection()
            if not sel or sel[0] == 0:
                return
            idx = sel[0]
            self.selected_files_list[idx-1], self.selected_files_list[idx] = \
                self.selected_files_list[idx], self.selected_files_list[idx-1]
            _refresh_list(idx - 1)

        def move_down():
            sel = listbox.curselection()
            if not sel or sel[0] >= len(self.selected_files_list) - 1:
                return
            idx = sel[0]
            self.selected_files_list[idx], self.selected_files_list[idx+1] = \
                self.selected_files_list[idx+1], self.selected_files_list[idx]
            _refresh_list(idx + 1)

        def remove_item():
            sel = listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            self.selected_files_list.pop(idx)
            new_idx = min(idx, len(self.selected_files_list) - 1)
            _refresh_list(new_idx)

        def _refresh_list(select_idx=0):
            listbox.delete(0, tk.END)
            for i, f in enumerate(self.selected_files_list):
                listbox.insert(tk.END, f"{i+1}. {os.path.basename(f)}")
            if self.selected_files_list and select_idx >= 0:
                listbox.selection_set(select_idx)
                listbox.see(select_idx)

        tk.Button(btn_frame, text="⬆ 上移", command=move_up,
                  font=("Microsoft YaHei", 9), width=8, cursor='hand2'
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_frame, text="⬇ 下移", command=move_down,
                  font=("Microsoft YaHei", 9), width=8, cursor='hand2'
                  ).pack(side=tk.LEFT, padx=4)
        tk.Button(btn_frame, text="✕ 移除", command=remove_item,
                  font=("Microsoft YaHei", 9), width=8, cursor='hand2'
                  ).pack(side=tk.LEFT, padx=4)

        def on_confirm():
            count = len(self.selected_files_list)
            if count == 0:
                self.selected_file.set("")
            elif count == 1:
                self.selected_file.set(self.selected_files_list[0])
            else:
                self.selected_file.set(f"已选择 {count} 个文件")
            func = self.current_function_var.get()
            if func == "PDF合并":
                self.merge_info_label.config(
                    text=f"已选择 {count} 个文件，将按选择顺序合并")
            self._update_order_btn()
            self.status_message.set(f"文件顺序已调整，共 {count} 个文件")
            dialog.destroy()

        tk.Button(btn_frame, text="✓ 确定", command=on_confirm,
                  font=("Microsoft YaHei", 9, "bold"), width=8, cursor='hand2'
                  ).pack(side=tk.RIGHT, padx=4)

    def _update_order_btn(self):
        """多文件时显示排序按钮，否则隐藏"""
        func = self.current_function_var.get()
        show = (len(self.selected_files_list) > 1
                and func in ("图片转PDF", "PDF合并", "PDF转Word", "PDF转图片", "PDF批量文本/图片提取", "PDF批量盖章"))
        if show:
            self.order_btn.pack(side=tk.LEFT, padx=(10, 0), ipady=6)
        else:
            self.order_btn.pack_forget()

    def _update_api_hint(self):
        if not self.panel_canvas:
            return
        ocr_on = self.ocr_enabled_var.get()
        formula_on = self.formula_api_enabled_var.get()
        if not ocr_on and not formula_on:
            self.panel_canvas.itemconfigure(self.cv_api_hint, text="")
            return
        has_key = bool(self.baidu_api_key and self.baidu_secret_key)
        parts = []
        if ocr_on:
            parts.append("OCR识别")
        if formula_on:
            parts.append("公式识别")
        feature_text = " + ".join(parts)
        if has_key:
            self.panel_canvas.itemconfigure(
                self.cv_api_hint,
                text=f"已启用: {feature_text}（百度API已配置）",
                fill="#228B22"
            )
        else:
            self.panel_canvas.itemconfigure(
                self.cv_api_hint,
                text=f"已启用: {feature_text}（⚠ 请在设置中配置API Key）",
                fill="#CC0000"
            )

    # ==========================================================
    # 文件操作
    # ==========================================================

    def check_dependencies(self):
        missing = []
        if not PDF2DOCX_AVAILABLE:
            missing.append("pdf2docx")
        if missing:
            msg = (f"警告：以下依赖库未安装：\n{', '.join(missing)}\n\n"
                   f"请运行: pip install {' '.join(missing)}")
            self.status_message.set(f"缺少依赖库: {', '.join(missing)}")
            messagebox.showwarning("缺少依赖", msg)

    def browse_file(self):
        func = self.current_function_var.get()

        if func in ("PDF转Word", "PDF转图片", "PDF合并", "PDF批量文本/图片提取", "PDF批量盖章"):
            # 多选PDF文件
            filenames = filedialog.askopenfilenames(
                title="选择PDF文件（可多选）",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
            )
            if filenames:
                self.selected_files_list = list(filenames)
                count = len(self.selected_files_list)
                if count == 1:
                    self.selected_file.set(filenames[0])
                    self.status_message.set(f"已选择: {os.path.basename(filenames[0])}")
                else:
                    self.selected_file.set(f"已选择 {count} 个PDF文件")
                    names = ", ".join(os.path.basename(f) for f in filenames[:3])
                    if count > 3:
                        names += f" 等共{count}个"
                    self.status_message.set(f"已选择: {names}")
                # 更新合并信息
                if func == "PDF合并":
                    self.merge_info_label.config(
                        text=f"已选择 {count} 个文件，将按选择顺序合并"
                    )

        elif func == "PDF拆分":
            # 单选PDF
            filename = filedialog.askopenfilename(
                title="选择PDF文件",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
            )
            if filename:
                self.selected_file.set(filename)
                self.selected_files_list = [filename]
                self.status_message.set(f"已选择: {os.path.basename(filename)}")

        elif func == "图片转PDF":
            # 多选图片
            filenames = filedialog.askopenfilenames(
                title="选择图片文件（可多选）",
                filetypes=[
                    ("图片文件", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tiff;*.tif;*.webp"),
                    ("所有文件", "*.*")
                ]
            )
            if filenames:
                self.selected_files_list = list(filenames)
                count = len(self.selected_files_list)
                if count == 1:
                    self.selected_file.set(filenames[0])
                    self.status_message.set(
                        f"已选择: {os.path.basename(filenames[0])}")
                else:
                    self.selected_file.set(f"已选择 {count} 张图片")
                    names = ", ".join(os.path.basename(f) for f in filenames[:3])
                    if count > 3:
                        names += f" 等共{count}个"
                    self.status_message.set(f"已选择: {names}")

        elif func in ("PDF加水印", "PDF加密/解密", "PDF压缩", "PDF提取/删页", "OCR可搜索PDF", "PDF转Excel"):
            # 单选PDF
            filename = filedialog.askopenfilename(
                title="选择PDF文件",
                filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
            )
            if filename:
                self.selected_file.set(filename)
                self.selected_files_list = [filename]
                self.status_message.set(f"已选择: {os.path.basename(filename)}")

        self._update_order_btn()

    def clear_selection(self):
        self.selected_file.set("")
        self.selected_files_list = []
        self._update_order_btn()
        self.progress_bar['value'] = 0
        self.set_progress_text("")
        self.status_message.set("就绪")
        self.total_pages = 0
        self.total_steps = 0
        self.start_time = None
        self.current_page_id = None
        self.current_page_index = None
        self.current_page_total = None
        self.current_phase = None
        self.page_start_time = None
        self.current_eta_text = ""
        self.base_status_text = ""
        self.conversion_active = False
        self.page_start_var.set("")
        self.page_end_var.set("")

    # ==========================================================
    # 拖拽文件支持
    # ==========================================================

    def _on_drop_files(self, files):
        """处理拖拽文件"""
        decoded = []
        for f in files:
            if isinstance(f, bytes):
                try:
                    decoded.append(f.decode('utf-8'))
                except UnicodeDecodeError:
                    try:
                        decoded.append(f.decode('gbk'))
                    except Exception:
                        decoded.append(f.decode('latin-1'))
            else:
                decoded.append(str(f))

        func = self.current_function_var.get()

        if func == '图片转PDF':
            valid = [f for f in decoded
                     if os.path.splitext(f)[1].lower() in SUPPORTED_IMAGE_EXTS]
            if not valid:
                self.status_message.set("拖拽的文件中没有支持的图片格式")
                return
        else:
            valid = [f for f in decoded if f.lower().endswith('.pdf')]
            if not valid:
                self.status_message.set("拖拽的文件中没有PDF文件")
                return

        self.selected_files_list = valid
        count = len(valid)
        if count == 1:
            self.selected_file.set(valid[0])
            self.status_message.set(f"拖拽导入: {os.path.basename(valid[0])}")
        else:
            self.selected_file.set(f"已拖拽 {count} 个文件")
            names = ", ".join(os.path.basename(f) for f in valid[:3])
            if count > 3:
                names += f" 等共{count}个"
            self.status_message.set(f"拖拽导入: {names}")

        # 更新合并信息
        if func == "PDF合并":
            self.merge_info_label.config(
                text=f"已选择 {count} 个文件，将按选择顺序合并"
            )

        self._update_order_btn()

    # ==========================================================
    # 转换入口
    # ==========================================================

    def start_conversion(self):
        func = self.current_function_var.get()

        # 验证文件选择
        if func == "图片转PDF":
            if not self.selected_files_list:
                messagebox.showwarning("提示", "请先选择图片文件！")
                return
        elif func == "PDF合并":
            if len(self.selected_files_list) < 2:
                messagebox.showwarning("提示", "请至少选择2个PDF文件进行合并！")
                return
        else:
            if not self.selected_files_list:
                messagebox.showwarning("提示", "请先选择文件！")
                return
            if func == "PDF批量文本/图片提取":
                if not self.batch_text_enabled_var.get() and not self.batch_image_enabled_var.get():
                    messagebox.showwarning("提示", "请至少选择文本或图片提取！")
                    return
            if func == "PDF批量盖章":
                mode_key = self._get_stamp_mode_key()
                if mode_key in ("seal", "seam") and not self._get_enabled_stamp_profiles():
                    messagebox.showwarning("提示", "请先选择章图，并至少勾选一个章图。")
                    return
                if mode_key == "qr" and not self.stamp_qr_text_var.get().strip():
                    messagebox.showwarning("提示", "请填写二维码内容。")
                    return
                if mode_key == "template" and not self.stamp_template_path:
                    messagebox.showwarning("提示", "请先选择模板JSON。")
                    return

        for f in self.selected_files_list:
            if not os.path.exists(f):
                messagebox.showerror("错误", f"文件不存在：\n{f}")
                return

        self.convert_btn.config(state=tk.DISABLED)
        self.conversion_active = True
        self.current_page_id = None
        self.current_page_index = None
        self.current_page_total = None
        self.current_phase = None
        self.page_start_time = None
        self.current_eta_text = ""
        self.base_status_text = ""
        self.start_page_timer()

        thread = threading.Thread(target=self.perform_conversion)
        thread.daemon = True
        thread.start()

    def perform_conversion(self):
        try:
            func = self.current_function_var.get()
            if func == "PDF转Word":
                self._do_convert_to_word()
            elif func == "PDF转图片":
                self._do_convert_to_images()
            elif func == "PDF合并":
                self._do_convert_merge()
            elif func == "PDF拆分":
                self._do_convert_split()
            elif func == "图片转PDF":
                self._do_convert_img2pdf()
            elif func == "PDF加水印":
                self._do_convert_watermark()
            elif func == "PDF加密/解密":
                self._do_convert_encrypt()
            elif func == "PDF压缩":
                self._do_convert_compress()
            elif func == "PDF提取/删页":
                self._do_convert_extract()
            elif func == "OCR可搜索PDF":
                self._do_convert_ocr()
            elif func == "PDF转Excel":
                self._do_convert_excel()
            elif func == "PDF批量文本/图片提取":
                self._do_convert_batch_extract()
            elif func == "PDF批量盖章":
                self._do_convert_batch_stamp()
        except Exception as e:
            logging.error(f"转换异常: {e}", exc_info=True)
            self.root.after(0, lambda: messagebox.showerror(
                "转换失败", f"转换过程中出错：\n{str(e)}"))
            self.root.after(0, lambda: self.status_message.set("转换失败"))
        finally:
            with self._state_lock:
                self.conversion_active = False
            self.stop_page_timer()
            self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))

    # ----------------------------------------------------------
    # PDF → Word（支持批量）
    # ----------------------------------------------------------

    def _do_convert_to_word(self):
        files = self.selected_files_list
        if not files:
            return

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        total_files = len(files)

        # 批量模式下忽略页范围（每个文件页数不同，统一应用不合理）
        if total_files > 1:
            start_page, end_page = 0, None
            # 用户设置了页范围时提示
            if self.page_start_var.get().strip() or self.page_end_var.get().strip():
                self.root.after(0, lambda: self.status_message.set(
                    "批量模式已自动忽略页范围，每个文件将全部转换"))
        else:
            start_page, end_page = self._parse_page_range_for_converter()
        results = []

        for file_idx, input_file in enumerate(files):
            output_file = self.generate_output_filename(input_file, '.docx')

            if total_files > 1:
                # 批量模式：用包装回调显示总体进度
                def make_progress_cb(fi, tf):
                    def cb(percent, progress_text, status_text):
                        overall = int((fi / tf + max(0, percent) / 100 / tf) * 100)
                        file_label = os.path.basename(files[fi])
                        self._simple_progress_callback(
                            overall,
                            f"[{fi + 1}/{tf}] {file_label}: {progress_text}",
                            status_text or f"正在转换: {file_label}"
                        )
                    return cb

                converter = PDFToWordConverter(
                    on_progress=make_progress_cb(file_idx, total_files),
                    pdf2docx_progress=None,  # 批量模式跳过详细进度
                )
            else:
                converter = PDFToWordConverter(
                    on_progress=self._simple_progress_callback,
                    pdf2docx_progress=self.update_progress,
                )

            result = converter.convert(
                input_file, output_file,
                start_page=start_page, end_page=end_page,
                ocr_enabled=self.ocr_enabled_var.get(),
                formula_api_enabled=self.formula_api_enabled_var.get(),
                api_key=self.baidu_api_key,
                secret_key=self.baidu_secret_key,
                xslt_path=self.xslt_path,
            )
            results.append((input_file, output_file, result))

            # 记录历史
            self.history.add({
                'function': 'PDF转Word',
                'input_files': [input_file],
                'output': output_file,
                'success': result['success'],
                'message': result.get('message', ''),
                'page_count': result.get('page_count', 0),
            })

        # 显示结果
        if total_files == 1:
            self._show_single_word_result(results[0])
        else:
            self._show_batch_word_result(results)

    def _show_single_word_result(self, result_tuple):
        """显示单文件Word转换结果"""
        _, output_file, result = result_tuple

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                "转换失败", result.get('message', '未知错误')))
            self.root.after(0, lambda: self.status_message.set("转换失败"))
            return

        mode_text = "OCR模式" if result.get('mode') == 'ocr' else ""
        success_msg = f"PDF已成功转换为Word！{mode_text}\n\n保存位置：\n{output_file}"
        if result.get('formula_count', 0) > 0:
            success_msg += f"\n\n已识别并转换 {result['formula_count']} 处数学公式为可编辑格式"
        if result.get('page_count', 0) > 0:
            success_msg += f"\n共处理 {result['page_count']} 页"
        if result.get('errors'):
            success_msg += f"\n\n⚠ {len(result['errors'])} 页识别出错（已用图片替代）"
        success_msg += "\n\n是否打开文件所在文件夹？"

        def _show():
            if messagebox.askyesno("转换成功", success_msg):
                self.open_folder(output_file)
            if result.get('skipped_pages'):
                skipped = self.format_skipped_pages(result['skipped_pages'])
                messagebox.showwarning("跳过异常页",
                                       f"以下页面在转换中被跳过：\n{skipped}")
            if result.get('errors'):
                err_detail = "\n".join(result['errors'][:10])
                messagebox.showwarning("OCR识别警告",
                                       f"以下页面识别失败：\n{err_detail}")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # PDF批量文本/图片提取
    # ----------------------------------------------------------

    def _do_convert_batch_extract(self):
        converter = PDFBatchExtractConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        text_mode_val = self.batch_text_mode_var.get()
        text_mode = "merge" if "合并" in text_mode_val else "per_page"

        result = converter.convert(
            files=list(self.selected_files_list),
            pages_str=self.batch_pages_var.get().strip(),
            extract_text=bool(self.batch_text_enabled_var.get()),
            extract_images=bool(self.batch_image_enabled_var.get()),
            text_format=self.batch_text_format_var.get(),
            text_mode=text_mode,
            preserve_layout=bool(self.batch_preserve_layout_var.get()),
            ocr_enabled=bool(self.batch_ocr_enabled_var.get()),
            api_key=self.baidu_api_key,
            secret_key=self.baidu_secret_key,
            image_per_page=bool(self.batch_image_per_page_var.get()),
            image_dedupe=bool(self.batch_image_dedupe_var.get()),
            image_format=self.batch_image_format_var.get(),
            zip_output=bool(self.batch_zip_enabled_var.get()),
        )

        # 记录历史
        self.history.add({
            'function': 'PDF批量文本/图片提取',
            'input_files': list(self.selected_files_list),
            'output': result.get('output_dir', ''),
            'success': result.get('success', False),
            'message': result.get('message', ''),
            'page_count': result.get('stats', {}).get('page_count', 0),
        })

        if not result.get('success'):
            self.root.after(0, lambda: messagebox.showerror(
                "批量提取失败", result.get('message', '未知错误')))
            self.root.after(0, lambda: self.status_message.set("批量提取失败"))
            return

        output_dir = result.get('output_dir', '')
        output_zip = result.get('output_zip', '')

        def _show():
            msg = (f"{result.get('message', '')}\n\n"
                   f"输出目录：\n{output_dir}")
            if output_zip:
                msg += f"\n\n已生成ZIP：\n{output_zip}"
            msg += "\n\n是否打开输出文件夹？"
            if messagebox.askyesno("批量提取完成", msg):
                self.open_folder(output_dir)
            self.status_message.set("批量提取完成")

        self.root.after(0, _show)

    def _do_convert_batch_stamp(self):
        converter = PDFBatchStampConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        mode_key = self._get_stamp_mode_key()
        seam_side_map = {"右侧": "right", "左侧": "left", "顶部": "top", "底部": "bottom"}
        seam_align_map = {"居中": "center", "顶部": "top", "底部": "bottom"}
        opacity_value = self._clamp_value(
            self.stamp_preview_profile.get("opacity", self.stamp_opacity_var.get()),
            0.05,
            1.0,
            0.85,
        )
        size_ratio = self._clamp_value(
            self.stamp_preview_profile.get("size_ratio", 0.18),
            0.03,
            0.7,
            0.18,
        )
        placement = {
            "x_ratio": self._clamp_value(self.stamp_preview_profile.get("x_ratio", 0.85), 0.0, 1.0, 0.85),
            "y_ratio": self._clamp_value(self.stamp_preview_profile.get("y_ratio", 0.85), 0.0, 1.0, 0.85),
            "size_ratio": size_ratio,
        }
        stamp_profiles = self._get_enabled_stamp_profiles()

        result = converter.convert(
            files=list(self.selected_files_list),
            mode=mode_key,
            pages_str=self.stamp_pages_var.get().strip(),
            opacity=opacity_value,
            position="right_bottom",
            size_ratio=size_ratio,
            seal_image_path=self._get_active_stamp_image_path(),
            qr_text=self.stamp_qr_text_var.get().strip(),
            seam_side=seam_side_map.get(self.stamp_seam_side_var.get(), "right"),
            seam_align=seam_align_map.get(self.stamp_seam_align_var.get(), "center"),
            seam_overlap_ratio=self.stamp_seam_overlap_var.get().strip() or "0.25",
            template_path=self.stamp_template_path,
            placement=placement,
            remove_white_bg=bool(self.stamp_remove_white_bg_var.get()),
            stamp_profiles=stamp_profiles,
        )

        self.history.add({
            'function': 'PDF批量盖章',
            'input_files': list(self.selected_files_list),
            'output': ', '.join(result.get('output_files', [])),
            'success': result.get('success', False),
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result.get('success'):
            self.root.after(0, lambda: messagebox.showerror(
                "批量盖章失败", result.get('message', '未知错误')))
            self.root.after(0, lambda: self.status_message.set("批量盖章失败"))
            return

        output_files = result.get('output_files', [])

        def _show():
            msg = (f"{result.get('message', '批量盖章完成')}\n\n"
                   f"输出文件数量：{len(output_files)}")
            if output_files:
                msg += f"\n\n示例输出：\n{output_files[0]}"
            msg += "\n\n是否打开输出文件夹？"
            if messagebox.askyesno("批量盖章完成", msg) and output_files:
                self.open_folder(output_files[0])
            self.status_message.set("批量盖章完成")

        self.root.after(0, _show)

    def _show_batch_word_result(self, results):
        """显示批量Word转换结果"""
        total = len(results)
        success_count = sum(1 for _, _, r in results if r['success'])
        fail_count = total - success_count
        total_pages = sum(r.get('page_count', 0) for _, _, r in results)

        def _show():
            if fail_count == 0:
                msg = (f"批量转换完成！\n\n"
                       f"成功: {success_count} 个文件\n"
                       f"共 {total_pages} 页\n\n"
                       f"输出文件保存在各PDF同目录下")
                messagebox.showinfo("批量转换完成", msg)
            else:
                msg = (f"批量转换部分完成\n\n"
                       f"成功: {success_count} 个\n"
                       f"失败: {fail_count} 个")
                for f, _, r in results:
                    if not r['success']:
                        msg += f"\n\n❌ {os.path.basename(f)}: {r.get('message', '未知错误')}"
                messagebox.showwarning("批量转换", msg)

            self.status_message.set(
                f"转换完成: {success_count}/{total} 成功, 共{total_pages}页")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # PDF → 图片
    # ----------------------------------------------------------

    def _do_convert_to_images(self):
        converter = PDFToImageConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        # 解析页范围
        start_text = self.page_start_var.get().strip()
        end_text = self.page_end_var.get().strip()
        start_page = int(start_text) if start_text and start_text.isdigit() else None
        end_page = int(end_text) if end_text and end_text.isdigit() else None

        result = converter.convert(
            files=self.selected_files_list,
            dpi=self.image_dpi_var.get(),
            img_format=self.image_format_var.get(),
            start_page=start_page,
            end_page=end_page,
        )

        # 记录历史
        self.history.add({
            'function': 'PDF转图片',
            'input_files': list(self.selected_files_list),
            'output': ', '.join(result.get('output_dirs', [])),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result['success'] and result.get('message'):
            self.root.after(0, lambda: messagebox.showerror(
                "转换失败", result['message']))
            self.root.after(0, lambda: self.status_message.set("转换失败"))
            return

        def _show():
            output_dirs = result['output_dirs']
            errors = result['errors']
            processed = result['page_count']
            files = self.selected_files_list
            dpi = result['dpi']
            img_format = result['format']

            if errors:
                err_msg = "\n".join(errors)
                msg = f"转换完成，但有 {len(errors)} 个文件出错：\n\n{err_msg}"
                if output_dirs:
                    msg += "\n\n成功的文件已保存到各PDF同目录下的文件夹中"
                messagebox.showwarning("部分完成", msg)
            else:
                if len(files) == 1:
                    msg = (f"PDF已成功转换为图片！\n\nDPI: {dpi}  格式: {img_format}\n"
                           f"共 {processed} 页\n\n保存位置：\n{output_dirs[0]}")
                else:
                    dir_list = "\n".join(output_dirs[:5])
                    if len(output_dirs) > 5:
                        dir_list += f"\n...等共 {len(output_dirs)} 个文件夹"
                    msg = (f"所有PDF已成功转换为图片！\n\nDPI: {dpi}  格式: {img_format}\n"
                           f"共 {len(files)} 个文件，{processed} 页\n\n保存位置：\n{dir_list}")
                messagebox.showinfo("转换成功", msg)

            if output_dirs:
                try:
                    os.startfile(output_dirs[0])
                except Exception:
                    pass

            self.status_message.set(
                f"转换完成：{len(files)}个文件，共{processed}页")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # PDF 合并
    # ----------------------------------------------------------

    def _do_convert_merge(self):
        converter = PDFMergeConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        result = converter.convert(files=self.selected_files_list)

        # 记录历史
        self.history.add({
            'function': 'PDF合并',
            'input_files': list(self.selected_files_list),
            'output': result.get('output_file', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                "合并失败", result['message']))
            self.root.after(0, lambda: self.status_message.set("合并失败"))
            return

        output_file = result['output_file']
        page_count = result['page_count']
        file_count = result['file_count']

        def _show():
            msg = (f"PDF合并成功！\n\n"
                   f"合并了 {file_count} 个文件，共 {page_count} 页\n\n"
                   f"保存位置：\n{output_file}\n\n"
                   f"是否打开文件所在文件夹？")
            if messagebox.askyesno("合并成功", msg):
                self.open_folder(output_file)
            self.status_message.set(
                f"合并完成: {file_count}个文件, {page_count}页")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # PDF 拆分
    # ----------------------------------------------------------

    def _do_convert_split(self):
        converter = PDFSplitConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        # 解析拆分模式
        mode_text = self.split_mode_var.get()
        mode_map = {
            "每页一个PDF": "every_page",
            "每N页一个PDF": "by_interval",
            "按范围拆分": "by_ranges",
        }
        mode = mode_map.get(mode_text, "every_page")

        interval = 1
        ranges = None
        if mode == "by_interval":
            try:
                interval = int(self.split_param_var.get())
                if interval < 1:
                    raise ValueError
            except (ValueError, TypeError):
                self.root.after(0, lambda: messagebox.showerror(
                    "参数错误", "请输入有效的页数（正整数）"))
                return
        elif mode == "by_ranges":
            ranges = self.split_param_var.get().strip()
            if not ranges:
                self.root.after(0, lambda: messagebox.showerror(
                    "参数错误", "请输入拆分范围，如：1-3,4-6,7-10"))
                return

        result = converter.convert(
            input_file=self.selected_files_list[0],
            mode=mode,
            interval=interval,
            ranges=ranges,
        )

        # 记录历史
        self.history.add({
            'function': 'PDF拆分',
            'input_files': list(self.selected_files_list),
            'output': result.get('output_dir', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                "拆分失败", result['message']))
            self.root.after(0, lambda: self.status_message.set("拆分失败"))
            return

        output_dir = result['output_dir']
        file_count = result['file_count']
        page_count = result['page_count']

        def _show():
            msg = (f"PDF拆分成功！\n\n"
                   f"共 {page_count} 页拆分为 {file_count} 个文件\n\n"
                   f"保存位置：\n{output_dir}")
            messagebox.showinfo("拆分成功", msg)
            try:
                os.startfile(output_dir)
            except Exception:
                pass
            self.status_message.set(f"拆分完成: {file_count}个文件")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # 图片 → PDF
    # ----------------------------------------------------------

    def _do_convert_img2pdf(self):
        converter = ImageToPDFConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        result = converter.convert(
            files=self.selected_files_list,
            page_size=self.page_size_var.get(),
        )

        # 记录历史
        self.history.add({
            'function': '图片转PDF',
            'input_files': list(self.selected_files_list),
            'output': result.get('output_file', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                "转换失败", result['message']))
            self.root.after(0, lambda: self.status_message.set("转换失败"))
            return

        output_file = result['output_file']
        page_count = result['page_count']

        def _show():
            msg = (f"图片转PDF成功！\n\n"
                   f"共 {page_count} 张图片\n\n"
                   f"保存位置：\n{output_file}\n\n"
                   f"是否打开文件所在文件夹？")
            if messagebox.askyesno("转换成功", msg):
                self.open_folder(output_file)
            self.status_message.set(f"转换完成: {page_count}张图片")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # PDF 加水印
    # ----------------------------------------------------------

    def _do_convert_watermark(self):
        converter = PDFWatermarkConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        # 位置映射
        pos_map = {
            "平铺": "tile", "居中": "center",
            "左上角": "top-left", "右上角": "top-right",
            "左下角": "bottom-left", "右下角": "bottom-right",
        }
        position = pos_map.get(self.watermark_position_var.get(), "tile")

        try:
            opacity = float(self.watermark_opacity_var.get())
        except (ValueError, TypeError):
            opacity = 0.3

        try:
            font_size = int(self.watermark_fontsize_var.get())
        except (ValueError, TypeError):
            font_size = 40

        try:
            rotation = int(self.watermark_rotation_var.get())
        except (ValueError, TypeError):
            rotation = 45

        result = converter.convert(
            input_file=self.selected_files_list[0],
            watermark_text=self.watermark_text_var.get().strip() or None,
            watermark_image=self.watermark_image_path,
            opacity=opacity,
            rotation=rotation,
            font_size=font_size,
            position=position,
        )

        # 记录历史
        self.history.add({
            'function': 'PDF加水印',
            'input_files': list(self.selected_files_list),
            'output': result.get('output_file', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                "水印失败", result['message']))
            self.root.after(0, lambda: self.status_message.set("添加水印失败"))
            return

        output_file = result['output_file']
        page_count = result['page_count']

        def _show():
            msg = (f"水印添加成功！\n\n"
                   f"共 {page_count} 页\n\n"
                   f"保存位置：\n{output_file}\n\n"
                   f"是否打开文件所在文件夹？")
            if messagebox.askyesno("水印成功", msg):
                self.open_folder(output_file)
            self.status_message.set(f"水印完成: {page_count}页")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # PDF 加密/解密
    # ----------------------------------------------------------

    def _do_convert_encrypt(self):
        converter = PDFEncryptConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        mode = self.encrypt_mode_var.get()
        input_file = self.selected_files_list[0]

        if mode == "加密":
            result = converter.encrypt(
                input_file=input_file,
                user_password=self.user_password_var.get(),
                owner_password=self.owner_password_var.get(),
                allow_print=self.allow_print_var.get(),
                allow_copy=self.allow_copy_var.get(),
                allow_modify=self.allow_modify_var.get(),
                allow_annotate=self.allow_annotate_var.get(),
            )
            func_name = 'PDF加密'
        else:
            result = converter.decrypt(
                input_file=input_file,
                password=self.user_password_var.get(),
            )
            func_name = 'PDF解密'

        # 记录历史
        self.history.add({
            'function': func_name,
            'input_files': list(self.selected_files_list),
            'output': result.get('output_file', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                f"{func_name}失败", result['message']))
            self.root.after(0, lambda: self.status_message.set(f"{func_name}失败"))
            return

        output_file = result['output_file']

        def _show():
            msg = (f"{result['message']}\n\n"
                   f"保存位置：\n{output_file}\n\n"
                   f"是否打开文件所在文件夹？")
            if messagebox.askyesno(f"{func_name}成功", msg):
                self.open_folder(output_file)
            self.status_message.set(f"{func_name}完成")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # PDF 压缩
    # ----------------------------------------------------------

    def _do_convert_compress(self):
        converter = PDFCompressConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        input_file = self.selected_files_list[0]
        compress_level = self.compress_level_var.get()

        result = converter.convert(
            input_file=input_file,
            compress_level=compress_level,
        )

        # 记录历史
        self.history.add({
            'function': 'PDF压缩',
            'input_files': list(self.selected_files_list),
            'output': result.get('output_file', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                "PDF压缩失败", result['message']))
            self.root.after(0, lambda: self.status_message.set("压缩失败"))
            return

        output_file = result['output_file']

        def _show():
            msg = (f"{result['message']}\n\n"
                   f"保存位置：\n{output_file}\n\n"
                   f"是否打开文件所在文件夹？")
            if messagebox.askyesno("PDF压缩完成", msg):
                self.open_folder(output_file)
            self.status_message.set("压缩完成")

        self.root.after(0, _show)

    def _on_compress_level_changed(self):
        """压缩级别切换时更新说明文字"""
        level = self.compress_level_var.get()
        preset = COMPRESS_PRESETS.get(level, {})
        self.compress_hint_var.set(preset.get('description', ''))

    # ----------------------------------------------------------
    # PDF 提取/删页
    # ----------------------------------------------------------

    def _do_convert_extract(self):
        converter = PDFExtractConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        input_file = self.selected_files_list[0]
        mode = self.extract_mode_var.get()
        pages_str = self.extract_pages_var.get()

        result = converter.convert(
            input_file=input_file,
            pages_str=pages_str,
            mode=mode,
        )

        func_name = f'PDF{mode}页面'

        # 记录历史
        self.history.add({
            'function': func_name,
            'input_files': list(self.selected_files_list),
            'output': result.get('output_file', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('result_pages', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                f"{func_name}失败", result['message']))
            self.root.after(0, lambda: self.status_message.set(f"{func_name}失败"))
            return

        output_file = result['output_file']

        def _show():
            msg = (f"{result['message']}\n\n"
                   f"保存位置：\n{output_file}\n\n"
                   f"是否打开文件所在文件夹？")
            if messagebox.askyesno(f"{func_name}完成", msg):
                self.open_folder(output_file)
            self.status_message.set(f"{func_name}完成")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # OCR可搜索PDF
    # ----------------------------------------------------------

    def _do_convert_ocr(self):
        converter = PDFOCRConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        input_file = self.selected_files_list[0]
        start_page, end_page = self._parse_page_range_for_converter()

        # _parse_page_range_for_converter 返回 (0-based start, end)
        # PDFOCRConverter.convert 期望 1-based start_page
        ocr_start = (start_page + 1) if start_page else None
        ocr_end = end_page  # end_page 已经是1-based

        result = converter.convert(
            input_file=input_file,
            api_key=self.baidu_api_key,
            secret_key=self.baidu_secret_key,
            start_page=ocr_start,
            end_page=ocr_end,
        )

        # 记录历史
        self.history.add({
            'function': 'OCR可搜索PDF',
            'input_files': list(self.selected_files_list),
            'output': result.get('output_file', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('page_count', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                "OCR失败", result['message']))
            self.root.after(0, lambda: self.status_message.set("OCR失败"))
            return

        output_file = result['output_file']

        def _show():
            msg = (f"{result['message']}\n\n"
                   f"保存位置：\n{output_file}\n\n"
                   f"是否打开文件所在文件夹？")
            if messagebox.askyesno("OCR完成", msg):
                self.open_folder(output_file)
            self.status_message.set("OCR可搜索PDF完成")

        self.root.after(0, _show)

    # ----------------------------------------------------------
    # PDF转Excel
    # ----------------------------------------------------------

    def _on_excel_strategy_changed(self):
        """Excel提取策略切换时更新说明文字"""
        strategy = self.excel_strategy_var.get()
        info = TABLE_STRATEGIES.get(strategy, {})
        self.excel_hint_var.set(info.get('description', ''))

    def _do_convert_excel(self):
        converter = PDFToExcelConverter(
            on_progress=self._simple_progress_callback
        )

        self.root.after(0, lambda: self.progress_bar.config(
            mode='determinate', maximum=100, value=0))
        self.start_time = time.time()

        input_file = self.selected_files_list[0]
        start_page, end_page = self._parse_page_range_for_converter()

        # _parse_page_range_for_converter 返回 (0-based start, end_page 1-based)
        # PDFToExcelConverter.convert 期望 1-based start_page
        excel_start = (start_page + 1) if start_page else None
        excel_end = end_page

        strategy = self.excel_strategy_var.get()
        merge_sheets = self.excel_merge_var.get()
        extract_mode = self.excel_extract_mode_var.get()

        result = converter.convert(
            input_file=input_file,
            start_page=excel_start,
            end_page=excel_end,
            strategy=strategy,
            merge_sheets=merge_sheets,
            extract_mode=extract_mode,
            api_key=self.baidu_api_key,
            secret_key=self.baidu_secret_key,
        )

        # 记录历史
        self.history.add({
            'function': 'PDF转Excel',
            'input_files': list(self.selected_files_list),
            'output': result.get('output_file', ''),
            'success': result['success'],
            'message': result.get('message', ''),
            'page_count': result.get('table_count', 0),
        })

        if not result['success']:
            self.root.after(0, lambda: messagebox.showerror(
                "提取失败", result['message']))
            self.root.after(0, lambda: self.status_message.set("PDF转Excel失败"))
            return

        output_file = result['output_file']

        def _show():
            msg = (f"{result['message']}\n\n"
                   f"保存位置：\n{output_file}\n\n"
                   f"是否打开文件所在文件夹？")
            if messagebox.askyesno("PDF转Excel完成", msg):
                self.open_folder(output_file)
            self.status_message.set("PDF转Excel完成")

        self.root.after(0, _show)

    # ==========================================================
    # 进度回调
    # ==========================================================

    def _simple_progress_callback(self, percent, progress_text, status_text):
        """通用进度回调（线程安全）— 供 converters 使用"""
        if percent >= 0:
            self.root.after(0, lambda: self.progress_bar.config(value=percent))
        if progress_text:
            self.root.after(0, lambda t=progress_text: self.set_progress_text(t))
        if status_text:
            with self._state_lock:
                self.base_status_text = status_text
            self.root.after(0, self.apply_status_text)

    def update_progress(self, phase, current, total, page_id):
        """pdf2docx ProgressConverter 的详细进度回调"""
        if total <= 0:
            return

        total_steps = total * 2
        if phase in ('start-parse', 'start-make'):
            phase_text = "解析" if phase == 'start-parse' else "生成"
            self.current_phase = phase_text
            self.current_page_id = page_id
            self.current_page_index = current
            self.current_page_total = total
            self.page_start_time = time.time()
            with self._state_lock:
                self.base_status_text = f"正在{phase_text}第 {page_id} 页，共 {total} 页"
            self.root.after(0, self.apply_status_text)
            return

        if phase in ('skip-parse', 'skip-make'):
            phase_text = "解析" if phase == 'skip-parse' else "生成"
            with self._state_lock:
                self.base_status_text = f"第 {page_id} 页{phase_text}失败，已跳过"
            self.root.after(0, self.apply_status_text)
            return

        if phase == 'parse':
            completed_steps = current
            percent = int(round((completed_steps / total_steps) * 100))
            phase_text = "解析"
        else:
            completed_steps = total + current
            percent = int(round((completed_steps / total_steps) * 100))
            phase_text = "生成"

        page_text = self.format_page_text(phase_text, current, total, page_id)
        with self._state_lock:
            self.base_status_text = f"正在{phase_text}第 {page_id} 页，共 {total} 页"

        eta_text = ""
        if self.start_time and completed_steps > 0:
            elapsed = time.time() - self.start_time
            remaining = max(total_steps - completed_steps, 0)
            eta_seconds = int(round(elapsed * remaining / completed_steps))
            eta_text = f"，预计剩余 {self.format_eta(eta_seconds)}"
        with self._state_lock:
            self.current_eta_text = eta_text

        def _apply():
            self.progress_bar.config(mode='determinate', maximum=100)
            self.progress_bar['value'] = percent
            self.set_progress_text(f"{page_text} ({percent}%)")
            self.apply_status_text()

        self.root.after(0, _apply)

    def apply_status_text(self):
        with self._state_lock:
            text = self.base_status_text or ""
            eta = self.current_eta_text
        if eta:
            text += eta
        if self.page_start_time:
            elapsed = int(time.time() - self.page_start_time)
            text += f"，当前页耗时 {self.format_eta(elapsed)}"
            if elapsed >= self.page_timeout_seconds:
                text += "，该页复杂请耐心等待"
        if text:
            self.status_message.set(text)

    def format_page_text(self, phase_text, current, total, page_id):
        if self.total_pages and total != self.total_pages:
            return f"{phase_text}页 {current}/{total} (原页 {page_id})"
        return f"{phase_text}页 {page_id}/{total}"

    @staticmethod
    def format_eta(seconds):
        minutes, sec = divmod(max(seconds, 0), 60)
        hours, minutes = divmod(minutes, 60)
        if hours > 0:
            return f"{hours}小时{minutes}分{sec}秒"
        if minutes > 0:
            return f"{minutes}分{sec}秒"
        return f"{sec}秒"

    # ==========================================================
    # 计时器
    # ==========================================================

    def start_page_timer(self):
        if self.page_timer_job is not None:
            return
        self.page_timer_job = self.root.after(1000, self.refresh_page_timer)

    def stop_page_timer(self):
        if self.page_timer_job is not None:
            try:
                self.root.after_cancel(self.page_timer_job)
            except Exception:
                pass
            self.page_timer_job = None

    def refresh_page_timer(self):
        self.apply_status_text()
        if self.conversion_active:
            self.page_timer_job = self.root.after(1000, self.refresh_page_timer)
        else:
            self.page_timer_job = None

    # ==========================================================
    # 设置窗口 & 历史窗口
    # ==========================================================

    def open_settings_window(self):
        from ui.dialogs import open_settings_window
        open_settings_window(self)

    def open_history_window(self):
        from ui.dialogs import open_history_window
        open_history_window(self)

    def apply_title_text(self):
        text = self.title_text_var.get().strip() or "PDF转换工具"
        self.title_text_var.set(text)
        self.save_settings()

    def on_opacity_change(self, _value=None):
        self.apply_panel_image()
        self.save_settings()

    def _get_baidu_client(self):
        if not REQUESTS_AVAILABLE:
            raise RuntimeError("requests库未安装")
        if not self.baidu_api_key or not self.baidu_secret_key:
            raise RuntimeError("百度OCR API未配置")
        if self._baidu_client is None:
            self._baidu_client = BaiduOCRClient(
                self.baidu_api_key, self.baidu_secret_key)
        return self._baidu_client

    # ==========================================================
    # 背景图片
    # ==========================================================

    def choose_background_image(self):
        filename = filedialog.askopenfilename(
            title="选择背景图片",
            filetypes=[("图片文件", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"),
                       ("所有文件", "*.*")]
        )
        if not filename:
            return
        if not PIL_AVAILABLE:
            messagebox.showerror(
                "错误", "Pillow库未安装，无法加载图片背景。\n请运行: pip install Pillow")
            return
        try:
            app_dir = get_app_dir()
            ext = os.path.splitext(filename)[1].lower() or ".png"
            target = os.path.join(app_dir, f"background{ext}")
            shutil.copyfile(filename, target)
            self.bg_image_path = target
            self.apply_background_image()
            self.save_settings()
        except Exception as e:
            messagebox.showerror("错误", f"无法设置背景图片：\n{str(e)}")

    def apply_background_image(self):
        if not PIL_AVAILABLE:
            return
        if not self.bg_image_path or not os.path.exists(self.bg_image_path):
            return
        try:
            img = Image.open(self.bg_image_path)
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            if width <= 1 or height <= 1:
                self.root.update_idletasks()
                width = self.root.winfo_width()
                height = self.root.winfo_height()
            img = img.resize((width, height), Image.LANCZOS).convert("RGB")
            self.bg_pil = img
            self.bg_image = ImageTk.PhotoImage(img)
            if self.bg_label is None:
                self.bg_label = tk.Label(self.root, image=self.bg_image)
                self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
                self.bg_label.lower()
            else:
                self.bg_label.configure(image=self.bg_image)
            self.root.after(0, self.apply_panel_image)
        except Exception as e:
            messagebox.showerror("错误", f"背景图片加载失败：\n{str(e)}")

    def on_root_resize(self, event):
        if not self.bg_image_path:
            return
        if self.resize_job is not None:
            try:
                self.root.after_cancel(self.resize_job)
            except Exception:
                pass
        self.resize_job = self.root.after(200, self.apply_background_image)

    def on_panel_resize(self, event):
        self.layout_canvas()
        if self.panel_resize_job is not None:
            try:
                self.root.after_cancel(self.panel_resize_job)
            except Exception:
                pass
        self.panel_resize_job = self.root.after(50, self.apply_panel_image)

    def refresh_layout(self):
        self.root.update_idletasks()
        self.layout_canvas()
        self.apply_panel_image()

    def apply_panel_image(self):
        if not PIL_AVAILABLE:
            return
        if not self.bg_pil or self.panel_canvas is None:
            return
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        panel_width = max(width - self.panel_padding * 2, 1)
        panel_height = max(height - self.panel_padding * 2, 1)
        if self.bg_pil.size[0] != width or self.bg_pil.size[1] != height:
            return
        left = self.panel_padding
        top = self.panel_padding
        right = left + panel_width
        bottom = top + panel_height
        panel_img = self.bg_pil.crop((left, top, right, bottom))
        opacity = max(0.2, min(1.0, self.panel_opacity_var.get() / 100.0))
        overlay = Image.new("RGB", panel_img.size, (255, 255, 255))
        panel_img = Image.blend(overlay, panel_img, opacity)
        self.panel_image = ImageTk.PhotoImage(panel_img)
        if self.panel_image_id is None:
            self.panel_image_id = self.panel_canvas.create_image(
                0, 0, anchor="nw", image=self.panel_image)
            self.panel_canvas.tag_lower(self.panel_image_id)
        else:
            self.panel_canvas.itemconfigure(
                self.panel_image_id, image=self.panel_image)
        self.panel_canvas.update_idletasks()

    # ==========================================================
    # 设置存取
    # ==========================================================

    def load_settings(self):
        if not os.path.exists(self.settings_path):
            return
        try:
            with open(self.settings_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            title_text = data.get('title_text')
            if title_text:
                self.title_text_var.set(title_text)
            bg_path = data.get('background_image')
            if bg_path and os.path.exists(bg_path):
                self.bg_image_path = bg_path
            opacity = data.get('panel_opacity', data.get('background_opacity'))
            if isinstance(opacity, (int, float)):
                self.panel_opacity_var.set(max(20.0, min(100.0, float(opacity))))
            # OCR和公式选项
            self.ocr_enabled_var.set(data.get('ocr_enabled', False))
            self.formula_api_enabled_var.set(data.get('formula_api_enabled', False))
            # API 配置
            self.baidu_api_key = simple_decrypt(data.get('baidu_api_key_enc', ''))
            self.baidu_secret_key = simple_decrypt(
                data.get('baidu_secret_key_enc', ''))
            self.xslt_path = data.get('xslt_path') or None
            # 功能选择和图片选项
            saved_func = data.get('current_function', 'PDF转Word')
            if saved_func in ALL_FUNCTIONS:
                self.current_function_var.set(saved_func)
                self._on_function_changed()
            saved_dpi = data.get('image_dpi', '200')
            if saved_dpi:
                self.image_dpi_var.set(str(saved_dpi))
            saved_fmt = data.get('image_format', 'PNG')
            if saved_fmt in ('PNG', 'JPEG'):
                self.image_format_var.set(saved_fmt)
            # 新增选项
            saved_split = data.get('split_mode', '每页一个PDF')
            if saved_split in ("每页一个PDF", "每N页一个PDF", "按范围拆分"):
                self.split_mode_var.set(saved_split)
            saved_page_size = data.get('page_size', 'A4')
            if saved_page_size in ("A4", "A3", "Letter", "Legal", "自适应"):
                self.page_size_var.set(saved_page_size)
            saved_excel_mode = data.get('excel_extract_mode', '结构提取')
            if saved_excel_mode in ("结构提取", "OCR提取"):
                self.excel_extract_mode_var.set(saved_excel_mode)
            # 批量提取选项
            self.batch_text_enabled_var.set(data.get('batch_text_enabled', True))
            self.batch_image_enabled_var.set(data.get('batch_image_enabled', True))
            self.batch_text_format_var.set(data.get('batch_text_format', 'txt'))
            self.batch_text_mode_var.set(data.get('batch_text_mode', '合并为一个文件'))
            self.batch_preserve_layout_var.set(data.get('batch_preserve_layout', True))
            self.batch_ocr_enabled_var.set(data.get('batch_ocr_enabled', False))
            self.batch_pages_var.set(data.get('batch_pages', ''))
            self.batch_image_per_page_var.set(data.get('batch_image_per_page', False))
            self.batch_image_dedupe_var.set(data.get('batch_image_dedupe', False))
            self.batch_image_format_var.set(data.get('batch_image_format', '原格式'))
            self.batch_zip_enabled_var.set(data.get('batch_zip_enabled', False))
            # 批量盖章选项
            saved_stamp_mode = data.get('stamp_mode', '普通章')
            if saved_stamp_mode in ("普通章", "二维码", "骑缝章", "模板"):
                self.stamp_mode_var.set(saved_stamp_mode)
            self.stamp_pages_var.set(data.get('stamp_pages', ''))
            self.stamp_opacity_var.set(str(data.get('stamp_opacity', '0.85')))
            self.stamp_position_var.set(data.get('stamp_position', '右下'))
            self.stamp_size_ratio_var.set(str(data.get('stamp_size_ratio', '0.18')))
            self.stamp_qr_text_var.set(data.get('stamp_qr_text', ''))
            self.stamp_seam_side_var.set(data.get('stamp_seam_side', '右侧'))
            self.stamp_seam_align_var.set(data.get('stamp_seam_align', '居中'))
            self.stamp_seam_overlap_var.set(str(data.get('stamp_seam_overlap', '0.25')))
            self.stamp_remove_white_bg_var.set(bool(data.get('stamp_remove_white_bg', False)))
            preview_profile = data.get('stamp_preview_profile', {})
            if isinstance(preview_profile, dict):
                self.stamp_preview_profile = {
                    "x_ratio": self._clamp_value(preview_profile.get("x_ratio", 0.85), 0.0, 1.0, 0.85),
                    "y_ratio": self._clamp_value(preview_profile.get("y_ratio", 0.85), 0.0, 1.0, 0.85),
                    "size_ratio": self._clamp_value(preview_profile.get("size_ratio", 0.18), 0.03, 0.7, 0.18),
                    "opacity": self._clamp_value(preview_profile.get("opacity", self.stamp_opacity_var.get()), 0.05, 1.0, 0.85),
                }
            else:
                self.stamp_preview_profile = {
                    "x_ratio": 0.85,
                    "y_ratio": 0.85,
                    "size_ratio": self._clamp_value(self.stamp_size_ratio_var.get(), 0.03, 0.7, 0.18),
                    "opacity": self._clamp_value(self.stamp_opacity_var.get(), 0.05, 1.0, 0.85),
                }
            self.stamp_opacity_var.set(f"{self.stamp_preview_profile.get('opacity', 0.85):.2f}")
            saved_image_paths = data.get('stamp_image_paths', [])
            if not isinstance(saved_image_paths, list):
                saved_image_paths = []
            if not saved_image_paths:
                legacy_single = data.get('stamp_image_path', '') or ''
                if legacy_single:
                    saved_image_paths = [legacy_single]
            self._set_stamp_images(saved_image_paths, selected_idx=data.get('stamp_selected_image_idx', 0))
            saved_profiles = data.get('stamp_profiles', {})
            if isinstance(saved_profiles, dict):
                loaded_profiles = {}
                for k, v in saved_profiles.items():
                    full = os.path.abspath(str(k))
                    if full in self.stamp_image_paths and isinstance(v, dict):
                        loaded_profiles[full] = self._normalize_stamp_profile(v)
                for full in self.stamp_image_paths:
                    loaded_profiles[full] = self._normalize_stamp_profile(loaded_profiles.get(full))
                self.stamp_profiles = loaded_profiles
            self.stamp_template_path = data.get('stamp_template_path', '') or ''
            if self.stamp_template_path and os.path.exists(self.stamp_template_path):
                nm2 = os.path.basename(self.stamp_template_path)
                self.stamp_template_label.config(text=nm2 if len(nm2) <= 16 else nm2[:13] + "...")
            if self.bg_image_path:
                self.apply_background_image()
            self._on_stamp_mode_changed()
            self._update_stamp_preview_info()
            self._update_api_hint()
        except Exception:
            pass

    def save_settings(self):
        data = {
            'title_text': self.title_text_var.get().strip(),
            'background_image': self.bg_image_path,
            'panel_opacity': float(self.panel_opacity_var.get()),
            'ocr_enabled': bool(self.ocr_enabled_var.get()),
            'formula_api_enabled': bool(self.formula_api_enabled_var.get()),
            'baidu_api_key_enc': simple_encrypt(self.baidu_api_key),
            'baidu_secret_key_enc': simple_encrypt(self.baidu_secret_key),
            'xslt_path': self.xslt_path or '',
            'current_function': self.current_function_var.get(),
            'image_dpi': self.image_dpi_var.get(),
            'image_format': self.image_format_var.get(),
            'split_mode': self.split_mode_var.get(),
            'page_size': self.page_size_var.get(),
            'excel_extract_mode': self.excel_extract_mode_var.get(),
            'batch_text_enabled': bool(self.batch_text_enabled_var.get()),
            'batch_image_enabled': bool(self.batch_image_enabled_var.get()),
            'batch_text_format': self.batch_text_format_var.get(),
            'batch_text_mode': self.batch_text_mode_var.get(),
            'batch_preserve_layout': bool(self.batch_preserve_layout_var.get()),
            'batch_ocr_enabled': bool(self.batch_ocr_enabled_var.get()),
            'batch_pages': self.batch_pages_var.get(),
            'batch_image_per_page': bool(self.batch_image_per_page_var.get()),
            'batch_image_dedupe': bool(self.batch_image_dedupe_var.get()),
            'batch_image_format': self.batch_image_format_var.get(),
            'batch_zip_enabled': bool(self.batch_zip_enabled_var.get()),
            'stamp_mode': self.stamp_mode_var.get(),
            'stamp_pages': self.stamp_pages_var.get(),
            'stamp_opacity': self.stamp_opacity_var.get(),
            'stamp_position': self.stamp_position_var.get(),
            'stamp_size_ratio': self.stamp_size_ratio_var.get(),
            'stamp_qr_text': self.stamp_qr_text_var.get(),
            'stamp_seam_side': self.stamp_seam_side_var.get(),
            'stamp_seam_align': self.stamp_seam_align_var.get(),
            'stamp_seam_overlap': self.stamp_seam_overlap_var.get(),
            'stamp_remove_white_bg': bool(self.stamp_remove_white_bg_var.get()),
            'stamp_preview_profile': {
                'x_ratio': self._clamp_value(self.stamp_preview_profile.get("x_ratio", 0.85), 0.0, 1.0, 0.85),
                'y_ratio': self._clamp_value(self.stamp_preview_profile.get("y_ratio", 0.85), 0.0, 1.0, 0.85),
                'size_ratio': self._clamp_value(self.stamp_preview_profile.get("size_ratio", 0.18), 0.03, 0.7, 0.18),
                'opacity': self._clamp_value(self.stamp_preview_profile.get("opacity", self.stamp_opacity_var.get()), 0.05, 1.0, 0.85),
            },
            'stamp_image_paths': list(self.stamp_image_paths),
            'stamp_selected_image_idx': int(self.stamp_selected_image_idx),
            'stamp_profiles': {
                p: self._normalize_stamp_profile(self.stamp_profiles.get(p))
                for p in self.stamp_image_paths if p
            },
            'stamp_image_path': self._get_active_stamp_image_path(),
            'stamp_template_path': self.stamp_template_path,
        }
        try:
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ==========================================================
    # 工具方法
    # ==========================================================

    def _parse_page_range_for_converter(self):
        """将UI的页范围文本转为 (start_page_0based, end_page_0based_exclusive) 或 (0, None)"""
        start_text = self.page_start_var.get().strip()
        end_text = self.page_end_var.get().strip()
        if not start_text and not end_text:
            return 0, None
        start_page = int(start_text) - 1 if start_text and start_text.isdigit() else 0
        end_page = int(end_text) if end_text and end_text.isdigit() else None
        start_page = max(0, start_page)
        # 验证起始页不超过结束页
        if end_page is not None and start_page >= end_page:
            return 0, None
        return start_page, end_page

    @staticmethod
    def format_skipped_pages(skipped_pages):
        pages = sorted(set(skipped_pages))
        if len(pages) <= 30:
            return ", ".join(str(p) for p in pages)
        head = ", ".join(str(p) for p in pages[:30])
        return f"{head} ...（共 {len(pages)} 页）"

    def generate_output_filename(self, input_file, extension):
        directory = os.path.dirname(input_file)
        basename = os.path.splitext(os.path.basename(input_file))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{basename}_converted_{timestamp}{extension}"
        return os.path.join(directory, output_filename)

    def open_folder(self, filepath):
        try:
            folder = os.path.dirname(os.path.abspath(filepath))
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹：\n{str(e)}")
