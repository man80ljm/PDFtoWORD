"""
PDF转Word转换工具
使用tkinter构建的图形界面应用程序
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
import threading
import logging
import time
import json
import shutil
import sys

# PDF转换相关库
try:
    from pdf2docx import Converter
    from pdf2docx.converter import ConversionException, MakedocxException
    from docx import Document
    PDF2DOCX_AVAILABLE = True

    class ProgressConverter(Converter):
        """带进度回调的PDF转Word转换器"""

        def __init__(self, pdf_file: str = None, password: str = None, stream: bytes = None, progress_callback=None):
            super().__init__(pdf_file=pdf_file, password=password, stream=stream)
            self.progress_callback = progress_callback
            self.skipped_pages = set()

        def _notify(self, phase: str, current: int, total: int, page_id: int):
            if self.progress_callback:
                self.progress_callback(phase, current, total, page_id)

        def parse_pages(self, **kwargs):
            """解析页面并回调进度"""
            logging.info(self._color_output('[3/4] Parsing pages...'))

            pages = [page for page in self._pages if not page.skip_parsing]
            total_pages = len(self._pages)
            num_pages = len(pages)
            for i, page in enumerate(pages, start=1):
                pid = page.id + 1
                self._notify('start-parse', i, num_pages, pid)
                logging.info('(%d/%d) Page %d', i, num_pages, pid)
                try:
                    page.parse(**kwargs)
                except Exception as e:
                    if not kwargs['debug'] and kwargs['ignore_page_error']:
                        logging.error('Ignore page %d due to parsing page error: %s', pid, e)
                        self.skipped_pages.add(pid)
                        self._notify('skip-parse', i, num_pages, pid)
                    else:
                        raise ConversionException(f'Error when parsing page {pid}: {e}')
                finally:
                    self._notify('parse', i, num_pages, pid)

            return self

        def make_docx(self, filename_or_stream=None, **kwargs):
            """生成docx并回调进度"""
            logging.info(self._color_output('[4/4] Creating pages...'))

            parsed_pages = list(filter(lambda page: page.finalized, self._pages))
            if not parsed_pages:
                raise ConversionException('No parsed pages. Please parse page first.')

            if not filename_or_stream:
                if self.filename_pdf:
                    filename_or_stream = f'{self.filename_pdf[0:-len(".pdf")]}.docx'
                    if os.path.exists(filename_or_stream):
                        os.remove(filename_or_stream)
                else:
                    raise ConversionException('Please specify a docx file name or a file-like object to write.')

            docx_file = Document()
            num_pages = len(parsed_pages)
            for i, page in enumerate(parsed_pages, start=1):
                if not page.finalized:
                    continue
                pid = page.id + 1
                self._notify('start-make', i, num_pages, pid)
                logging.info('(%d/%d) Page %d', i, num_pages, pid)
                try:
                    page.make_docx(docx_file)
                except Exception as e:
                    if not kwargs['debug'] and kwargs['ignore_page_error']:
                        logging.error('Ignore page %d due to making page error: %s', pid, e)
                        self.skipped_pages.add(pid)
                        self._notify('skip-make', i, num_pages, pid)
                    else:
                        raise MakedocxException(f'Error when make page {pid}: {e}')
                finally:
                    self._notify('make', i, num_pages, pid)

            docx_file.save(filename_or_stream)
except ImportError:
    PDF2DOCX_AVAILABLE = False
    ProgressConverter = None

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

class PDFConverterApp:
    """PDF转换工具主应用类"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("程新伟专属转换器 - PDF转Word")
        self.root.geometry("500x550")
        self.root.resizable(False, False)
        
        # 设置应用图标（如果可用）
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
        
        # 变量
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
        self.page_start_var = tk.StringVar()
        self.page_end_var = tk.StringVar()
        self.title_text_var = tk.StringVar(value="程新伟专属转换器")
        self.settings_path = os.path.join(self.get_app_dir(), "settings.json")
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
        
        # 创建UI
        self.create_ui()

        # 加载设置
        self.load_settings()
        
        # 检查依赖
        self.check_dependencies()
    
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
        self.cv_settings = self.panel_canvas.create_window(5, 5, window=self.settings_btn, anchor="nw")

        # 标题（透明背景）
        self.cv_title = self.panel_canvas.create_text(
            0, 35, text=self.title_text_var.get(),
            font=("Microsoft YaHei", 26, "bold"), anchor="n"
        )
        self.title_text_var.trace_add("write", self._on_title_var_changed)

        # 副标题（透明背景）
        self.cv_subtitle = self.panel_canvas.create_text(
            0, 75, text="PDF转Word工具（支持大文件）",
            font=("Microsoft YaHei", 10), anchor="n"
        )

        # 分区标题：选择PDF文件（透明背景）
        self.cv_section1 = self.panel_canvas.create_text(
            15, 105, text="选择PDF文件",
            font=("Microsoft YaHei", 11, "bold"), anchor="nw"
        )

        # 文件输入框 + 浏览按钮
        file_frame = tk.Frame(self.panel_canvas)
        self.file_entry = tk.Entry(
            file_frame, textvariable=self.selected_file,
            font=("Microsoft YaHei", 10), state='readonly'
        )
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        tk.Button(
            file_frame, text="浏览...", command=self.browse_file,
            font=("Microsoft YaHei", 10), padx=20, cursor='hand2'
        ).pack(side=tk.LEFT, padx=(10, 0), ipady=6)
        self.cv_file_frame = self.panel_canvas.create_window(
            15, 130, window=file_frame, anchor="nw", width=1
        )

        # 分区标题：页范围（透明背景）
        self.cv_section2 = self.panel_canvas.create_text(
            15, 185, text="页范围（可选）",
            font=("Microsoft YaHei", 11, "bold"), anchor="nw"
        )

        # 页范围输入
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

        # 进度条
        self.progress_bar = ttk.Progressbar(self.panel_canvas, mode='determinate')
        self.cv_progress_bar = self.panel_canvas.create_window(
            20, 260, window=self.progress_bar, anchor="nw", width=1, height=25
        )

        # 进度文本（透明背景）
        self.cv_progress_text = self.panel_canvas.create_text(
            0, 295, text="", font=("Microsoft YaHei", 9), anchor="n"
        )

        # 转换 / 清除按钮
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
            0, 350, window=btn_frame, anchor="n"
        )

        # 状态栏文字（透明背景）
        self.cv_status_text = self.panel_canvas.create_text(
            15, 0, text=self.status_message.get(),
            font=("Microsoft YaHei", 9), anchor="sw"
        )
        self.status_message.trace_add("write", self._on_status_var_changed)

        # 绑定事件
        self.root.bind("<Configure>", self.on_root_resize)
        self.panel_canvas.bind("<Configure>", self.on_panel_resize)
        self.root.after(50, self.refresh_layout)
    
    def _on_title_var_changed(self, *args):
        """标题变量变化时更新Canvas文字"""
        if self.panel_canvas:
            self.panel_canvas.itemconfigure(self.cv_title, text=self.title_text_var.get())

    def _on_status_var_changed(self, *args):
        """状态变量变化时更新Canvas文字"""
        if self.panel_canvas:
            self.panel_canvas.itemconfigure(self.cv_status_text, text=self.status_message.get())

    def set_progress_text(self, text):
        """更新进度文本"""
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
        self.panel_canvas.coords(self.cv_progress_bar, 20, 260)
        self.panel_canvas.itemconfigure(self.cv_progress_bar, width=w - 40)
        self.panel_canvas.coords(self.cv_progress_text, cx, 295)
        self.panel_canvas.coords(self.cv_btn_frame, cx, 350)
        self.panel_canvas.coords(self.cv_status_text, 15, h - 10)
    
    def check_dependencies(self):
        """检查依赖库"""
        missing = []
        
        if not PDF2DOCX_AVAILABLE:
            missing.append("pdf2docx")
        if missing:
            msg = f"警告：以下依赖库未安装：\n{', '.join(missing)}\n\n请运行: pip install {' '.join(missing)}"
            self.status_message.set(f"缺少依赖库: {', '.join(missing)}")
            messagebox.showwarning("缺少依赖", msg)
    
    def browse_file(self):
        """浏览并选择PDF文件"""
        filename = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if filename:
            self.selected_file.set(filename)
            self.status_message.set(f"已选择: {os.path.basename(filename)}")
    
    def clear_selection(self):
        """清除选择"""
        self.selected_file.set("")
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
    
    def start_conversion(self):
        """开始转换"""
        if not self.selected_file.get():
            messagebox.showwarning("提示", "请先选择一个PDF文件！")
            return
        
        if not os.path.exists(self.selected_file.get()):
            messagebox.showerror("错误", "选择的文件不存在！")
            return
        
        # 禁用转换按钮
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
        
        # 在新线程中执行转换
        thread = threading.Thread(target=self.perform_conversion)
        thread.daemon = True
        thread.start()
    
    def perform_conversion(self):
        """执行转换（在后台线程中）"""
        try:
            self.convert_to_word()
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("转换失败", f"转换过程中出错：\n{str(e)}"))
            self.root.after(0, lambda: self.status_message.set("转换失败"))
        finally:
            # 重新启用转换按钮
            self.conversion_active = False
            self.stop_page_timer()
            self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))
    
    def convert_to_word(self):
        """将PDF转换为Word"""
        if not PDF2DOCX_AVAILABLE or ProgressConverter is None:
            self.root.after(0, lambda: messagebox.showerror("错误", "pdf2docx库未安装！\n请运行: pip install pdf2docx"))
            return
        
        # 更新状态
        self.base_status_text = "正在初始化转换..."
        self.root.after(0, self.apply_status_text)
        self.root.after(0, lambda: self.set_progress_text("准备中..."))
        
        # 生成输出文件名
        input_file = self.selected_file.get()
        output_file = self.generate_output_filename(input_file, '.docx')
        
        # 执行转换
        self.root.after(0, lambda: self.progress_bar.config(mode='determinate', maximum=100, value=0))
        
        try:
            cv = ProgressConverter(input_file, progress_callback=self.update_progress)
            self.total_pages = len(cv.fitz_doc)
            if self.total_pages <= 0:
                raise ConversionException("无法读取PDF页数")
            try:
                start_idx, end_idx, range_total = self.get_page_range(self.total_pages)
            except ValueError as e:
                self.root.after(0, lambda: messagebox.showerror("页范围错误", str(e)))
                self.root.after(0, lambda: self.status_message.set("页范围无效"))
                cv.close()
                return
            self.total_steps = range_total * 2
            self.start_time = time.time()
            self.root.after(0, lambda: self.set_progress_text(f"共 {range_total} 页，开始转换..."))
            cv.convert(output_file, start=start_idx, end=end_idx)
            cv.close()
            
            # 转换成功
            self.root.after(0, lambda: self.progress_bar.config(value=100))
            self.root.after(0, lambda: self.set_progress_text("转换完成！(100%)"))
            self.root.after(0, lambda: self.status_message.set(f"已保存到: {output_file}"))
            
            # 提示用户
            result = messagebox.askyesno(
                "转换成功",
                f"PDF已成功转换为Word！\n\n保存位置：\n{output_file}\n\n是否打开文件所在文件夹？"
            )
            
            if result:
                self.open_folder(output_file)

            if cv.skipped_pages:
                skipped_text = self.format_skipped_pages(cv.skipped_pages)
                messagebox.showwarning("跳过异常页", f"以下页面在转换中被跳过：\n{skipped_text}")
                
        except Exception as e:
            raise e

    def update_progress(self, phase: str, current: int, total: int, page_id: int):
        """更新进度条和提示信息"""
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
            self.base_status_text = f"正在{phase_text}第 {page_id} 页，共 {total} 页"
            self.root.after(0, self.apply_status_text)
            return

        if phase in ('skip-parse', 'skip-make'):
            phase_text = "解析" if phase == 'skip-parse' else "生成"
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
        self.base_status_text = f"正在{phase_text}第 {page_id} 页，共 {total} 页"

        eta_text = ""
        if self.start_time and completed_steps > 0:
            elapsed = time.time() - self.start_time
            remaining_steps = max(total_steps - completed_steps, 0)
            eta_seconds = int(round(elapsed * remaining_steps / completed_steps))
            eta_text = f"，预计剩余 {self.format_eta(eta_seconds)}"
        self.current_eta_text = eta_text

        def _apply():
            self.progress_bar.config(mode='determinate', maximum=100)
            self.progress_bar['value'] = percent
            self.set_progress_text(f"{page_text} ({percent}%)")
            self.apply_status_text()

        self.root.after(0, _apply)

    @staticmethod
    def format_eta(seconds: int) -> str:
        """格式化预计剩余时间"""
        minutes, sec = divmod(max(seconds, 0), 60)
        hours, minutes = divmod(minutes, 60)
        if hours > 0:
            return f"{hours}小时{minutes}分{sec}秒"
        if minutes > 0:
            return f"{minutes}分{sec}秒"
        return f"{sec}秒"

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

    def apply_status_text(self):
        text = self.base_status_text or ""
        if self.current_eta_text:
            text += self.current_eta_text
        if self.page_start_time:
            elapsed = int(time.time() - self.page_start_time)
            text += f"，当前页耗时 {self.format_eta(elapsed)}"
            if elapsed >= self.page_timeout_seconds:
                text += "，该页复杂请耐心等待"
        if text:
            self.status_message.set(text)

    def format_page_text(self, phase_text: str, current: int, total: int, page_id: int) -> str:
        if self.total_pages and total != self.total_pages:
            return f"{phase_text}页 {current}/{total} (原页 {page_id})"
        return f"{phase_text}页 {page_id}/{total}"

    def open_settings_window(self):
        """打开设置窗口"""
        win = tk.Toplevel(self.root)
        win.title("设置")
        win.geometry("360x260")
        win.resizable(False, False)

        container = tk.Frame(win, padx=16, pady=16)
        container.pack(fill=tk.BOTH, expand=True)

        title_label = tk.Label(container, text="标题文字:", font=("Microsoft YaHei", 10))
        title_label.pack(anchor=tk.W)

        title_entry = tk.Entry(container, textvariable=self.title_text_var, font=("Microsoft YaHei", 10))
        title_entry.pack(fill=tk.X, pady=(4, 12))

        bg_btn = tk.Button(
            container,
            text="更换背景",
            font=("Microsoft YaHei", 10),
            command=self.choose_background_image
        )
        bg_btn.pack(anchor=tk.W)

        opacity_label = tk.Label(container, text="面板透明度:", font=("Microsoft YaHei", 10))
        opacity_label.pack(anchor=tk.W, pady=(12, 0))

        opacity_scale = tk.Scale(
            container,
            from_=0,
            to=100,
            orient=tk.HORIZONTAL,
            resolution=1,
            showvalue=True,
            variable=self.panel_opacity_var,
            command=self.on_opacity_change
        )
        opacity_scale.pack(fill=tk.X, pady=(4, 0))

        apply_btn = tk.Button(
            container,
            text="应用标题",
            font=("Microsoft YaHei", 10),
            command=self.apply_title_text
        )
        apply_btn.pack(anchor=tk.W, pady=(12, 0))

    def apply_title_text(self):
        text = self.title_text_var.get().strip() or "程新伟专属转换器"
        self.title_text_var.set(text)
        self.save_settings()

    def on_opacity_change(self, _value=None):
        self.apply_panel_image()
        self.save_settings()

    def choose_background_image(self):
        filename = filedialog.askopenfilename(
            title="选择背景图片",
            filetypes=[("图片文件", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"), ("所有文件", "*.*")]
        )
        if not filename:
            return

        if not PIL_AVAILABLE:
            messagebox.showerror("错误", "Pillow库未安装，无法加载图片背景。\n请运行: pip install Pillow")
            return

        try:
            app_dir = self.get_app_dir()
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
            self.panel_image_id = self.panel_canvas.create_image(0, 0, anchor="nw", image=self.panel_image)
            self.panel_canvas.tag_lower(self.panel_image_id)
        else:
            self.panel_canvas.itemconfigure(self.panel_image_id, image=self.panel_image)
        self.panel_canvas.update_idletasks()

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
            if self.bg_image_path:
                self.apply_background_image()
        except Exception:
            pass

    def save_settings(self):
        data = {
            'title_text': self.title_text_var.get().strip(),
            'background_image': self.bg_image_path,
            'panel_opacity': float(self.panel_opacity_var.get())
        }
        try:
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    @staticmethod
    def get_app_dir():
        if getattr(sys, 'frozen', False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def get_page_range(self, total_pages: int):
        start_text = self.page_start_var.get().strip()
        end_text = self.page_end_var.get().strip()

        if not start_text and not end_text:
            return 0, None, total_pages

        if start_text and not start_text.isdigit():
            raise ValueError("起始页必须是数字")
        if end_text and not end_text.isdigit():
            raise ValueError("结束页必须是数字")

        start_page = int(start_text) if start_text else 1
        end_page = int(end_text) if end_text else total_pages

        if start_page < 1 or end_page < 1:
            raise ValueError("页码必须从1开始")
        if start_page > end_page:
            raise ValueError("起始页不能大于结束页")
        if end_page > total_pages:
            raise ValueError("结束页超出总页数")

        start_idx = start_page - 1
        end_idx = end_page
        return start_idx, end_idx, end_page - start_idx

    @staticmethod
    def format_skipped_pages(skipped_pages):
        pages = sorted(set(skipped_pages))
        if len(pages) <= 30:
            return ", ".join(str(p) for p in pages)
        head = ", ".join(str(p) for p in pages[:30])
        return f"{head} ...（共 {len(pages)} 页）"
    
    def generate_output_filename(self, input_file, extension):
        """生成输出文件名"""
        # 获取输入文件的目录和基本名称
        directory = os.path.dirname(input_file)
        basename = os.path.splitext(os.path.basename(input_file))[0]
        
        # 添加时间戳避免覆盖
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{basename}_converted_{timestamp}{extension}"
        
        return os.path.join(directory, output_filename)
    
    def open_folder(self, filepath):
        """打开文件所在文件夹"""
        try:
            folder = os.path.dirname(os.path.abspath(filepath))
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹：\n{str(e)}")


def main():
    """主函数"""
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
