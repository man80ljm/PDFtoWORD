"""
对话框模块

包含设置窗口（外观、API）和转换历史记录窗口。
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox

from core.ocr_client import BaiduOCRClient


def open_settings_window(app):
    """打开设置窗口（含API配置）

    Args:
        app: PDFConverterApp 实例
    """
    win = tk.Toplevel(app.root)
    win.title("设置")
    win.geometry("480x520")
    win.resizable(False, False)

    # 使用 Notebook 分页签
    notebook = ttk.Notebook(win)
    notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    # ========== 页签1：外观设置 ==========
    tab_appearance = tk.Frame(notebook, padx=12, pady=12)
    notebook.add(tab_appearance, text="外观设置")

    tk.Label(tab_appearance, text="标题文字:", font=("Microsoft YaHei", 10)).pack(anchor=tk.W)
    title_entry = tk.Entry(tab_appearance, textvariable=app.title_text_var,
                           font=("Microsoft YaHei", 10))
    title_entry.pack(fill=tk.X, pady=(4, 12))

    tk.Button(tab_appearance, text="更换背景", font=("Microsoft YaHei", 10),
              command=app.choose_background_image).pack(anchor=tk.W)

    tk.Label(tab_appearance, text="面板透明度:", font=("Microsoft YaHei", 10)
             ).pack(anchor=tk.W, pady=(12, 0))
    tk.Scale(tab_appearance, from_=0, to=100, orient=tk.HORIZONTAL,
             resolution=1, showvalue=True, variable=app.panel_opacity_var,
             command=app.on_opacity_change).pack(fill=tk.X, pady=(4, 0))

    tk.Button(tab_appearance, text="应用标题", font=("Microsoft YaHei", 10),
              command=app.apply_title_text).pack(anchor=tk.W, pady=(12, 0))

    # ========== 页签2：API设置 ==========
    tab_api = tk.Frame(notebook, padx=12, pady=12)
    notebook.add(tab_api, text="API设置")

    # 百度OCR配置
    tk.Label(tab_api, text="百度OCR API（用于文字识别和公式识别）",
             font=("Microsoft YaHei", 10, "bold")).pack(anchor=tk.W, pady=(0, 8))

    tk.Label(tab_api, text="API Key:", font=("Microsoft YaHei", 9)).pack(anchor=tk.W)
    api_key_var = tk.StringVar(value=app.baidu_api_key)
    tk.Entry(tab_api, textvariable=api_key_var, font=("Microsoft YaHei", 9),
             width=50).pack(fill=tk.X, pady=(2, 6))

    tk.Label(tab_api, text="Secret Key:", font=("Microsoft YaHei", 9)).pack(anchor=tk.W)
    secret_key_var = tk.StringVar(value=app.baidu_secret_key)
    tk.Entry(tab_api, textvariable=secret_key_var, font=("Microsoft YaHei", 9),
             width=50, show="*").pack(fill=tk.X, pady=(2, 8))

    # 测试连接
    test_status_var = tk.StringVar(value="")
    test_frame = tk.Frame(tab_api)
    test_frame.pack(fill=tk.X, pady=(0, 8))

    def do_test():
        ak = api_key_var.get().strip()
        sk = secret_key_var.get().strip()
        if not ak or not sk:
            test_status_var.set("⚠ 请填写API Key和Secret Key")
            return
        test_status_var.set("⏳ 正在测试...")
        test_btn.config(state=tk.DISABLED)

        def _test_thread():
            client = BaiduOCRClient(ak, sk)
            ok, msg = client.test_connection()
            def _update():
                test_btn.config(state=tk.NORMAL)
                if ok:
                    test_status_var.set("✅ 连接成功")
                else:
                    test_status_var.set(f"❌ 失败: {msg[:50]}")
            win.after(0, _update)

        threading.Thread(target=_test_thread, daemon=True).start()

    test_btn = tk.Button(test_frame, text="测试连接", font=("Microsoft YaHei", 9),
              command=do_test)
    test_btn.pack(side=tk.LEFT)
    tk.Label(test_frame, textvariable=test_status_var,
             font=("Microsoft YaHei", 9)).pack(side=tk.LEFT, padx=(10, 0))

    # 说明
    hint_text = (
        "注册地址：https://cloud.baidu.com/product/ocr\n"
        "1. 注册百度智能云账号\n"
        "2. 创建文字识别应用，获取API Key和Secret Key\n"
        "3. 同一个应用可同时使用文字识别和公式识别\n"
        "4. 免费额度：通用文字500次/月"
    )
    tk.Label(tab_api, text=hint_text, font=("Microsoft YaHei", 8),
             fg="#666666", justify=tk.LEFT, wraplength=420).pack(anchor=tk.W, pady=(4, 12))

    # XSLT路径（高级选项）
    tk.Label(tab_api, text="高级选项（通常无需修改）:",
             font=("Microsoft YaHei", 8), fg="#aaaaaa").pack(anchor=tk.W, pady=(8, 0))
    xslt_hint = "留空自动检测Office安装路径，仅Office路径异常时手动填写"
    tk.Label(tab_api, text=f"MML2OMML.XSL: {xslt_hint}",
             font=("Microsoft YaHei", 8), fg="#aaaaaa").pack(anchor=tk.W)
    xslt_var = tk.StringVar(value=app.xslt_path or "")
    tk.Entry(tab_api, textvariable=xslt_var, font=("Microsoft YaHei", 8),
             fg="#aaaaaa").pack(fill=tk.X, pady=(2, 0))

    # 保存按钮
    def save_api_settings():
        app.baidu_api_key = api_key_var.get().strip()
        app.baidu_secret_key = secret_key_var.get().strip()
        app.xslt_path = xslt_var.get().strip() or None
        app._baidu_client = None  # 重建客户端
        app.save_settings()
        app._update_api_hint()
        messagebox.showinfo("设置", "API设置已保存", parent=win)

    tk.Button(tab_api, text="保存设置", font=("Microsoft YaHei", 10, "bold"),
              command=save_api_settings).pack(anchor=tk.E, pady=(12, 0))


def open_history_window(app):
    """打开转换历史记录窗口

    Args:
        app: PDFConverterApp 实例
    """
    win = tk.Toplevel(app.root)
    win.title("转换历史记录")
    win.geometry("650x420")
    win.resizable(True, True)
    win.transient(app.root)

    # 工具栏
    toolbar = tk.Frame(win)
    toolbar.pack(fill=tk.X, padx=10, pady=5)
    count_label = tk.Label(
        toolbar, text=f"共 {app.history.count} 条记录",
        font=("Microsoft YaHei", 9)
    )
    count_label.pack(side=tk.LEFT)
    tk.Button(
        toolbar, text="清空历史", font=("Microsoft YaHei", 9),
        command=lambda: _clear_history()
    ).pack(side=tk.RIGHT)

    # 列表
    columns = ('time', 'function', 'files', 'result', 'pages')
    tree_frame = tk.Frame(win)
    tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

    tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
    tree.heading('time', text='时间')
    tree.heading('function', text='功能')
    tree.heading('files', text='文件')
    tree.heading('result', text='结果')
    tree.heading('pages', text='页数')

    tree.column('time', width=130, minwidth=100)
    tree.column('function', width=80, minwidth=60)
    tree.column('files', width=240, minwidth=120)
    tree.column('result', width=100, minwidth=60)
    tree.column('pages', width=60, minwidth=40)

    scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)

    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _load_records():
        for item in tree.get_children():
            tree.delete(item)
        for record in app.history.get_all():
            files = record.get('input_files', [])
            if files:
                file_text = os.path.basename(files[0])
                if len(files) > 1:
                    file_text += f" 等{len(files)}个"
            else:
                file_text = ""
            result_text = "✅ 成功" if record.get('success') else "❌ 失败"
            tree.insert('', 'end', values=(
                record.get('timestamp', ''),
                record.get('function', ''),
                file_text,
                result_text,
                record.get('page_count', 0),
            ))

    def _clear_history():
        if messagebox.askyesno("确认", "确定要清空所有转换历史记录吗？",
                               parent=win):
            app.history.clear()
            _load_records()
            count_label.config(text="共 0 条记录")

    _load_records()
