# 📄 PDF转换工具 v1.0.0 —— PDF全能工具箱

> 🎯 一个长得好看、用着顺手、功能全面的 PDF 工具箱。  
> 支持 PDF 转 Word、转图片、合并、拆分、图片转 PDF、加水印、加密/解密，一站式搞定！

![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%207%2F10%2F11-0078D6?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?logo=opensourceinitiative&logoColor=white)
![GUI](https://img.shields.io/badge/GUI-tkinter-orange?logo=buffer&logoColor=white)
![Status](https://img.shields.io/badge/Status-v1.0.0%20%E7%A8%B3%E5%AE%9A%E5%8F%AF%E7%94%A8-brightgreen)

---

## ✨ 功能亮点

| 功能 | 说明 | 状态 |
|------|------|------|
| 📝 PDF → Word | 高质量转换，保留排版；支持批量多文件 | ✅ |
| 🔍 OCR + 公式识别 | 扫描件/数学公式通过百度云API智能识别 | ✅ |
| 🖼️ PDF → 图片 | 批量导出，可选DPI和格式(PNG/JPEG) | ✅ |
| 📎 PDF 合并 | 多个PDF按顺序合并为一个 | ✅ |
| ✂️ PDF 拆分 | 按页/按N页/按范围三种拆分模式 | ✅ |
| 🎨 图片 → PDF | 多张图片合成PDF，支持多种页面尺寸 | ✅ |
| 💧 PDF 加水印 | 文字/图片水印，支持透明度、平铺/居中/角落定位 | ✅ |
| 🔒 PDF 加密/解密 | AES-256加密，可设打印/复制/修改/批注权限；一键解密 | ✅ |
| 📐 页范围选择 | 只转你想转的那几页 | ✅ |
| 🔀 文件排序 | 合并/图片转PDF时可自由调整文件顺序 | ✅ |
| 📊 实时进度条 | 显示百分比、当前页码、预计剩余时间 | ✅ |
| 📋 转换历史 | 自动记录每次转换操作 | ✅ |
| 🖱️ 拖拽导入 | 直接把文件拖进窗口 | ✅ |
| 🖼️ 自定义背景 | 换张好看的壁纸，透明面板颜值拉满 | ✅ |
| ✏️ 自定义标题 | 想叫啥名就叫啥名 | ✅ |
| 💾 设置持久化 | 关掉再开，所有设置都还在 | ✅ |

---

## 🚀 快速开始

### 方式一：直接用打包好的 exe（推荐懒人使用 🦥）

双击 `dist/PDF转换工具.exe`，开箱即用，不需要装 Python。

> ⚡ 首次启动可能稍慢（几秒），别急，它在解压自己。

### 方式二：从源码运行（适合爱折腾的你 🔧）

```bash
# 1. 克隆项目
git clone https://github.com/your-username/PDFtoWORD.git
cd PDFtoWORD

# 2. 创建虚拟环境（推荐 Python 3.8，兼容 Win7）
py -3.8 -m venv .venv38
.venv38\Scripts\activate

# 3. 安装依赖
pip install -r requirements.txt

# 4. 启动
python pdf_converter.py
```

---

## 📁 项目结构

```
PDFtoWORD/
├── 📄 pdf_converter.py          # 入口文件（启动应用）
├── 📂 core/                     # 核心模块
│   ├── __init__.py              # 公共工具函数
│   ├── math_utils.py            # LaTeX/MathML 公式处理
│   ├── ocr_client.py            # 百度OCR API客户端
│   ├── progress_converter.py    # pdf2docx进度回调封装
│   └── history.py               # 转换历史记录管理
├── 📂 converters/               # 转换器（每种功能一个文件）
│   ├── __init__.py
│   ├── pdf_to_word.py           # PDF → Word
│   ├── pdf_to_image.py          # PDF → 图片
│   ├── pdf_merge.py             # PDF 合并
│   ├── pdf_split.py             # PDF 拆分
│   ├── image_to_pdf.py          # 图片 → PDF
│   ├── pdf_watermark.py         # PDF 加水印
│   └── pdf_encrypt.py           # PDF 加密/解密
├── 📂 ui/                       # 界面层
│   ├── __init__.py
│   ├── app.py                   # 主应用类（UI + 事件 + 设置）
│   └── dialogs.py               # 设置窗口 + 历史窗口
├── 📋 requirements.txt          # 依赖清单
├── 📘 README.md                 # 项目说明
├── 📝 更新说明.md               # 版本更新记录
├── 📝 打包说明.md               # 打包相关说明
├── 🧪 create_test_pdf.py        # 测试PDF生成器
└── 📜 LICENSE                   # MIT许可证
```

---

## 🛠️ 技术栈

| 技术 | 用途 |
|------|------|
| 🐍 Python 3.8 | 主语言，支持 Win7 |
| 🖼️ tkinter | GUI 框架（Python 自带） |
| 📄 pdf2docx | PDF 转 Word 核心引擎 |
| 📑 PyMuPDF (fitz) | PDF 读写/图片/水印/加密 |
| 📝 python-docx | Word文档操作（OCR模式使用） |
| 🎨 Pillow | 图片处理 + 背景图渲染 + 水印透明度 |
| 🌐 requests | 百度云OCR API调用 |
| 🧮 latex2mathml + lxml | LaTeX公式 → MathML转换 |
| 🖱️ windnd | 文件拖拽支持 |
| 📦 PyInstaller | 打包为 exe |

---

## 📦 依赖清单

```
pdf2docx>=0.5.6    # PDF → Word 核心
PyMuPDF>=1.23.0    # PDF处理（图片/水印/加密/拆分/合并）
python-docx>=0.8.11 # Word文档操作（OCR模式写入）
Pillow>=9.0.0      # 图片处理 + 水印透明度
requests>=2.28.0   # HTTP请求（OCR API）
latex2mathml>=3.0  # LaTeX → MathML
lxml>=4.9.0        # XML处理
windnd>=1.0.7      # 窗口拖拽
```

---

## ⚙️ 设置说明

点击界面左上角 **⚙** 齿轮按钮：

- **标题文字**：修改主界面大标题
- **更换背景**：选一张图片作为背景
- **面板透明度**：0%~100%，越低越透明
- **百度API配置**：设置API Key和Secret Key以启用OCR/公式识别
- **测试连接**：验证API配置是否正确

点击 **📋** 按钮查看转换历史记录。

> 设置保存在程序目录下的 `settings.json`，删掉它即恢复默认。

---

## 🤔 常见问题

<details>
<summary><b>Q: 转换出来格式不对怎么办？</b></summary>
PDF 转 Word 本身是"尽力而为"。简单排版效果很好，复杂表格嵌套可能有偏差，建议转完后微调。
</details>

<details>
<summary><b>Q: 扫描版 PDF 怎么转？</b></summary>
勾选"OCR识别(扫描件)"，需要先在设置中配置百度云OCR的 API Key 和 Secret Key。
</details>

<details>
<summary><b>Q: 数学公式能识别吗？</b></summary>
可以！勾选"公式智能识别"，同样需要配置百度云API。识别后自动转为 Word 可编辑的 MathML 公式。
</details>

<details>
<summary><b>Q: 转换大文件好慢？</b></summary>
进度条会告诉你预计剩余时间。超过 60 秒的页面会提示"该页复杂请耐心等待"。
</details>

<details>
<summary><b>Q: exe 在 Win7 上报错 api-ms-win-core-path 缺失？</b></summary>
请用 Python 3.8 的虚拟环境重新打包。Python 3.9+ 不支持 Win7。
</details>

<details>
<summary><b>Q: 可以批量转换吗？</b></summary>
可以！PDF转Word和PDF转图片都支持多文件选择/拖拽批量转换。
</details>

---

## 📜 许可证

MIT License — 随便用，随便改 ☕

---

<p align="center">
  <b>用 ❤️ 和 ☕ 驱动开发</b><br>
  <i>如果觉得好用，给个 ⭐ 呗~</i>
</p>
