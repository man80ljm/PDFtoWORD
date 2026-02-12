# PDF转换工具

一个基于 `tkinter` 的桌面 PDF 工具箱。  
一句话：这不是“能不能处理 PDF”，而是“你想怎么处理，我都给你安排明白”。

## 已实现功能

- 文档转换：`PDF转Word`、`PDF转图片`、`图片转PDF`、`PDF转Excel`
- 文档组织：`PDF合并`、`PDF拆分`
- 页面级处理：`PDF提取/删页`、`PDF页面重排/旋转/倒序`（含拖拽预览）
- OCR能力：`OCR可搜索PDF`（扫描件转可搜索文本）
- 水印系统：`PDF加水印`（文字/图片、角度/透明度、平铺与单点、预览后应用）
- 权限安全：`PDF加密/解密`（打开密码与权限控制）
- 体积优化：`PDF压缩`（多档压缩策略）
- 批量提取：`PDF批量文本/图片提取`
- 批量盖章：`PDF批量盖章`（普通章/骑缝章/二维码/模板/签名）
- 书签处理：`PDF添加/移除书签`（添加、删除、导入/导出 JSON、自动生成）

## 功能亮点（硬核版）

- 固定窗口布局下做了多行选项自适应，复杂参数不再挤爆界面。
- 关键功能提供预览交互再落地，减少“导出后才发现不对”的返工。
- 页码范围支持混合格式：`1,2,5-8`，批量场景更稳。
- 批处理能力覆盖提取、盖章、签名、页面处理等高频办公链路。
- `onefile` 打包可直接分发，开箱就干活，不跟环境扯皮。
- 目标很直接：把“PDF工具箱”做成“PDF生产力武器”。

## 运行环境

- Windows 10/11（推荐）
- Python 3.8+

## 本地运行

```powershell
py -3.8 -m venv .venv38; .\.venv38\Scripts\Activate.ps1; pip install -r requirements.txt; python pdf_converter.py
```

## 打包 Onefile EXE

推荐使用项目内 `spec` 文件打包（资源和隐藏依赖已配置）。

```powershell
.\.venv38\Scripts\python.exe -m PyInstaller --noconfirm --clean "PDF转换工具.spec"
```

打包完成后可执行文件在：

```text
dist/PDF转换工具.exe
```

## 打包说明

- 首次启动 `onefile` 会先自解压，启动慢几秒属于正常现象。
- 如需修改图标，请替换项目根目录 `logo.ico` 后重新打包。
- 若新增第三方库导致启动报错，请在 `PDF转换工具.spec` 的 `hiddenimports` 中补充后重打包。

## 主要目录

```text
PDFtoWORD/
  pdf_converter.py
  requirements.txt
  PDF转换工具.spec
  converters/
  core/
  ui/
```

## License

MIT
