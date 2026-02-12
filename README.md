# PDF转换工具

一个基于 `tkinter` 的桌面 PDF 工具箱，支持常见 PDF 转换、批处理、盖章/签名预览、页面重排、书签处理等功能。

## 已实现功能

- `PDF转Word`
- `PDF转图片`
- `PDF合并`
- `PDF拆分`
- `图片转PDF`
- `PDF加水印`
- `PDF加密/解密`
- `PDF压缩`
- `PDF提取/删页`
- `OCR可搜索PDF`
- `PDF转Excel`
- `PDF批量文本/图片提取`
- `PDF批量盖章`（普通章/骑缝章/二维码/模板/签名）
- `PDF页面重排/旋转/倒序`
- `PDF添加/移除书签`

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

