"""
图片 → PDF 转换器

将多张图片合并为一个PDF文件，支持多种页面尺寸。
通过 on_progress 回调报告进度，不直接操作UI。
"""

import logging
import os
from datetime import datetime

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

# 标准页面尺寸（单位：点，72 pts/inch）
PAGE_SIZES = {
    'A4': (595.28, 841.89),
    'A3': (841.89, 1190.55),
    'Letter': (612, 792),
    'Legal': (612, 1008),
    '自适应': None,       # 页面大小匹配图片
}

SUPPORTED_IMAGE_EXTS = {
    '.png', '.jpg', '.jpeg', '.bmp', '.gif',
    '.tiff', '.tif', '.webp',
}


class ImageToPDFConverter:
    """图片→PDF 转换器，与 UI 完全解耦。

    用法::

        converter = ImageToPDFConverter(on_progress=my_callback)
        result = converter.convert(files, page_size='A4')
    """

    def __init__(self, on_progress=None):
        """
        Args:
            on_progress: fn(percent, progress_text, status_text)
        """
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, files, output_path=None, page_size='A4', margin=20):
        """将图片转换为PDF。

        Args:
            files: 图片文件路径列表
            output_path: 输出PDF路径，None 则自动生成
            page_size: 页面尺寸 ('A4', 'A3', 'Letter', 'Legal', '自适应')
            margin: 页边距(点)，自适应模式下为 0

        Returns:
            dict with keys:
                success (bool), message (str),
                output_file (str), page_count (int)
        """
        result = {
            'success': False, 'message': '',
            'output_file': '', 'page_count': 0,
        }

        if not FITZ_AVAILABLE:
            result['message'] = "PyMuPDF (fitz) 未安装！\n请运行: pip install PyMuPDF"
            return result

        if not files:
            result['message'] = "请先选择图片文件！"
            return result

        # 过滤有效图片
        valid_files = []
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext in SUPPORTED_IMAGE_EXTS:
                valid_files.append(f)
            else:
                logging.warning(f"跳过不支持的文件格式: {f}")

        if not valid_files:
            result['message'] = (
                "没有有效的图片文件！\n"
                "支持的格式：PNG, JPG, BMP, GIF, TIFF, WebP"
            )
            return result

        if not output_path:
            dir_path = os.path.dirname(valid_files[0])
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(dir_path, f"图片转PDF_{timestamp}.pdf")

        is_auto_size = (page_size == '自适应')
        target_size = PAGE_SIZES.get(page_size) if not is_auto_size else None

        try:
            doc = fitz.open()

            for idx, img_path in enumerate(valid_files):
                self._report(
                    percent=int((idx + 1) / len(valid_files) * 100),
                    progress_text=f"正在处理第 {idx + 1}/{len(valid_files)} 张图片",
                    status_text=f"正在转换: {os.path.basename(img_path)}"
                )

                try:
                    img = fitz.open(img_path)
                    # 将图片渲染为单页PDF
                    img_pdf = fitz.open("pdf", img.convert_to_pdf())
                    img_page = img_pdf[0]
                    img_rect = img_page.rect
                    img.close()

                    if is_auto_size:
                        # 页面尺寸匹配图片
                        page = doc.new_page(
                            width=img_rect.width, height=img_rect.height
                        )
                        page.show_pdf_page(page.rect, img_pdf, 0)
                    else:
                        # 将图片适配到指定页面尺寸（带页边距）
                        pw, ph = target_size
                        page = doc.new_page(width=pw, height=ph)

                        avail_w = pw - 2 * margin
                        avail_h = ph - 2 * margin

                        # 等比缩放
                        scale_w = avail_w / img_rect.width if img_rect.width > 0 else 1
                        scale_h = avail_h / img_rect.height if img_rect.height > 0 else 1
                        scale = min(scale_w, scale_h)

                        new_w = img_rect.width * scale
                        new_h = img_rect.height * scale

                        # 居中放置
                        x0 = (pw - new_w) / 2
                        y0 = (ph - new_h) / 2
                        target_rect = fitz.Rect(x0, y0, x0 + new_w, y0 + new_h)

                        page.show_pdf_page(target_rect, img_pdf, 0)

                    img_pdf.close()

                except Exception as e:
                    logging.error(f"处理图片失败: {img_path}, {e}")
                    result['message'] = (
                        f"处理图片失败: {os.path.basename(img_path)}\n{e}"
                    )
                    doc.close()
                    return result

            doc.save(output_path)
            doc.close()

            self._report(percent=100, progress_text="转换完成！")

            result['success'] = True
            result['output_file'] = output_path
            result['page_count'] = len(valid_files)
            result['message'] = f"成功将 {len(valid_files)} 张图片转换为PDF"

        except Exception as e:
            logging.error(f"图片转PDF失败: {e}", exc_info=True)
            result['message'] = f"转换失败：{str(e)}"

        return result
