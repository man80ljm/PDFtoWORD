"""
PDF → 图片 批量转换器

支持多文件批量转换，每个PDF输出到以文件名命名的文件夹。
通过 on_progress 回调报告进度，不直接操作UI。
"""

import logging
import os
import time
from datetime import datetime

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False


class PDFToImageConverter:
    """PDF→图片 批量转换器，与 UI 完全解耦。

    用法::

        converter = PDFToImageConverter(on_progress=my_callback)
        result = converter.convert(files, dpi=200, img_format='PNG')
    """

    def __init__(self, on_progress=None):
        """
        Args:
            on_progress: fn(percent, progress_text, status_text)
        """
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, files, dpi=200, img_format='PNG',
                start_page=None, end_page=None):
        """批量转换PDF为图片。

        Args:
            files: PDF文件路径列表
            dpi: 输出DPI (36-1200)
            img_format: 'PNG' 或 'JPEG'
            start_page: 起始页（1-based），None=第1页
            end_page: 结束页（1-based），None=最后一页

        Returns:
            dict with keys:
                success (bool), message (str),
                output_dirs (list[str]), page_count (int),
                file_count (int), errors (list[str]),
                dpi (int), format (str)
        """
        result = {
            'success': False, 'message': '',
            'output_dirs': [], 'page_count': 0,
            'file_count': len(files), 'errors': [],
            'dpi': dpi, 'format': img_format,
        }

        if not FITZ_AVAILABLE:
            result['message'] = "PyMuPDF (fitz) 未安装！\n请运行: pip install PyMuPDF"
            return result

        if not files:
            result['message'] = "请先选择PDF文件！"
            return result

        # 校验DPI
        try:
            dpi = int(dpi)
            if dpi < 36 or dpi > 1200:
                raise ValueError
        except (ValueError, TypeError):
            result['message'] = "DPI必须是36-1200之间的整数"
            return result

        if img_format not in ("PNG", "JPEG"):
            img_format = "PNG"
        ext = ".png" if img_format == "PNG" else ".jpg"
        zoom = dpi / 72.0

        # 计算实际处理页数（考虑页范围）
        use_range = start_page is not None or end_page is not None
        total_pages_all = 0
        file_page_counts = []
        for f in files:
            try:
                doc = fitz.open(f)
                count = len(doc)
                doc.close()
                file_page_counts.append(count)
                if use_range:
                    s = max(1, min(start_page or 1, count))
                    e = max(s, min(end_page or count, count))
                    total_pages_all += (e - s + 1)
                else:
                    total_pages_all += count
            except Exception as e:
                result['message'] = f"无法打开: {os.path.basename(f)}\n{e}"
                return result

        if total_pages_all == 0:
            result['message'] = "所有PDF文件均无内容"
            return result
        processed = 0
        output_dirs = []
        errors = []

        for file_idx, pdf_path in enumerate(files):
            basename = os.path.splitext(os.path.basename(pdf_path))[0]
            output_dir = os.path.join(os.path.dirname(pdf_path), basename)

            if os.path.exists(output_dir):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_dir = os.path.join(os.path.dirname(pdf_path),
                                          f"{basename}_{timestamp}")
            os.makedirs(output_dir, exist_ok=True)
            output_dirs.append(output_dir)

            try:
                doc = fitz.open(pdf_path)
                page_count = len(doc)

                # 确定页范围
                s_idx = 0
                e_idx = page_count
                if use_range:
                    s = start_page if start_page else 1
                    e = end_page if end_page else page_count
                    s = max(1, min(s, page_count))
                    e = max(s, min(e, page_count))
                    s_idx = s - 1
                    e_idx = e

                file_label = os.path.basename(pdf_path)
                for page_idx in range(s_idx, e_idx):
                    page = doc[page_idx]
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat, alpha=False)

                    img_filename = f"{page_idx + 1}{ext}"
                    img_path = os.path.join(output_dir, img_filename)

                    if img_format == "JPEG":
                        pix.save(img_path, jpg_quality=95)
                    else:
                        pix.save(img_path)

                    processed += 1
                    progress = int(processed / total_pages_all * 100)
                    page_num = page_idx + 1
                    self._report(
                        percent=progress,
                        progress_text=f"[{file_idx+1}/{len(files)}] {file_label} - "
                                      f"第{page_num}页 ({progress}%)",
                        status_text=f"正在转换: {file_label}",
                    )

                doc.close()
            except Exception as e:
                errors.append(f"{os.path.basename(pdf_path)}: {str(e)}")
                logging.error(f"PDF转图片失败 [{pdf_path}]: {e}")

        result['success'] = len(errors) == 0 or processed > 0
        result['output_dirs'] = output_dirs
        result['page_count'] = processed
        result['errors'] = errors
        result['dpi'] = dpi
        result['format'] = img_format
        self._report(percent=100, progress_text="转换完成！(100%)")

        return result
