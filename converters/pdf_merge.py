"""
PDF 合并工具

将多个PDF文件按选择顺序合并为一个PDF文件。
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


class PDFMergeConverter:
    """PDF合并转换器，与 UI 完全解耦。

    用法::

        converter = PDFMergeConverter(on_progress=my_callback)
        result = converter.convert(files)
    """

    def __init__(self, on_progress=None):
        """
        Args:
            on_progress: fn(percent, progress_text, status_text)
        """
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, files, output_path=None):
        """合并多个PDF文件。

        Args:
            files: PDF文件路径列表（按合并顺序）
            output_path: 输出文件路径，None 则自动生成

        Returns:
            dict with keys:
                success (bool), message (str),
                output_file (str), page_count (int),
                file_count (int)
        """
        result = {
            'success': False, 'message': '',
            'output_file': '', 'page_count': 0,
            'file_count': len(files) if files else 0,
        }

        if not FITZ_AVAILABLE:
            result['message'] = "PyMuPDF (fitz) 未安装！\n请运行: pip install PyMuPDF"
            return result

        if not files or len(files) < 2:
            result['message'] = "请至少选择2个PDF文件进行合并！"
            return result

        if not output_path:
            dir_path = os.path.dirname(files[0])
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(dir_path, f"合并结果_{timestamp}.pdf")

        try:
            merged = fitz.open()
            total_pages = 0

            for idx, pdf_path in enumerate(files):
                self._report(
                    percent=int(idx / len(files) * 90),
                    progress_text=f"正在合并第 {idx + 1}/{len(files)} 个文件",
                    status_text=f"正在处理: {os.path.basename(pdf_path)}"
                )

                doc = fitz.open(pdf_path)
                merged.insert_pdf(doc)
                total_pages += len(doc)
                doc.close()

            self._report(
                percent=95,
                progress_text="正在保存合并后的文件...",
                status_text="正在保存..."
            )

            merged.save(output_path)
            merged.close()

            self._report(percent=100, progress_text="合并完成！")

            result['success'] = True
            result['output_file'] = output_path
            result['page_count'] = total_pages
            result['message'] = f"成功合并 {len(files)} 个文件，共 {total_pages} 页"

        except Exception as e:
            logging.error(f"PDF合并失败: {e}", exc_info=True)
            result['message'] = f"合并失败：{str(e)}"

        return result
