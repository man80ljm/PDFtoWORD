"""
PDF 拆分工具

支持三种拆分模式：每页一个PDF、按固定间隔拆分、按自定义范围拆分。
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


class PDFSplitConverter:
    """PDF拆分转换器，与 UI 完全解耦。

    用法::

        converter = PDFSplitConverter(on_progress=my_callback)
        result = converter.convert(input_file, mode='every_page')
    """

    def __init__(self, on_progress=None):
        """
        Args:
            on_progress: fn(percent, progress_text, status_text)
        """
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, input_file, mode='every_page', interval=1,
                ranges=None, output_dir=None):
        """拆分PDF文件。

        Args:
            input_file: 输入PDF路径
            mode: 拆分模式
                - 'every_page': 每页一个PDF
                - 'by_interval': 每N页一个PDF
                - 'by_ranges': 按自定义范围拆分（如 "1-3,4-6,7-10"）
            interval: 每N页拆分（mode='by_interval'时使用）
            ranges: 自定义范围字符串（mode='by_ranges'时使用）
            output_dir: 输出目录，None则自动生成

        Returns:
            dict with keys:
                success (bool), message (str),
                output_dir (str), output_files (list[str]),
                page_count (int), file_count (int)
        """
        result = {
            'success': False, 'message': '',
            'output_dir': '', 'output_files': [],
            'page_count': 0, 'file_count': 0,
        }

        if not FITZ_AVAILABLE:
            result['message'] = "PyMuPDF (fitz) 未安装！\n请运行: pip install PyMuPDF"
            return result

        if not input_file:
            result['message'] = "请先选择PDF文件！"
            return result

        try:
            doc = fitz.open(input_file)
            total_pages = len(doc)
        except Exception as e:
            result['message'] = f"无法打开PDF文件：{e}"
            return result

        if total_pages == 0:
            doc.close()
            result['message'] = "PDF文件无内容"
            return result

        # 准备输出目录
        basename = os.path.splitext(os.path.basename(input_file))[0]
        if not output_dir:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = os.path.join(
                os.path.dirname(input_file),
                f"{basename}_拆分_{timestamp}"
            )
        os.makedirs(output_dir, exist_ok=True)
        result['output_dir'] = output_dir
        result['page_count'] = total_pages

        # 确定分组
        try:
            if mode == 'every_page':
                groups = [[i] for i in range(total_pages)]
            elif mode == 'by_interval':
                interval = max(1, int(interval))
                groups = []
                for start in range(0, total_pages, interval):
                    end = min(start + interval, total_pages)
                    groups.append(list(range(start, end)))
            elif mode == 'by_ranges':
                groups = self._parse_ranges(ranges, total_pages)
            else:
                groups = [[i] for i in range(total_pages)]
        except ValueError as e:
            doc.close()
            result['message'] = str(e)
            return result

        # 执行拆分
        output_files = []
        for group_idx, pages in enumerate(groups):
            self._report(
                percent=int((group_idx + 1) / len(groups) * 100),
                progress_text=f"正在拆分第 {group_idx + 1}/{len(groups)} 个部分",
                status_text="拆分中..."
            )

            try:
                new_doc = fitz.open()
                for page_num in pages:
                    new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

                # 以页范围命名
                if len(pages) == 1:
                    out_name = f"{basename}_第{pages[0] + 1}页.pdf"
                else:
                    out_name = f"{basename}_第{pages[0] + 1}-{pages[-1] + 1}页.pdf"

                out_path = os.path.join(output_dir, out_name)
                new_doc.save(out_path)
                new_doc.close()
                output_files.append(out_path)
            except Exception as e:
                logging.error(f"拆分第{group_idx + 1}部分失败: {e}")
                result['message'] = f"拆分第{group_idx + 1}部分时出错：{e}"
                doc.close()
                return result

        doc.close()

        result['success'] = True
        result['output_files'] = output_files
        result['file_count'] = len(output_files)
        result['message'] = f"成功拆分为 {len(output_files)} 个文件"

        self._report(percent=100, progress_text="拆分完成！")
        return result

    @staticmethod
    def _parse_ranges(ranges_str, total_pages):
        """解析范围字符串，如 "1-3,4-6,7-10"

        Returns:
            list[list[int]]: 0-based 页码分组
        """
        if not ranges_str or not ranges_str.strip():
            raise ValueError("请输入拆分范围，如：1-3,4-6,7-10")

        groups = []
        # 兼容中英文逗号 / 分号
        text = ranges_str.replace('，', ',').replace('；', ';').replace(';', ',')
        parts = text.split(',')

        for part in parts:
            part = part.strip()
            if not part:
                continue

            if '-' in part:
                bounds = part.split('-', 1)
                try:
                    start = int(bounds[0].strip())
                    end = int(bounds[1].strip())
                except ValueError:
                    raise ValueError(f"无效的范围格式: {part}")
                if start < 1 or end < 1:
                    raise ValueError(f"页码必须从1开始: {part}")
                if start > total_pages or end > total_pages:
                    raise ValueError(f"页码超出范围(共{total_pages}页): {part}")
                if start > end:
                    raise ValueError(f"起始页不能大于结束页: {part}")
                groups.append(list(range(start - 1, end)))
            else:
                try:
                    page = int(part)
                except ValueError:
                    raise ValueError(f"无效的页码: {part}")
                if page < 1 or page > total_pages:
                    raise ValueError(f"页码超出范围(共{total_pages}页): {part}")
                groups.append([page - 1])

        if not groups:
            raise ValueError("未指定有效的拆分范围")
        return groups
