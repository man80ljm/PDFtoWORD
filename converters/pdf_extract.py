"""
PDF 页面提取/删除工具

支持提取指定页面或删除指定页面，生成新的PDF。
页码格式：1,3,5-10（逗号分隔，支持范围）
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


class PDFExtractConverter:
    """PDF页面提取/删除转换器，与 UI 完全解耦。

    用法::

        converter = PDFExtractConverter(on_progress=my_callback)
        result = converter.convert("input.pdf", pages_str="1,3,5-10", mode="提取")
    """

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, input_file, output_path=None, pages_str='', mode='提取'):
        """提取或删除PDF指定页面。

        Args:
            input_file: 输入PDF路径
            output_path: 输出路径，None则自动生成
            pages_str: 页码字符串，如 "1,3,5-10"（1-based）
            mode: '提取' = 仅保留指定页；'删除' = 移除指定页

        Returns:
            dict with keys:
                success (bool), message (str),
                output_file (str), page_count (int),
                original_pages (int), result_pages (int)
        """
        result = {
            'success': False, 'message': '',
            'output_file': '', 'page_count': 0,
            'original_pages': 0, 'result_pages': 0,
        }

        if not FITZ_AVAILABLE:
            result['message'] = "PyMuPDF (fitz) 未安装！\n请运行: pip install PyMuPDF"
            return result

        if not input_file or not os.path.exists(input_file):
            result['message'] = "请先选择PDF文件！"
            return result

        if not pages_str.strip():
            result['message'] = "请输入页码！\n格式：1,3,5-10"
            return result

        try:
            doc = fitz.open(input_file)
            total_pages = len(doc)

            if total_pages == 0:
                doc.close()
                result['message'] = "PDF文件无内容"
                return result

            result['original_pages'] = total_pages

            # 解析页码字符串
            specified_pages, error = self._parse_pages(pages_str, total_pages)
            if error:
                doc.close()
                result['message'] = error
                return result

            if not specified_pages:
                doc.close()
                result['message'] = "未指定有效页码"
                return result

            # 根据模式计算最终保留的页面
            all_pages = set(range(total_pages))
            if mode == '提取':
                keep_pages = sorted(specified_pages)
                action_text = "提取"
            else:  # 删除
                keep_pages = sorted(all_pages - specified_pages)
                action_text = "删除"

            if not keep_pages:
                doc.close()
                result['message'] = f"{action_text}后没有剩余页面！"
                return result

            # 生成输出路径
            if not output_path:
                dir_path = os.path.dirname(input_file)
                basename = os.path.splitext(os.path.basename(input_file))[0]
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_path = os.path.join(
                    dir_path, f"{basename}_{action_text}_{timestamp}.pdf")

            self._report(percent=10,
                         progress_text=f"正在{action_text}页面...",
                         status_text=f"原始 {total_pages} 页，{action_text} {len(specified_pages)} 页")

            # 使用 select 保留指定页面
            doc.select(keep_pages)

            self._report(percent=70,
                         progress_text="正在保存...",
                         status_text=f"输出 {len(keep_pages)} 页")

            doc.save(output_path, garbage=3, deflate=True)
            doc.close()

            result_pages = len(keep_pages)
            result['success'] = True
            result['output_file'] = output_path
            result['page_count'] = result_pages
            result['result_pages'] = result_pages

            if mode == '提取':
                pages_display = self._format_pages(specified_pages)
                result['message'] = (
                    f"成功提取 {result_pages} 页\n"
                    f"提取的页码：{pages_display}\n"
                    f"原始 {total_pages} 页 → 输出 {result_pages} 页"
                )
            else:
                pages_display = self._format_pages(specified_pages)
                result['message'] = (
                    f"成功删除 {len(specified_pages)} 页\n"
                    f"删除的页码：{pages_display}\n"
                    f"原始 {total_pages} 页 → 剩余 {result_pages} 页"
                )

            self._report(percent=100, progress_text=f"页面{action_text}完成！")

        except Exception as e:
            logging.error(f"PDF页面{mode}失败: {e}", exc_info=True)
            result['message'] = f"页面{mode}失败：{str(e)}"

        return result

    @staticmethod
    def _parse_pages(pages_str, total_pages):
        """解析页码字符串，返回 (set_of_0based_indices, error_msg)。

        支持格式：
            "1,3,5"       → 单独页
            "2-8"         → 范围
            "1,3-5,8,10-12" → 混合
        """
        pages = set()
        # 统一替换中文逗号/分号
        pages_str = pages_str.replace('，', ',').replace('；', ',').replace(';', ',')

        for part in pages_str.split(','):
            part = part.strip()
            if not part:
                continue

            if '-' in part:
                # 范围：如 "3-8"
                segments = part.split('-', 1)
                try:
                    start = int(segments[0].strip())
                    end = int(segments[1].strip())
                except ValueError:
                    return None, f"无效的页码范围：{part}\n格式示例：1,3,5-10"

                if start < 1 or end < 1:
                    return None, f"页码必须大于0：{part}"
                if start > total_pages or end > total_pages:
                    return None, f"页码超出范围（共{total_pages}页）：{part}"
                if start > end:
                    return None, f"起始页不能大于结束页：{part}"

                for p in range(start, end + 1):
                    pages.add(p - 1)  # 转为 0-based
            else:
                # 单个页码
                try:
                    p = int(part)
                except ValueError:
                    return None, f"无效的页码：{part}\n格式示例：1,3,5-10"

                if p < 1:
                    return None, f"页码必须大于0：{part}"
                if p > total_pages:
                    return None, f"页码 {p} 超出范围（共{total_pages}页）"

                pages.add(p - 1)  # 转为 0-based

        return pages, None

    @staticmethod
    def _format_pages(pages_0based):
        """将0-based页码集合格式化为友好显示（1-based，合并连续范围）"""
        if not pages_0based:
            return ""
        sorted_pages = sorted(p + 1 for p in pages_0based)  # 转回1-based
        ranges = []
        start = sorted_pages[0]
        end = start

        for p in sorted_pages[1:]:
            if p == end + 1:
                end = p
            else:
                ranges.append(f"{start}" if start == end else f"{start}-{end}")
                start = end = p
        ranges.append(f"{start}" if start == end else f"{start}-{end}")

        return ", ".join(ranges)
