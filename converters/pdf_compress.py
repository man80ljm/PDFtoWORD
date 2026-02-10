"""
PDF 压缩工具

通过降低图片质量/DPI、清理冗余对象来减小PDF文件大小。
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


# 压缩级别预设：(图片质量, 目标DPI, garbage回收级别, deflate压缩)
COMPRESS_PRESETS = {
    '轻度压缩': {
        'image_quality': 85,
        'max_dpi': 200,
        'garbage': 2,
        'description': '文件略微减小，画质几乎无损',
    },
    '标准压缩': {
        'image_quality': 60,
        'max_dpi': 150,
        'garbage': 3,
        'description': '平衡文件大小与画质（推荐）',
    },
    '极限压缩': {
        'image_quality': 30,
        'max_dpi': 96,
        'garbage': 4,
        'description': '最大程度压缩，画质会明显下降',
    },
}


class PDFCompressConverter:
    """PDF压缩转换器，与 UI 完全解耦。

    用法::

        converter = PDFCompressConverter(on_progress=my_callback)
        result = converter.convert("input.pdf", compress_level="标准压缩")
    """

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, input_file, output_path=None, compress_level='标准压缩'):
        """压缩PDF文件。

        Args:
            input_file: 输入PDF路径
            output_path: 输出路径，None则自动生成
            compress_level: 压缩级别 ('轻度压缩' / '标准压缩' / '极限压缩')

        Returns:
            dict with keys:
                success (bool), message (str),
                output_file (str), page_count (int),
                original_size (int), compressed_size (int),
                ratio (float)  -- 压缩率百分比
        """
        result = {
            'success': False, 'message': '',
            'output_file': '', 'page_count': 0,
            'original_size': 0, 'compressed_size': 0,
            'ratio': 0.0,
        }

        if not FITZ_AVAILABLE:
            result['message'] = "PyMuPDF (fitz) 未安装！\n请运行: pip install PyMuPDF"
            return result

        if not input_file or not os.path.exists(input_file):
            result['message'] = "请先选择PDF文件！"
            return result

        preset = COMPRESS_PRESETS.get(compress_level, COMPRESS_PRESETS['标准压缩'])
        image_quality = preset['image_quality']
        max_dpi = preset['max_dpi']
        garbage_level = preset['garbage']

        if not output_path:
            dir_path = os.path.dirname(input_file)
            basename = os.path.splitext(os.path.basename(input_file))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(dir_path, f"{basename}_压缩_{timestamp}.pdf")

        original_size = os.path.getsize(input_file)
        result['original_size'] = original_size

        try:
            doc = fitz.open(input_file)
            total_pages = len(doc)

            if total_pages == 0:
                doc.close()
                result['message'] = "PDF文件无内容"
                return result

            result['page_count'] = total_pages
            self._report(percent=0, progress_text="正在分析PDF...",
                         status_text=f"共 {total_pages} 页，开始压缩")

            # 阶段1：压缩每页中的图片（去重，同一图片可能在多页引用）
            images_processed = 0
            processed_xrefs = set()
            for page_idx in range(total_pages):
                page = doc[page_idx]
                image_list = page.get_images(full=True)

                for img_idx, img_info in enumerate(image_list):
                    xref = img_info[0]
                    if xref in processed_xrefs:
                        continue
                    processed_xrefs.add(xref)
                    try:
                        if self._compress_image(doc, xref, image_quality, max_dpi):
                            images_processed += 1
                    except Exception as e:
                        logging.debug(f"跳过图片 xref={xref}: {e}")

                # 进度：图片压缩占 80%
                progress = int((page_idx + 1) / total_pages * 80)
                self._report(
                    percent=progress,
                    progress_text=f"压缩图片: 第{page_idx + 1}/{total_pages}页 ({progress}%)",
                    status_text=f"已处理 {images_processed} 张图片"
                )

            # 阶段2：保存 - 使用 garbage 回收和 deflate 压缩
            self._report(percent=85, progress_text="正在优化文件结构...",
                         status_text="清理冗余对象、压缩数据流")

            doc.save(
                output_path,
                garbage=garbage_level,       # 回收未使用的对象
                deflate=True,                # 压缩数据流
                deflate_images=True,         # 压缩图片流
                deflate_fonts=True,          # 压缩字体流
                clean=True,                  # 清理内容流
                linear=True,                 # 线性化（加快网络加载）
            )
            doc.close()

            # 计算压缩结果
            compressed_size = os.path.getsize(output_path)
            result['compressed_size'] = compressed_size

            if original_size > 0:
                ratio = (1 - compressed_size / original_size) * 100
                result['ratio'] = round(ratio, 1)
            else:
                ratio = 0

            # 如果压缩后反而更大，提示用户
            if compressed_size >= original_size:
                result['success'] = True
                result['output_file'] = output_path
                result['message'] = (
                    f"文件已处理，但大小未减小（原始 {self._format_size(original_size)} → "
                    f"{self._format_size(compressed_size)}）。\n"
                    f"该PDF可能已经是最优状态或不含可压缩的图片。"
                )
            else:
                result['success'] = True
                result['output_file'] = output_path
                result['message'] = (
                    f"压缩完成！\n"
                    f"原始大小：{self._format_size(original_size)}\n"
                    f"压缩后：{self._format_size(compressed_size)}\n"
                    f"减小了 {abs(ratio):.1f}%（节省 {self._format_size(original_size - compressed_size)}）"
                )

            self._report(percent=100, progress_text="压缩完成！")
            logging.info(
                f"PDF压缩: {self._format_size(original_size)} → "
                f"{self._format_size(compressed_size)} ({ratio:.1f}%)"
            )

        except Exception as e:
            logging.error(f"PDF压缩失败: {e}", exc_info=True)
            result['message'] = f"压缩失败：{str(e)}"

        return result

    def _compress_image(self, doc, xref, quality, max_dpi):
        """压缩文档中的单个图片。

        将图片提取 → Pillow 重编码为低质量 JPEG → 通过 update_stream 写回。
        """
        # 检查是否为图片 xref
        if not doc.xref_is_image(xref):
            return False

        img_info = doc.extract_image(xref)
        if not img_info:
            return False

        img_data = img_info['image']
        width = img_info.get('width', 0)
        height = img_info.get('height', 0)

        # 跳过太小的图片（图标等）或已经很小的数据
        if width <= 50 or height <= 50 or len(img_data) < 5000:
            return False

        try:
            from PIL import Image as PILImage
            import io

            pil_img = PILImage.open(io.BytesIO(img_data))

            # 透明通道转白底 RGB（JPEG 不支持透明）
            has_alpha = pil_img.mode in ('RGBA', 'P', 'LA')
            if has_alpha:
                background = PILImage.new('RGB', pil_img.size, (255, 255, 255))
                if pil_img.mode == 'P':
                    pil_img = pil_img.convert('RGBA')
                if pil_img.mode in ('RGBA', 'LA'):
                    background.paste(pil_img, mask=pil_img.split()[-1])
                else:
                    background.paste(pil_img)
                pil_img = background
            elif pil_img.mode != 'RGB':
                pil_img = pil_img.convert('RGB')

            # 根据 max_dpi 缩小尺寸
            new_w, new_h = pil_img.size
            if max_dpi < 200 and (width > 1000 or height > 1000):
                scale = max_dpi / 200.0
                new_w = max(int(width * scale), 100)
                new_h = max(int(height * scale), 100)
                pil_img = pil_img.resize((new_w, new_h), PILImage.LANCZOS)
                new_w, new_h = pil_img.size

            # 重新编码为 JPEG
            buf = io.BytesIO()
            pil_img.save(buf, format='JPEG', quality=quality, optimize=True)
            new_data = buf.getvalue()

            # 仅在确实变小时替换
            if len(new_data) < len(img_data) * 0.95:
                # 直接替换 xref 的流数据并更新字典属性
                doc.update_stream(xref, new_data, compress=False)
                doc.xref_set_key(xref, "Width", str(new_w))
                doc.xref_set_key(xref, "Height", str(new_h))
                doc.xref_set_key(xref, "ColorSpace", "/DeviceRGB")
                doc.xref_set_key(xref, "BitsPerComponent", "8")
                doc.xref_set_key(xref, "Filter", "/DCTDecode")
                doc.xref_set_key(xref, "DecodeParms", "null")
                # JPEG 无透明，清除 SMask 引用
                if has_alpha:
                    doc.xref_set_key(xref, "SMask", "null")
                return True

        except ImportError:
            pass
        except Exception as e:
            logging.debug(f"图片压缩跳过 xref={xref}: {e}")

        return False

    @staticmethod
    def _format_size(size_bytes):
        """格式化文件大小"""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
