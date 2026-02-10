"""
PDF 加水印工具

支持文字水印和图片水印，可设置透明度、旋转角度、位置等。
通过 on_progress 回调报告进度，不直接操作UI。
"""

import io
import logging
import math
import os
from datetime import datetime

try:
    from PIL import Image as PILImage
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False


class PDFWatermarkConverter:
    """PDF加水印转换器，与 UI 完全解耦。

    用法::

        converter = PDFWatermarkConverter(on_progress=my_callback)
        result = converter.convert(
            input_file, watermark_text="机密文件",
            opacity=0.3, rotation=45
        )
    """

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, input_file, output_path=None,
                watermark_text=None, watermark_image=None,
                opacity=0.3, rotation=45, font_size=40,
                color=(0.6, 0.6, 0.6), position='tile',
                start_page=None, end_page=None):
        """给PDF添加水印。

        Args:
            input_file: 输入PDF路径
            output_path: 输出路径，None则自动生成
            watermark_text: 文字水印内容
            watermark_image: 图片水印路径
            opacity: 透明度 (0.0-1.0)，越小越透明
            rotation: 旋转角度
            font_size: 文字水印字号
            color: 文字颜色 (R, G, B)，0-1范围
            position: 位置模式
                - 'tile': 铺满（默认）
                - 'center': 居中
                - 'top-left', 'top-right', 'bottom-left', 'bottom-right': 四角
            start_page: 起始页（1-based），None=第1页
            end_page: 结束页（1-based），None=最后一页

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

        if not input_file:
            result['message'] = "请先选择PDF文件！"
            return result

        if not watermark_text and not watermark_image:
            result['message'] = "请输入水印文字或选择水印图片！"
            return result

        if not output_path:
            dir_path = os.path.dirname(input_file)
            basename = os.path.splitext(os.path.basename(input_file))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(dir_path, f"{basename}_水印_{timestamp}.pdf")

        try:
            doc = fitz.open(input_file)
            total_pages = len(doc)

            if total_pages == 0:
                doc.close()
                result['message'] = "PDF文件无内容"
                return result

            # 确定页范围
            s_idx = 0
            e_idx = total_pages
            if start_page is not None or end_page is not None:
                s = max(1, min(start_page or 1, total_pages))
                e = max(s, min(end_page or total_pages, total_pages))
                s_idx = s - 1
                e_idx = e

            opacity = max(0.01, min(1.0, float(opacity)))
            processed = 0
            pages_to_process = e_idx - s_idx

            for page_idx in range(s_idx, e_idx):
                page = doc[page_idx]

                if watermark_image:
                    self._add_image_watermark(
                        page, watermark_image, opacity, position
                    )
                elif watermark_text:
                    self._add_text_watermark(
                        page, watermark_text, opacity, rotation,
                        font_size, color, position
                    )

                processed += 1
                progress = int(processed / pages_to_process * 100)
                self._report(
                    percent=progress,
                    progress_text=f"正在添加水印: 第{page_idx + 1}页 ({progress}%)",
                    status_text=f"处理第 {page_idx + 1}/{total_pages} 页"
                )

            doc.save(output_path)
            doc.close()

            self._report(percent=100, progress_text="水印添加完成！")

            result['success'] = True
            result['output_file'] = output_path
            result['page_count'] = pages_to_process
            result['message'] = f"成功为 {pages_to_process} 页添加水印"

        except Exception as e:
            logging.error(f"PDF加水印失败: {e}", exc_info=True)
            result['message'] = f"加水印失败：{str(e)}"

        return result

    def _add_text_watermark(self, page, text, opacity, rotation,
                            font_size, color, position):
        """在页面上添加文字水印"""
        rect = page.rect
        # PyMuPDF 的 insert_text rotate 参数只接受 0/90/180/270
        # 对于任意角度旋转，使用 0 度插入 + morph 变换实现
        valid_rotate = int(round(rotation / 90.0)) * 90 % 360

        if position == 'tile':
            # 铺满模式：在页面上平铺多个水印
            gap_x = max(font_size * len(text) * 0.8, 200)
            gap_y = max(font_size * 3, 150)

            # 扩展范围以覆盖旋转后的区域
            rad = math.radians(rotation)
            cos_a = abs(math.cos(rad))
            sin_a = abs(math.sin(rad))
            expand = max(rect.width, rect.height) * 0.5

            y = -expand
            while y < rect.height + expand:
                x = -expand
                while x < rect.width + expand:
                    rx = x * math.cos(rad) + y * math.sin(rad)
                    ry = -x * math.sin(rad) + y * math.cos(rad)

                    point = fitz.Point(
                        rect.width / 2 + rx - len(text) * font_size * 0.3,
                        rect.height / 2 + ry
                    )

                    shape = page.new_shape()
                    shape.insert_text(
                        point, text,
                        fontsize=font_size,
                        fontname="china-s",
                        color=color,
                        rotate=valid_rotate,
                    )
                    shape.finish(color=color, fill=color, fill_opacity=opacity)
                    shape.commit()

                    x += gap_x
                y += gap_y

        else:
            # 单个水印
            point = self._get_position_point(rect, position, text, font_size)

            shape = page.new_shape()
            shape.insert_text(
                point, text,
                fontsize=font_size,
                fontname="china-s",
                color=color,
                rotate=valid_rotate,
            )
            shape.finish(color=color, fill=color, fill_opacity=opacity)
            shape.commit()

    def _add_image_watermark(self, page, image_path, opacity, position):
        """在页面上添加图片水印（支持透明度）"""
        rect = page.rect
        try:
            # 使用 Pillow 预处理图片透明度
            if PIL_AVAILABLE:
                pil_img = PILImage.open(image_path).convert("RGBA")
                # 应用透明度：将 alpha 通道乘以 opacity
                alpha = pil_img.split()[3]
                alpha = alpha.point(lambda a: int(a * opacity))
                pil_img.putalpha(alpha)
                img_buf = io.BytesIO()
                pil_img.save(img_buf, format="PNG")
                img_data = img_buf.getvalue()
                img_w, img_h = pil_img.size
            else:
                # 降级：无 Pillow 时直接插入（无透明度）
                with open(image_path, "rb") as f:
                    img_data = f.read()
                tmp_img = fitz.open(image_path)
                tmp_page = tmp_img[0] if tmp_img.page_count else None
                if tmp_page:
                    img_w = tmp_page.rect.width
                    img_h = tmp_page.rect.height
                else:
                    img_w, img_h = 200, 200
                tmp_img.close()

            if position == 'tile':
                # 铺满模式使用较小的图片
                scale = min(rect.width / 4 / img_w,
                            rect.height / 4 / img_h, 1.0)
                scaled_w = img_w * scale
                scaled_h = img_h * scale
                gap_x = scaled_w * 1.5
                gap_y = scaled_h * 1.5

                y = 0
                while y < rect.height:
                    x = 0
                    while x < rect.width:
                        target = fitz.Rect(x, y, x + scaled_w, y + scaled_h)
                        page.insert_image(target, stream=img_data)
                        x += gap_x
                    y += gap_y
            else:
                # 单个水印 - 缩放到页面的 1/3
                scale = min(rect.width / 3 / img_w,
                            rect.height / 3 / img_h, 1.0)
                scaled_w = img_w * scale
                scaled_h = img_h * scale

                if position == 'center':
                    x0 = (rect.width - scaled_w) / 2
                    y0 = (rect.height - scaled_h) / 2
                elif position == 'top-left':
                    x0, y0 = 20, 20
                elif position == 'top-right':
                    x0, y0 = rect.width - scaled_w - 20, 20
                elif position == 'bottom-left':
                    x0, y0 = 20, rect.height - scaled_h - 20
                elif position == 'bottom-right':
                    x0, y0 = rect.width - scaled_w - 20, rect.height - scaled_h - 20
                else:
                    x0 = (rect.width - scaled_w) / 2
                    y0 = (rect.height - scaled_h) / 2

                target = fitz.Rect(x0, y0, x0 + scaled_w, y0 + scaled_h)
                page.insert_image(target, stream=img_data)

        except Exception as e:
            logging.error(f"添加图片水印失败: {e}")
            raise

    @staticmethod
    def _get_position_point(rect, position, text, font_size):
        """根据位置模式计算文字起点"""
        text_w = len(text) * font_size * 0.6
        text_h = font_size

        if position == 'center':
            return fitz.Point(
                (rect.width - text_w) / 2,
                (rect.height + text_h) / 2
            )
        elif position == 'top-left':
            return fitz.Point(30, 30 + text_h)
        elif position == 'top-right':
            return fitz.Point(rect.width - text_w - 30, 30 + text_h)
        elif position == 'bottom-left':
            return fitz.Point(30, rect.height - 30)
        elif position == 'bottom-right':
            return fitz.Point(rect.width - text_w - 30, rect.height - 30)
        else:
            return fitz.Point(
                (rect.width - text_w) / 2,
                (rect.height + text_h) / 2
            )
