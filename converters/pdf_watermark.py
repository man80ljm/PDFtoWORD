"""
PDF 加水印工具

支持文字水印和图片水印，可设置透明度、旋转角度、位置等。
通过 on_progress 回调报告进度，不直接操作UI。
"""

import io
import logging
import os
import random
import re
from datetime import datetime

try:
    from PIL import Image as PILImage, ImageDraw as PILImageDraw, ImageFont as PILImageFont
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
                start_page=None, end_page=None, pages_str="",
                size_scale=1.0, layout='grid',
                spacing_scale=1.0,
                random_size=False, random_strength=0.35):
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

            # 确定页范围：优先按 pages_str（支持 1,2,5-8），留空走旧的 start/end
            page_indices = []
            if pages_str and str(pages_str).strip():
                parsed = self._parse_pages_str(pages_str, total_pages)
                if parsed is None:
                    doc.close()
                    result['message'] = f"页码格式不正确：{pages_str}"
                    return result
                page_indices = parsed
            else:
                s_idx = 0
                e_idx = total_pages
                if start_page is not None or end_page is not None:
                    s = max(1, min(start_page or 1, total_pages))
                    e = max(s, min(end_page or total_pages, total_pages))
                    s_idx = s - 1
                    e_idx = e
                page_indices = list(range(s_idx, e_idx))

            opacity = max(0.01, min(1.0, float(opacity)))
            size_scale = max(0.2, min(3.0, float(size_scale)))
            spacing_scale = max(0.5, min(2.0, float(spacing_scale)))
            random_strength = max(0.0, min(1.0, float(random_strength)))
            processed = 0
            pages_to_process = len(page_indices)
            if pages_to_process <= 0:
                doc.close()
                result['message'] = "没有匹配到可处理页码"
                return result

            for page_idx in page_indices:
                page = doc[page_idx]

                if watermark_image:
                    self._add_image_watermark(
                        page, watermark_image, opacity, position,
                        rotation=rotation, size_scale=size_scale,
                        spacing_scale=spacing_scale,
                        layout=layout, random_size=random_size,
                        random_strength=random_strength, page_seed=page_idx + 1
                    )
                elif watermark_text:
                    self._add_text_watermark(
                        page, watermark_text, opacity, rotation,
                        font_size, color, position,
                        size_scale=size_scale, layout=layout, spacing_scale=spacing_scale,
                        random_size=random_size,
                        random_strength=random_strength, page_seed=page_idx + 1
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
                            font_size, color, position,
                            size_scale=1.0, layout='grid', spacing_scale=1.0,
                            random_size=False, random_strength=0.35,
                            page_seed=1):
        """在页面上添加文字水印"""
        rect = page.rect
        # 文字尺寸以 font_size 为准，避免与 size_scale 叠乘导致预览和导出不一致
        base_font = max(8, int(font_size))
        is_tile = self._is_tile_mode(position)
        tile_layout = self._resolve_tile_layout(position, layout)
        color255 = self._normalize_color255(color)

        if not PIL_AVAILABLE:
            # 降级路径（无 Pillow 时）
            valid_rotate = int(round(rotation / 90.0)) * 90 % 360
            point = self._get_position_point(rect, position, text, base_font)
            shape = page.new_shape()
            shape.insert_text(
                point, text,
                fontsize=base_font,
                fontname="china-s",
                color=(
                    color255[0] / 255.0,
                    color255[1] / 255.0,
                    color255[2] / 255.0,
                ),
                rotate=valid_rotate,
            )
            shape.finish(
                color=(
                    color255[0] / 255.0,
                    color255[1] / 255.0,
                    color255[2] / 255.0,
                ),
                fill=(
                    color255[0] / 255.0,
                    color255[1] / 255.0,
                    color255[2] / 255.0,
                ),
                fill_opacity=opacity,
            )
            shape.commit()
            return

        stamp_cache = {}

        if is_tile:
            base_w = max(20, int(base_font * max(1, len(text)) * 0.6))
            base_h = max(16, int(base_font * 1.5))
            for cx, cy, row, col in self._iter_positions(
                page_w=rect.width,
                page_h=rect.height,
                base_w=base_w,
                base_h=base_h,
                spacing_scale=spacing_scale,
                tile_layout=tile_layout,
            ):
                scale_factor = self._tile_size_factor(
                    page_seed=page_seed,
                    row=row,
                    col=col,
                    enabled=random_size,
                    strength=random_strength,
                )
                draw_font = max(8, int(base_font * scale_factor))
                key = (draw_font, int(opacity * 1000), int(round(rotation)))
                cached = stamp_cache.get(key)
                if cached is None:
                    stamp = self._render_text_stamp(
                        text=text,
                        font_px=draw_font,
                        color255=color255,
                        opacity=opacity,
                        rotation=rotation,
                    )
                    stamp_bytes = self._pil_to_png_bytes(stamp)
                    cached = (stamp_bytes, stamp.width, stamp.height)
                    stamp_cache[key] = cached
                stamp_bytes, sw, sh = cached
                x = cx - sw / 2
                y = cy - sh / 2
                page.insert_image(
                    fitz.Rect(x, y, x + sw, y + sh),
                    stream=stamp_bytes,
                    keep_proportion=True,
                    overlay=True,
                )
        else:
            stamp = self._render_text_stamp(
                text=text,
                font_px=base_font,
                color255=color255,
                opacity=opacity,
                rotation=rotation,
            )
            sw, sh = stamp.size
            x0, y0 = self._single_anchor_xy(rect, position, sw, sh)
            page.insert_image(
                fitz.Rect(x0, y0, x0 + sw, y0 + sh),
                stream=self._pil_to_png_bytes(stamp),
                keep_proportion=True,
                overlay=True,
            )

    def _add_image_watermark(self, page, image_path, opacity, position,
                             rotation=45, size_scale=1.0, layout='grid',
                             spacing_scale=1.0,
                             random_size=False, random_strength=0.35,
                             page_seed=1):
        """在页面上添加图片水印（支持透明度）"""
        rect = page.rect
        try:
            if PIL_AVAILABLE:
                pil_img = PILImage.open(image_path).convert("RGBA")
                alpha = pil_img.split()[3]
                alpha = alpha.point(lambda a: int(a * opacity))
                pil_img.putalpha(alpha)
                if abs(float(rotation)) > 0.01:
                    pil_img = pil_img.rotate(float(rotation), expand=True, resample=PILImage.BICUBIC)
                img_w, img_h = pil_img.size
            else:
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
                if self._is_tile_mode(position):
                    scale = min(rect.width / 4 / img_w, rect.height / 4 / img_h, 1.0) * size_scale
                else:
                    scale = min(rect.width / 3 / img_w, rect.height / 3 / img_h, 1.0) * size_scale
                scaled_w = max(12, img_w * scale)
                scaled_h = max(12, img_h * scale)
                x0 = (rect.width - scaled_w) / 2
                y0 = (rect.height - scaled_h) / 2
                target = fitz.Rect(x0, y0, x0 + scaled_w, y0 + scaled_h)
                page.insert_image(target, stream=img_data)
                return

            is_tile = self._is_tile_mode(position)
            tile_layout = self._resolve_tile_layout(position, layout)

            if is_tile:
                # 与预览窗口一致：平铺时以页面宽度的 22% 作为基准宽度，再叠加 size_scale
                scaled_w = max(16, rect.width * 0.22 * size_scale)
                scaled_h = max(16, scaled_w * img_h / max(1, img_w))
                resized_cache = {}
                for cx, cy, row, col in self._iter_positions(
                    page_w=rect.width,
                    page_h=rect.height,
                    base_w=scaled_w,
                    base_h=scaled_h,
                    spacing_scale=spacing_scale,
                    tile_layout=tile_layout,
                ):
                    factor = self._tile_size_factor(
                        page_seed=page_seed,
                        row=row,
                        col=col,
                        enabled=random_size,
                        strength=random_strength,
                    )
                    cur_w = max(10, int(scaled_w * factor))
                    cur_h = max(10, int(scaled_h * factor))
                    key = (cur_w, cur_h)
                    if key not in resized_cache:
                        render_img = pil_img.resize((cur_w, cur_h), PILImage.LANCZOS)
                        buf = io.BytesIO()
                        render_img.save(buf, format="PNG")
                        resized_cache[key] = buf.getvalue()
                    x = cx - cur_w / 2
                    y = cy - cur_h / 2
                    target = fitz.Rect(x, y, x + cur_w, y + cur_h)
                    page.insert_image(target, stream=resized_cache[key], overlay=True)
            else:
                # 与预览窗口一致：单点模式以页面宽度的 33% 作为基准宽度
                scaled_w = max(16, rect.width * 0.33 * size_scale)
                scaled_h = max(16, scaled_w * img_h / max(1, img_w))
                buf = io.BytesIO()
                pil_img.resize((max(10, int(scaled_w)), max(10, int(scaled_h))), PILImage.LANCZOS).save(buf, format="PNG")
                x0, y0 = self._single_anchor_xy(
                    rect=rect,
                    position=position,
                    item_w=max(10, int(scaled_w)),
                    item_h=max(10, int(scaled_h)),
                )
                target = fitz.Rect(x0, y0, x0 + max(10, int(scaled_w)), y0 + max(10, int(scaled_h)))
                page.insert_image(target, stream=buf.getvalue())

        except Exception as e:
            logging.error(f"添加图片水印失败: {e}")
            raise

    @staticmethod
    def _is_tile_mode(position):
        return str(position).startswith("tile")

    @staticmethod
    def _resolve_tile_layout(position, layout):
        pos = str(position or "")
        if pos.startswith("tile-"):
            return pos.split("-", 1)[1]
        return layout if layout in ("grid", "diag", "row", "col") else "grid"

    @staticmethod
    def _tile_size_factor(page_seed, row, col, enabled, strength):
        if not enabled:
            return 1.0
        spread = max(0.0, min(1.0, float(strength)))
        if spread <= 0.0:
            return 1.0
        rnd = random.Random((row + 1) * 9176 + (col + 1) * 101 + 1000003)
        return max(0.25, 1.0 + (rnd.random() * 2.0 - 1.0) * spread)

    @staticmethod
    def _iter_positions(page_w, page_h, base_w, base_h, spacing_scale, tile_layout):
        spacing = max(0.5, min(2.0, float(spacing_scale)))
        gap_x = max(base_w * 1.5, 90.0) * spacing
        gap_y = max(base_h * 1.7, 90.0) * spacing
        if tile_layout == "row":
            gap_y *= 2.2
        elif tile_layout == "col":
            gap_x *= 2.2
        y = gap_y * 0.5
        row = 0
        while y < page_h + base_h:
            x = gap_x * 0.5 + (gap_x * 0.5 if (tile_layout == "diag" and row % 2 == 1) else 0.0)
            col = 0
            while x < page_w + base_w:
                yield x, y, row, col
                x += gap_x
                col += 1
            y += gap_y
            row += 1

    @staticmethod
    def _normalize_color255(color):
        try:
            if isinstance(color, (list, tuple)) and len(color) == 3:
                vals = [float(color[0]), float(color[1]), float(color[2])]
            else:
                vals = [0.6, 0.6, 0.6]
        except Exception:
            vals = [0.6, 0.6, 0.6]
        out = []
        for v in vals:
            if v <= 1.0:
                out.append(int(max(0, min(255, round(v * 255)))))
            else:
                out.append(int(max(0, min(255, round(v)))))
        return tuple(out)

    @staticmethod
    def _load_font(font_px):
        candidates = [
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            r"C:\Windows\Fonts\simsun.ttc",
            r"C:\Windows\Fonts\arial.ttf",
        ]
        for fp in candidates:
            if os.path.exists(fp):
                try:
                    return PILImageFont.truetype(fp, int(font_px))
                except Exception:
                    pass
        return PILImageFont.load_default()

    def _render_text_stamp(self, text, font_px, color255, opacity, rotation):
        font = self._load_font(max(8, int(font_px)))
        tmp = PILImage.new("RGBA", (10, 10), (0, 0, 0, 0))
        d0 = PILImageDraw.Draw(tmp)
        bbox = d0.textbbox((0, 0), text, font=font)
        w = max(1, bbox[2] - bbox[0])
        h = max(1, bbox[3] - bbox[1])
        pad = max(4, int(font_px * 0.3))
        img = PILImage.new("RGBA", (w + pad * 2, h + pad * 2), (0, 0, 0, 0))
        d1 = PILImageDraw.Draw(img)
        d1.text(
            (pad - bbox[0], pad - bbox[1]),
            text,
            font=font,
            fill=(
                int(color255[0]),
                int(color255[1]),
                int(color255[2]),
                int(255 * max(0.01, min(1.0, float(opacity)))),
            ),
        )
        if abs(float(rotation)) > 0.01:
            img = img.rotate(float(rotation), expand=True, resample=PILImage.BICUBIC)
        return img

    @staticmethod
    def _pil_to_png_bytes(pil_img):
        buf = io.BytesIO()
        pil_img.save(buf, format="PNG")
        return buf.getvalue()

    @staticmethod
    def _parse_pages_str(pages_str, total_pages):
        """解析页码字符串（1-based）为 0-based 索引列表。"""
        if not pages_str or not str(pages_str).strip():
            return []

        pages = set()
        text = str(pages_str).strip()
        text = text.replace("～", "-").replace("~", "-").replace("—", "-").replace("–", "-")

        for part in re.split(r"[^0-9\-]+", text):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                seg = part.split("-", 1)
                if len(seg) != 2:
                    return None
                try:
                    start = int(seg[0].strip())
                    end = int(seg[1].strip())
                except ValueError:
                    return None
                if start < 1 or end < 1 or start > end:
                    return None
                for p in range(start, end + 1):
                    idx = p - 1
                    if 0 <= idx < int(total_pages):
                        pages.add(idx)
            else:
                try:
                    p = int(part)
                except ValueError:
                    return None
                if p < 1:
                    return None
                idx = p - 1
                if 0 <= idx < int(total_pages):
                    pages.add(idx)
        return sorted(pages)

    @staticmethod
    def _single_anchor_xy(rect, position, item_w, item_h):
        margin = 20
        if position == 'center':
            return (rect.width - item_w) / 2, (rect.height - item_h) / 2
        if position == 'top-center':
            return (rect.width - item_w) / 2, margin
        if position == 'bottom-center':
            return (rect.width - item_w) / 2, rect.height - item_h - margin
        if position == 'top-left':
            return margin, margin
        if position == 'top-right':
            return rect.width - item_w - margin, margin
        if position == 'bottom-left':
            return margin, rect.height - item_h - margin
        if position == 'bottom-right':
            return rect.width - item_w - margin, rect.height - item_h - margin
        return (rect.width - item_w) / 2, (rect.height - item_h) / 2

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
        elif position == 'top-center':
            return fitz.Point((rect.width - text_w) / 2, 30 + text_h)
        elif position == 'bottom-center':
            return fitz.Point((rect.width - text_w) / 2, rect.height - 30)
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
