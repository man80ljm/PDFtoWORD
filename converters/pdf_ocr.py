"""
PDF OCR（生成可搜索PDF）

将扫描版PDF的每一页通过OCR识别文字，然后在原PDF上叠加
一层透明（不可见）文字，使PDF变为可搜索、可复制文字的版本。
原始页面外观完全不变。

需要百度OCR API Key。
通过 on_progress 回调报告进度，不直接操作UI。
"""

import base64
import io
import logging
import os
import time
from datetime import datetime

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

try:
    from PIL import Image, ImageOps, ImageFilter
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class PDFOCRConverter:
    """PDF OCR转换器 — 生成可搜索PDF，与 UI 完全解耦。

    用法::

        converter = PDFOCRConverter(on_progress=my_callback)
        result = converter.convert("scanned.pdf", api_key="xxx", secret_key="yyy")
    """

    # 百度通用文字识别(含位置)
    OCR_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate"
    TOKEN_URL = "https://aip.baidubce.com/oauth/2.0/token"

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)
        self._access_token = None
        self._token_time = 0

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, input_file, output_path=None,
                api_key='', secret_key='',
                start_page=None, end_page=None, dpi=None, ocr_mode="平衡",
                skip_text_pages=True, min_existing_text_chars=24):
        """对扫描版PDF进行OCR并生成可搜索PDF。

        Args:
            input_file: 输入PDF路径
            output_path: 输出路径，None则自动生成
            api_key: 百度OCR API Key
            secret_key: 百度OCR Secret Key
            start_page: 起始页(1-based)，None=第1页
            end_page: 结束页(1-based)，None=最后一页
            dpi: 渲染DPI（可选，未传时按 ocr_mode 自动设置）
            ocr_mode: OCR质量模式（快速/平衡/高精）

        Returns:
            dict: success, message, output_file, page_count, words_count
        """
        result = {
            'success': False, 'message': '',
            'output_file': '', 'page_count': 0,
            'words_count': 0,
            'skipped_text_pages': 0,
        }

        if not FITZ_AVAILABLE:
            result['message'] = "PyMuPDF (fitz) 未安装！"
            return result

        if not REQUESTS_AVAILABLE:
            result['message'] = "requests 未安装！"
            return result

        if not api_key or not secret_key:
            result['message'] = "请先在设置中配置百度OCR API Key和Secret Key！"
            return result

        if not input_file or not os.path.exists(input_file):
            result['message'] = "请先选择PDF文件！"
            return result

        if not output_path:
            dir_path = os.path.dirname(input_file)
            basename = os.path.splitext(os.path.basename(input_file))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(
                dir_path, f"{basename}_可搜索_{timestamp}.pdf")

        profile = self._get_ocr_mode_profile(ocr_mode)
        if isinstance(dpi, (int, float)) and dpi > 0:
            render_dpi = int(dpi)
        else:
            render_dpi = profile["dpi"]
        retry_dpi = max(render_dpi, profile["retry_dpi"])
        retry_score = profile["retry_score"]

        try:
            # 获取 access_token
            self._report(percent=0, progress_text="正在连接百度OCR...",
                         status_text="获取API凭证")
            token = self._get_access_token(api_key, secret_key)

            doc = fitz.open(input_file)
            total_pages = len(doc)

            if total_pages == 0:
                doc.close()
                result['message'] = "PDF文件无内容"
                return result

            # 确定页范围
            s_idx = 0
            e_idx = total_pages
            if start_page is not None:
                s_idx = max(0, min(start_page - 1, total_pages - 1))
            if end_page is not None:
                e_idx = max(s_idx + 1, min(end_page, total_pages))

            pages_to_process = e_idx - s_idx
            total_words = 0

            for i, page_idx in enumerate(range(s_idx, e_idx)):
                page = doc[page_idx]
                page_num = page_idx + 1

                # 进度
                percent = int((i / pages_to_process) * 90)
                self._report(
                    percent=percent,
                    progress_text=f"OCR识别第 {page_num} 页... ({percent}%)",
                    status_text=f"第 {page_num}/{total_pages} 页"
                )

                # 已有可搜索文本的页直接跳过，可显著降低耗时
                if skip_text_pages:
                    existing_text = page.get_text("text") or ""
                    if self._has_enough_text(existing_text, min_existing_text_chars):
                        result['skipped_text_pages'] += 1
                        continue

                # 渲染页面为图片
                pix = page.get_pixmap(dpi=render_dpi)
                img_bytes = pix.tobytes("png")

                # API 频率控制
                if i > 0:
                    time.sleep(0.5)

                # 调用 OCR 获取带位置的文字
                try:
                    words_with_loc = self._ocr_with_location(img_bytes, token, render_dpi)
                except Exception as e:
                    logging.warning(f"第{page_num}页OCR失败: {e}")
                    continue

                # 低置信度时自动提高DPI重试一次
                if self._score_loc_words(words_with_loc) < retry_score and render_dpi < retry_dpi:
                    try:
                        pix_hi = page.get_pixmap(dpi=retry_dpi)
                        words_hi = self._ocr_with_location(
                            pix_hi.tobytes("png"), token, retry_dpi
                        )
                        if self._score_loc_words(words_hi) > self._score_loc_words(words_with_loc):
                            words_with_loc = words_hi
                    except Exception as e:
                        logging.debug(f"第{page_num}页高DPI重试失败: {e}")

                if not words_with_loc:
                    continue

                # 在页面上叠加不可见文字
                page_rect = page.rect
                for word_info in words_with_loc:
                    text = word_info['text']
                    x0 = word_info['x']
                    y0 = word_info['y']
                    w = word_info['w']
                    h = word_info['h']

                    if not text.strip():
                        continue

                    # 根据文字区域高度估算字号
                    fontsize = max(h * 0.8, 4)

                    # 插入不可见文字（透明色）
                    # render_mode=3 = invisible text（PDF标准的不可见渲染模式）
                    text_point = fitz.Point(x0, y0 + h * 0.85)
                    try:
                        rc = page.insert_text(
                            text_point,
                            text,
                            fontsize=fontsize,
                            fontname="china-s",
                            color=(0, 0, 0),
                            render_mode=3,  # 3 = invisible
                        )
                        if rc >= 0:
                            total_words += len(text)
                    except Exception as e:
                        logging.debug(f"文字插入失败: {e}")

            # 保存
            self._report(percent=92, progress_text="正在保存...",
                         status_text="写入可搜索PDF")

            doc.save(output_path, garbage=3, deflate=True)
            doc.close()

            result['success'] = True
            result['output_file'] = output_path
            result['page_count'] = pages_to_process
            result['words_count'] = total_words
            result['message'] = (
                f"OCR完成！\n"
                f"处理了 {pages_to_process} 页，识别 {total_words} 个字符\n"
                f"跳过已有文本页 {result['skipped_text_pages']} 页\n"
                f"PDF已变为可搜索版本，外观不变"
            )
            self._report(percent=100, progress_text="OCR完成！")

        except Exception as e:
            logging.error(f"PDF OCR失败: {e}", exc_info=True)
            result['message'] = f"OCR失败：{str(e)}"

        return result

    def _get_access_token(self, api_key, secret_key):
        """获取百度API access_token"""
        if self._access_token and (time.time() - self._token_time) < 86400 * 25:
            return self._access_token
        params = {
            "grant_type": "client_credentials",
            "client_id": api_key,
            "client_secret": secret_key,
        }
        resp = requests.post(self.TOKEN_URL, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if "access_token" not in data:
            raise RuntimeError(
                f"百度API认证失败: {data.get('error_description', data)}")
        self._access_token = data["access_token"]
        self._token_time = time.time()
        return self._access_token

    def _ocr_with_location(self, image_bytes, token, dpi):
        """调用百度通用文字识别（含位置），返回带坐标的文字列表。

        返回格式: [{'text': str, 'x': float, 'y': float, 'w': float, 'h': float}, ...]
        坐标已从像素坐标转换为PDF页面坐标（72 DPI 基准）。
        """
        best_words = self._ocr_with_location_once(image_bytes, token, dpi)
        best_score = self._score_loc_words(best_words)
        logging.info(f"OCR-with-location pass#1 score={best_score}, words={len(best_words)}")

        # 低置信度页触发预处理重试
        if best_score < 120 and PIL_AVAILABLE:
            for mode in ("normal", "strong"):
                try:
                    normalized = self._normalize_scan_image(image_bytes, mode=mode)
                    words2 = self._ocr_with_location_once(normalized, token, dpi)
                    score2 = self._score_loc_words(words2)
                    logging.info(
                        f"OCR-with-location retry[{mode}] score={score2}, words={len(words2)}"
                    )
                    if score2 > best_score:
                        best_words = words2
                        best_score = score2
                except Exception as e:
                    logging.debug(f"OCR-with-location retry failed ({mode}): {e}")

        return best_words

    @staticmethod
    def _score_loc_words(words):
        if not words:
            return 0
        chars = sum(len((w.get("text") or "").strip()) for w in words)
        lines = sum(1 for w in words if (w.get("text") or "").strip())
        return lines * 10 + chars

    @staticmethod
    def _normalize_scan_image(image_bytes, mode="normal"):
        """对扫描页做轻量增强，返回 PNG bytes。"""
        if not PIL_AVAILABLE:
            return image_bytes
        try:
            img = Image.open(io.BytesIO(image_bytes))
            gray = ImageOps.grayscale(img)
            if mode == "strong":
                gray = ImageOps.autocontrast(gray, cutoff=2)
                gray = gray.filter(ImageFilter.MedianFilter(size=3))
                gray = gray.point(lambda p: 255 if p > 170 else 0)
            else:
                gray = ImageOps.autocontrast(gray, cutoff=1)
                gray = gray.filter(ImageFilter.SHARPEN)

            out = io.BytesIO()
            gray.save(out, format="PNG")
            return out.getvalue()
        except Exception as e:
            logging.debug(f"Normalize scan image failed ({mode}): {e}")
            return image_bytes

    def _ocr_with_location_once(self, image_bytes, token, dpi):
        # 压缩图片（百度API限制4MB）
        compressed = self._compress_for_api(image_bytes)
        img_b64 = base64.b64encode(compressed).decode()

        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "image": img_b64,
            "language_type": "CHN_ENG",
            "detect_direction": "true",
            "paragraph": "false",
            "recognize_granularity": "small",
        }

        resp = requests.post(
            f"{self.OCR_URL}?access_token={token}",
            headers=headers, data=data, timeout=60
        )
        resp.raise_for_status()
        result = resp.json()

        if "error_code" in result:
            raise RuntimeError(
                f"OCR失败[{result.get('error_code')}]: {result.get('error_msg')}")

        # 像素坐标 → PDF坐标 的缩放比例
        # 图片以 dpi 渲染，PDF标准为 72 DPI
        scale = 72.0 / dpi

        words = []
        for item in result.get("words_result", []):
            text = item.get("words", "")
            loc = item.get("location", {})
            if not text or not loc:
                continue

            # 百度返回的是像素坐标
            px_left = loc.get("left", 0)
            px_top = loc.get("top", 0)
            px_width = loc.get("width", 100)
            px_height = loc.get("height", 20)

            words.append({
                'text': text,
                'x': px_left * scale,
                'y': px_top * scale,
                'w': px_width * scale,
                'h': px_height * scale,
            })

        return words

    @staticmethod
    def _compress_for_api(image_bytes, max_size=3 * 1024 * 1024):
        """压缩图片以满足百度API大小限制"""
        try:
            from PIL import Image as PILImage
        except ImportError:
            return image_bytes

        if len(image_bytes) <= max_size:
            return image_bytes

        img = PILImage.open(io.BytesIO(image_bytes))
        if img.mode == 'RGBA':
            img = img.convert('RGB')

        for quality in [85, 70, 55, 40]:
            buf = io.BytesIO()
            img.save(buf, 'JPEG', quality=quality)
            data = buf.getvalue()
            if len(base64.b64encode(data)) <= max_size:
                return data

        # 还是太大，缩小尺寸
        w, h = img.size
        img = img.resize((w // 2, h // 2), PILImage.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, 'JPEG', quality=70)
        return buf.getvalue()

    @staticmethod
    def _has_enough_text(text, min_chars=24):
        raw = (text or "").strip()
        if not raw:
            return False
        compact = "".join(raw.split())
        if len(compact) < min_chars:
            return False
        effective = sum(1 for ch in compact if ch.isalnum() or ("\u4e00" <= ch <= "\u9fff"))
        return effective >= max(12, min_chars // 2)

    @staticmethod
    def _get_ocr_mode_profile(ocr_mode):
        mode = (ocr_mode or "平衡").strip()
        profiles = {
            "快速": {"dpi": 220, "retry_dpi": 300, "retry_score": 100},
            "平衡": {"dpi": 300, "retry_dpi": 360, "retry_score": 120},
            "高精": {"dpi": 360, "retry_dpi": 420, "retry_score": 140},
        }
        return profiles.get(mode, profiles["平衡"])
