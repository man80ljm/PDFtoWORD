"""
PDF 批量文本/图片提取工具

支持：
- 批量 PDF 文件
- 页码范围字符串（如 1,3,5-10）
- 文本导出（txt/json/csv）
- 图片导出（原图提取，可选格式转换）
- 可选去重、按页目录
- 可选 OCR（无文本页）
"""

import csv
import hashlib
import io
import json
import logging
import os
import re
from datetime import datetime

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

from core.ocr_client import BaiduOCRClient, REQUESTS_AVAILABLE


class PDFBatchExtractConverter:
    """PDF 批量文本/图片提取转换器（与 UI 解耦）。"""

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(
        self,
        files,
        output_dir=None,
        pages_str="",
        extract_text=True,
        extract_images=True,
        text_format="txt",
        text_mode="merge",
        preserve_layout=True,
        ocr_enabled=False,
        ocr_mode="平衡",
        api_key=None,
        secret_key=None,
        image_per_page=False,
        image_dedupe=False,
        image_format="原格式",
        zip_output=False,
        keyword_filter="",
        regex_filter="",
        regex_enabled=False,
    ):
        """批量提取 PDF 文本/图片。

        Returns:
            dict: success, message, output_dir, output_zip, stats, errors
        """
        result = {
            "success": False,
            "message": "",
            "output_dir": "",
            "output_zip": "",
            "stats": {
                "file_count": 0,
                "page_count": 0,
                "text_pages": 0,
                "image_count": 0,
                "ocr_pages": 0,
                "skipped_files": 0,
            },
            "errors": [],
        }

        if not FITZ_AVAILABLE:
            result["message"] = "PyMuPDF (fitz) 未安装！\n请运行: pip install PyMuPDF"
            return result

        if not files:
            result["message"] = "未选择 PDF 文件"
            return result

        if extract_text is False and extract_images is False:
            result["message"] = "请至少选择文本或图片提取"
            return result

        if ocr_enabled and (not REQUESTS_AVAILABLE):
            result["message"] = "requests 未安装，无法使用 OCR"
            return result

        if ocr_enabled and (not api_key or not secret_key):
            result["message"] = "已启用 OCR，但未配置百度 OCR API Key/Secret Key"
            return result

        text_format_norm = (text_format or "txt").strip().lower()
        if text_format_norm == "excel":
            text_format_norm = "xlsx"

        if image_format != "原格式" and not PIL_AVAILABLE:
            result["message"] = "Pillow 未安装，无法进行图片格式转换"
            return result

        if text_format_norm == "xlsx" and not OPENPYXL_AVAILABLE:
            result["message"] = "openpyxl 未安装，无法导出 xlsx"
            return result

        # 过滤器准备
        keywords = []
        if keyword_filter and keyword_filter.strip():
            for part in re.split(r"[,，;\s]+", keyword_filter.strip()):
                token = part.strip()
                if token:
                    keywords.append(token)

        regex_obj = None
        if regex_enabled and regex_filter and regex_filter.strip():
            try:
                regex_obj = re.compile(regex_filter)
            except re.error as e:
                result["message"] = f"正则表达式无效: {e}"
                return result

        # 输出目录
        if not output_dir:
            first = files[0]
            base_dir = os.path.dirname(first)
            base_name = os.path.splitext(os.path.basename(first))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = os.path.join(base_dir, f"{base_name}_批量提取_{timestamp}")

        os.makedirs(output_dir, exist_ok=True)
        text_root = os.path.join(output_dir, "文本")
        image_root = os.path.join(output_dir, "图片")
        if extract_text:
            os.makedirs(text_root, exist_ok=True)
        if extract_images:
            os.makedirs(image_root, exist_ok=True)

        # 统计总页数
        total_pages = 0
        page_counts = {}
        for f in files:
            try:
                doc = fitz.open(f)
                if doc.is_encrypted:
                    # 尝试空密码
                    if not doc.authenticate(""):
                        doc.close()
                        result["errors"].append(f"加密PDF无法打开: {f}")
                        result["stats"]["skipped_files"] += 1
                        continue
                total_pages += len(doc)
                page_counts[f] = len(doc)
                doc.close()
            except Exception as e:
                result["errors"].append(f"无法读取PDF: {f} ({e})")
                result["stats"]["skipped_files"] += 1

        if total_pages == 0:
            result["message"] = "未能读取任何可用 PDF"
            return result

        # OCR 客户端
        ocr_client = None
        if ocr_enabled:
            ocr_client = BaiduOCRClient(api_key, secret_key)
        ocr_dpi = self._ocr_mode_to_dpi(ocr_mode)

        processed_pages = 0
        dedupe_hashes = set()

        all_text_rows = []  # for csv/json
        summary = []

        for file_idx, pdf_path in enumerate(files, 1):
            if pdf_path not in page_counts:
                continue

            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            per_pdf_text_dir = os.path.join(text_root, base_name)
            per_pdf_img_dir = os.path.join(image_root, base_name)
            if extract_text:
                os.makedirs(per_pdf_text_dir, exist_ok=True)
            if extract_images:
                os.makedirs(per_pdf_img_dir, exist_ok=True)

            try:
                doc = fitz.open(pdf_path)
                if doc.is_encrypted:
                    if not doc.authenticate(""):
                        doc.close()
                        result["errors"].append(f"加密PDF无法打开: {pdf_path}")
                        result["stats"]["skipped_files"] += 1
                        continue

                total = len(doc)
                pages = self._parse_pages_str(pages_str, total)
                if pages is None:
                    doc.close()
                    result["message"] = f"页码范围格式不正确: {pages_str}"
                    return result
                if not pages:
                    pages = list(range(total))

                per_pdf_text = []
                per_pdf_text_rows = []
                per_pdf_images = 0
                per_pdf_ocr_pages = 0

                for page_idx in pages:
                    if page_idx < 0 or page_idx >= total:
                        continue
                    page = doc[page_idx]

                    # 文本提取
                    page_text = ""
                    ocr_used = False
                    if extract_text:
                        raw_page_text = page.get_text("text") or ""
                        page_text = raw_page_text
                        if not preserve_layout:
                            page_text = self._clean_text(page_text)

                        # 智能OCR触发：不仅“纯空页”触发，也会在文本极少时触发
                        if ocr_enabled and self._needs_ocr_text_fallback(raw_page_text):
                            try:
                                pix = page.get_pixmap(dpi=ocr_dpi)
                                img_bytes = pix.tobytes("png")
                                lines = ocr_client.recognize_text(img_bytes)
                                page_text = "\n".join(lines) if lines else ""
                                ocr_used = bool(page_text.strip())
                            except Exception as e:
                                result["errors"].append(
                                    f"OCR失败: {base_name} 第{page_idx + 1}页 ({e})"
                                )

                        if self._match_text_filter(page_text, keywords, regex_obj):
                            per_pdf_text_rows.append({
                                "file": base_name,
                                "page": page_idx + 1,
                                "text": page_text,
                                "ocr": ocr_used,
                            })
                            per_pdf_text.append((page_idx + 1, page_text))

                    # 图片提取
                    if extract_images:
                        images = page.get_images(full=True)
                        for img_i, img in enumerate(images, 1):
                            xref = img[0]
                            try:
                                extracted = doc.extract_image(xref)
                            except Exception:
                                continue
                            if not extracted:
                                continue
                            img_bytes = extracted.get("image")
                            img_ext = extracted.get("ext", "bin")

                            if image_format != "原格式":
                                img_bytes, img_ext = self._convert_image_format(
                                    img_bytes, image_format
                                )

                            if image_dedupe:
                                h = hashlib.sha256(img_bytes).hexdigest()
                                if h in dedupe_hashes:
                                    continue
                                dedupe_hashes.add(h)

                            if image_per_page:
                                page_dir = os.path.join(per_pdf_img_dir, f"第{page_idx + 1}页")
                                os.makedirs(page_dir, exist_ok=True)
                                target_dir = page_dir
                            else:
                                target_dir = per_pdf_img_dir

                            filename = f"{base_name}_第{page_idx + 1}页_img{img_i}.{img_ext}"
                            out_path = os.path.join(target_dir, filename)
                            with open(out_path, "wb") as f:
                                f.write(img_bytes)
                            per_pdf_images += 1

                    processed_pages += 1
                    percent = int((processed_pages / total_pages) * 100)
                    self._report(
                        percent=percent,
                        progress_text=f"正在处理 {base_name} 第 {page_idx + 1}/{total} 页",
                        status_text=f"总进度 {processed_pages}/{total_pages} 页",
                    )

                    if ocr_used:
                        per_pdf_ocr_pages += 1

                # 写文本输出
                if extract_text:
                    if text_format_norm == "txt":
                        if text_mode == "per_page":
                            for page_no, text in per_pdf_text:
                                txt_path = os.path.join(
                                    per_pdf_text_dir, f"{base_name}_第{page_no}页.txt"
                                )
                                with open(txt_path, "w", encoding="utf-8") as f:
                                    f.write(text or "")
                        else:
                            txt_path = os.path.join(per_pdf_text_dir, f"{base_name}.txt")
                            with open(txt_path, "w", encoding="utf-8") as f:
                                for page_no, text in per_pdf_text:
                                    if text_mode == "merge":
                                        if text:
                                            f.write(text)
                                        f.write("\n")
                    elif text_format_norm == "csv":
                        all_text_rows.extend(per_pdf_text_rows)
                    elif text_format_norm == "xlsx":
                        all_text_rows.extend(per_pdf_text_rows)
                    else:
                        all_text_rows.extend(per_pdf_text_rows)

                summary.append({
                    "file": base_name,
                    "pages": len(pages),
                    "text_pages": len(per_pdf_text),
                    "image_count": per_pdf_images,
                    "ocr_pages": per_pdf_ocr_pages,
                })

                result["stats"]["page_count"] += len(pages)
                result["stats"]["text_pages"] += len(per_pdf_text)
                result["stats"]["image_count"] += per_pdf_images
                result["stats"]["ocr_pages"] += per_pdf_ocr_pages

                doc.close()

            except Exception as e:
                result["errors"].append(f"处理失败: {pdf_path} ({e})")

        # 汇总输出
        if extract_text and all_text_rows:
            if text_format_norm == "csv":
                csv_path = os.path.join(output_dir, "文本", "全文_汇总.csv")
                with open(csv_path, "w", encoding="utf-8", newline="") as f:
                    writer = csv.DictWriter(f, fieldnames=["file", "page", "text", "ocr"])
                    writer.writeheader()
                    for row in all_text_rows:
                        writer.writerow(row)
            elif text_format_norm == "json":
                json_path = os.path.join(output_dir, "文本", "全文_汇总.json")
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(all_text_rows, f, ensure_ascii=False, indent=2)
            elif text_format_norm == "xlsx":
                xlsx_path = os.path.join(output_dir, "文本", "全文_汇总.xlsx")
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "text_rows"
                ws.append(["file", "page", "text", "ocr"])
                for row in all_text_rows:
                    ws.append([
                        row.get("file", ""),
                        row.get("page", ""),
                        row.get("text", ""),
                        "1" if row.get("ocr") else "0",
                    ])
                wb.save(xlsx_path)

        # summary
        summary_path = os.path.join(output_dir, "summary.json")
        try:
            with open(summary_path, "w", encoding="utf-8") as f:
                json.dump(summary, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        # zip 输出（可选）
        output_zip = ""
        if zip_output:
            try:
                import zipfile

                zip_path = output_dir.rstrip("\\/") + ".zip"
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for root, _, files_in in os.walk(output_dir):
                        for name in files_in:
                            full = os.path.join(root, name)
                            rel = os.path.relpath(full, output_dir)
                            zf.write(full, rel)
                output_zip = zip_path
            except Exception as e:
                result["errors"].append(f"打包失败: {e}")

        result["success"] = True
        result["output_dir"] = output_dir
        result["output_zip"] = output_zip
        result["stats"]["file_count"] = len(files)
        result["message"] = self._build_message(result)

        self._report(percent=100, progress_text="批量提取完成")
        return result

    @staticmethod
    def _clean_text(text):
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        # 合并连续空白
        parts = text.split()
        return " ".join(parts)

    @staticmethod
    def _parse_pages_str(pages_str, total_pages):
        if not pages_str or not pages_str.strip():
            return []

        pages = set()
        text = pages_str.replace("，", ",").replace("；", ",").replace(";", ",")
        for part in text.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                seg = part.split("-", 1)
                try:
                    start = int(seg[0].strip())
                    end = int(seg[1].strip())
                except ValueError:
                    return None
                if start < 1 or end < 1 or start > end:
                    return None
                for p in range(start, end + 1):
                    if p <= total_pages:
                        pages.add(p - 1)
            else:
                try:
                    p = int(part)
                except ValueError:
                    return None
                if p < 1:
                    return None
                if p <= total_pages:
                    pages.add(p - 1)
        return sorted(pages)

    @staticmethod
    def _convert_image_format(img_bytes, target_format):
        target = target_format.upper()
        if target not in ("JPG", "JPEG", "PNG"):
            return img_bytes, "bin"
        img = Image.open(io.BytesIO(img_bytes))
        if target in ("JPG", "JPEG"):
            img = img.convert("RGB")
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=95)
            return buf.getvalue(), "jpg"
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue(), "png"

    @staticmethod
    def _build_message(result):
        stats = result.get("stats", {})
        msg = (
            "批量提取完成！\n"
            f"文件数: {stats.get('file_count', 0)}\n"
            f"处理页数: {stats.get('page_count', 0)}\n"
            f"文本页数: {stats.get('text_pages', 0)}\n"
            f"图片数量: {stats.get('image_count', 0)}\n"
            f"OCR页数: {stats.get('ocr_pages', 0)}"
        )
        if result.get("errors"):
            msg += f"\n\n有 {len(result['errors'])} 条警告/错误，详情见 summary.json"
        return msg

    @staticmethod
    def _match_text_filter(text, keywords, regex_obj):
        if not keywords and regex_obj is None:
            return True

        normalized = text or ""

        if keywords:
            hit_kw = any(k in normalized for k in keywords)
            if not hit_kw:
                return False

        if regex_obj is not None:
            return bool(regex_obj.search(normalized))

        return True

    @staticmethod
    def _needs_ocr_text_fallback(raw_text):
        """文本兜底触发条件：无文本或有效字符过少。"""
        text = (raw_text or "").strip()
        if not text:
            return True

        # 去掉空白后统计可见字符
        compact = re.sub(r"\s+", "", text)
        if len(compact) < 16:
            return True

        # 至少包含一定数量的中文/英文/数字字符，否则视为低质量文本
        effective = re.findall(r"[\u4e00-\u9fffA-Za-z0-9]", compact)
        return len(effective) < 12

    @staticmethod
    def _ocr_mode_to_dpi(ocr_mode):
        mode = (ocr_mode or "平衡").strip()
        mapping = {
            "快速": 220,
            "平衡": 300,
            "高精": 360,
        }
        return mapping.get(mode, 300)
