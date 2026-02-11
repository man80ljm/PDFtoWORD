"""
Batch PDF stamping converter.

Modes:
- seal: stamp image on selected pages
- qr: generate QR image and stamp on selected pages
- seam: split one stamp image across selected pages along page edge
- template: apply a JSON template with seal/qr/text elements
"""

import io
import json
import logging
import os
from datetime import datetime

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    from PIL import Image, ImageEnhance
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import qrcode
    QRCODE_AVAILABLE = True
except ImportError:
    QRCODE_AVAILABLE = False


class PDFBatchStampConverter:
    """Batch PDF stamp converter (UI-decoupled)."""

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(
        self,
        files,
        mode="seal",
        pages_str="",
        opacity=0.85,
        position="right_bottom",
        size_ratio=0.18,
        seal_image_path="",
        qr_text="",
        seam_side="right",
        seam_align="center",
        seam_overlap_ratio=0.25,
        template_path="",
    ):
        result = {
            "success": False,
            "message": "",
            "output_files": [],
            "file_count": 0,
            "page_count": 0,
            "error_count": 0,
            "errors": [],
        }

        if not FITZ_AVAILABLE:
            result["message"] = "PyMuPDF (fitz) is not installed. Run: pip install PyMuPDF"
            return result
        if not PIL_AVAILABLE:
            result["message"] = "Pillow is not installed. Run: pip install Pillow"
            return result
        if not files:
            result["message"] = "No input PDF files selected"
            return result

        mode = (mode or "seal").strip().lower()
        opacity = max(0.05, min(1.0, float(opacity or 0.85)))
        size_ratio = max(0.03, min(0.6, float(size_ratio or 0.18)))
        seam_overlap_ratio = max(0.05, min(0.95, float(seam_overlap_ratio or 0.25)))

        if mode in ("seal", "seam") and (not seal_image_path or not os.path.exists(seal_image_path)):
            result["message"] = f"Stamp image not found: {seal_image_path}"
            return result
        if mode == "qr":
            if not qr_text.strip():
                result["message"] = "QR text is empty"
                return result
            if not QRCODE_AVAILABLE:
                result["message"] = "qrcode is not installed. Run: pip install qrcode[pil]"
                return result
        if mode == "template":
            if not template_path or not os.path.exists(template_path):
                result["message"] = f"Template JSON not found: {template_path}"
                return result

        total_pages = 0
        readable_files = []
        for f in files:
            try:
                d = fitz.open(f)
                if d.is_encrypted and not d.authenticate(""):
                    d.close()
                    result["errors"].append(f"Encrypted PDF skipped: {f}")
                    continue
                total_pages += len(d)
                readable_files.append(f)
                d.close()
            except Exception as e:
                result["errors"].append(f"Open failed: {f} ({e})")

        if not readable_files:
            result["message"] = "No readable PDF files"
            result["error_count"] = len(result["errors"])
            return result

        seal_bytes = None
        if mode in ("seal", "seam"):
            seal_bytes = self._image_with_opacity(seal_image_path, opacity)

        template_obj = None
        if mode == "template":
            try:
                with open(template_path, "r", encoding="utf-8") as f:
                    template_obj = json.load(f)
            except Exception as e:
                result["message"] = f"Template JSON parse failed: {e}"
                return result

        processed_pages = 0
        for file_idx, pdf_path in enumerate(readable_files, 1):
            try:
                doc = fitz.open(pdf_path)
                page_count = len(doc)
                pages = self._parse_pages_str(pages_str, page_count)
                if pages is None:
                    doc.close()
                    result["errors"].append(f"Invalid page range for file: {os.path.basename(pdf_path)}")
                    continue
                if not pages:
                    pages = list(range(page_count))

                if mode == "seal":
                    self._apply_seal(doc, pages, seal_bytes, position, size_ratio)
                elif mode == "qr":
                    qr_bytes = self._make_qr_png_bytes(qr_text.strip(), opacity=opacity)
                    self._apply_seal(doc, pages, qr_bytes, position, size_ratio)
                elif mode == "seam":
                    self._apply_seam(doc, pages, seal_image_path, seam_side, seam_align, seam_overlap_ratio, opacity)
                elif mode == "template":
                    self._apply_template(doc, pages, template_obj, opacity_default=opacity, size_ratio_default=size_ratio)
                else:
                    doc.close()
                    result["errors"].append(f"Unsupported mode: {mode}")
                    continue

                out_path = self._make_output_path(pdf_path, suffix="盖章")
                doc.save(out_path, garbage=3, deflate=True)
                doc.close()
                result["output_files"].append(out_path)
                result["file_count"] += 1
                result["page_count"] += len(pages)

            except Exception as e:
                logging.error(f"Stamp failed: {pdf_path}: {e}", exc_info=True)
                result["errors"].append(f"Stamp failed: {os.path.basename(pdf_path)} ({e})")
            finally:
                processed_pages += 1
                pct = int((processed_pages / max(1, len(readable_files))) * 100)
                self._report(
                    pct,
                    progress_text=f"Stamping {file_idx}/{len(readable_files)}: {os.path.basename(pdf_path)}",
                    status_text=f"Processed {file_idx}/{len(readable_files)} files",
                )

        result["error_count"] = len(result["errors"])
        result["success"] = result["file_count"] > 0
        if result["success"]:
            result["message"] = (
                f"Batch stamping completed\n"
                f"Success files: {result['file_count']}\n"
                f"Stamped pages: {result['page_count']}\n"
                f"Warnings: {result['error_count']}"
            )
        else:
            result["message"] = "Batch stamping failed"
        self._report(100, progress_text="Batch stamping completed")
        return result

    def _apply_seal(self, doc, pages, image_bytes, position, size_ratio):
        img_size = self._image_size_from_bytes(image_bytes)
        for p in pages:
            page = doc[p]
            rect = self._build_rect(page.rect, img_size[0], img_size[1], position, size_ratio)
            page.insert_image(rect, stream=image_bytes, keep_proportion=True, overlay=True)

    def _apply_seam(self, doc, pages, image_path, side, align, overlap_ratio, opacity):
        src = Image.open(image_path).convert("RGBA")
        src = self._apply_alpha(src, opacity)
        n = max(1, len(pages))
        side = (side or "right").lower()
        align = (align or "center").lower()

        if side in ("left", "right"):
            step = src.width / n
            slices = []
            for i in range(n):
                x1 = int(round(i * step))
                x2 = int(round((i + 1) * step))
                x2 = max(x2, x1 + 1)
                slices.append(src.crop((x1, 0, x2, src.height)))
            for idx, p in enumerate(pages):
                page = doc[p]
                sl = slices[idx]
                sl_bytes = self._pil_to_png_bytes(sl)
                pr = page.rect
                target_w = pr.width * 0.14
                target_h = target_w * (sl.height / max(1, sl.width))
                y = self._aligned_y(pr.height, target_h, align)
                if side == "right":
                    x = pr.width - target_w * overlap_ratio
                else:
                    x = -target_w * (1.0 - overlap_ratio)
                r = fitz.Rect(x, y, x + target_w, y + target_h)
                page.insert_image(r, stream=sl_bytes, keep_proportion=True, overlay=True)
        else:
            step = src.height / n
            slices = []
            for i in range(n):
                y1 = int(round(i * step))
                y2 = int(round((i + 1) * step))
                y2 = max(y2, y1 + 1)
                slices.append(src.crop((0, y1, src.width, y2)))
            for idx, p in enumerate(pages):
                page = doc[p]
                sl = slices[idx]
                sl_bytes = self._pil_to_png_bytes(sl)
                pr = page.rect
                target_h = pr.height * 0.14
                target_w = target_h * (sl.width / max(1, sl.height))
                x = self._aligned_x(pr.width, target_w, align)
                if side == "bottom":
                    y = pr.height - target_h * overlap_ratio
                else:
                    y = -target_h * (1.0 - overlap_ratio)
                r = fitz.Rect(x, y, x + target_w, y + target_h)
                page.insert_image(r, stream=sl_bytes, keep_proportion=True, overlay=True)

    def _apply_template(self, doc, pages, template_obj, opacity_default=0.85, size_ratio_default=0.18):
        if isinstance(template_obj, dict):
            elems = template_obj.get("elements", [])
        elif isinstance(template_obj, list):
            elems = template_obj
        else:
            elems = []
        if not elems:
            return

        for p in pages:
            page = doc[p]
            page_no = p + 1
            pr = page.rect
            for e in elems:
                if not isinstance(e, dict):
                    continue
                scope = e.get("pages")
                if isinstance(scope, list) and page_no not in scope:
                    continue
                etype = str(e.get("type", "seal")).lower()
                x_ratio = float(e.get("x_ratio", 0.75))
                y_ratio = float(e.get("y_ratio", 0.75))
                w_ratio = float(e.get("w_ratio", size_ratio_default))
                h_ratio = float(e.get("h_ratio", 0.0))
                opacity = float(e.get("opacity", opacity_default))
                opacity = max(0.05, min(1.0, opacity))

                if etype == "seal":
                    image_path = e.get("image_path", "")
                    if not image_path or not os.path.exists(image_path):
                        continue
                    img_bytes = self._image_with_opacity(image_path, opacity)
                    iw, ih = self._image_size_from_bytes(img_bytes)
                    rw = pr.width * max(0.02, min(0.95, w_ratio))
                    rh = (rw * ih / max(1, iw)) if h_ratio <= 0 else pr.height * h_ratio
                    x = pr.width * x_ratio
                    y = pr.height * y_ratio
                    r = fitz.Rect(x, y, x + rw, y + rh)
                    page.insert_image(r, stream=img_bytes, keep_proportion=True, overlay=True)

                elif etype == "qr":
                    if not QRCODE_AVAILABLE:
                        continue
                    txt = str(e.get("text", "")).strip()
                    if not txt:
                        continue
                    qr_bytes = self._make_qr_png_bytes(txt, opacity=opacity)
                    iw, ih = self._image_size_from_bytes(qr_bytes)
                    rw = pr.width * max(0.02, min(0.95, w_ratio))
                    rh = (rw * ih / max(1, iw)) if h_ratio <= 0 else pr.height * h_ratio
                    x = pr.width * x_ratio
                    y = pr.height * y_ratio
                    r = fitz.Rect(x, y, x + rw, y + rh)
                    page.insert_image(r, stream=qr_bytes, keep_proportion=True, overlay=True)

                elif etype == "text":
                    txt = str(e.get("text", "")).strip()
                    if not txt:
                        continue
                    font_size = float(e.get("font_size", 12))
                    color = e.get("color", [1, 0, 0])
                    if not isinstance(color, (list, tuple)) or len(color) != 3:
                        color = [1, 0, 0]
                    rw = pr.width * max(0.04, min(0.95, w_ratio))
                    rh = pr.height * (h_ratio if h_ratio > 0 else 0.05)
                    x = pr.width * x_ratio
                    y = pr.height * y_ratio
                    r = fitz.Rect(x, y, x + rw, y + rh)
                    page.insert_textbox(
                        r, txt, fontsize=font_size,
                        color=(float(color[0]), float(color[1]), float(color[2])),
                        overlay=True
                    )

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
    def _make_output_path(input_file, suffix="盖章"):
        base = os.path.splitext(os.path.basename(input_file))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"{base}_{suffix}_{ts}.pdf"
        return os.path.join(os.path.dirname(input_file), out_name)

    @staticmethod
    def _apply_alpha(img_rgba, opacity):
        if img_rgba.mode != "RGBA":
            img_rgba = img_rgba.convert("RGBA")
        alpha = img_rgba.getchannel("A")
        alpha = ImageEnhance.Brightness(alpha).enhance(max(0.05, min(1.0, float(opacity))))
        img_rgba.putalpha(alpha)
        return img_rgba

    def _image_with_opacity(self, image_path, opacity):
        img = Image.open(image_path).convert("RGBA")
        img = self._apply_alpha(img, opacity)
        return self._pil_to_png_bytes(img)

    @staticmethod
    def _pil_to_png_bytes(img):
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    @staticmethod
    def _image_size_from_bytes(img_bytes):
        img = Image.open(io.BytesIO(img_bytes))
        return img.size[0], img.size[1]

    @staticmethod
    def _build_rect(page_rect, img_w, img_h, position, size_ratio):
        p = (position or "right_bottom").lower()
        margin = 10
        target_w = page_rect.width * max(0.03, min(0.7, size_ratio))
        target_h = target_w * (img_h / max(1, img_w))

        if p == "left_top":
            x = margin
            y = margin
        elif p == "right_top":
            x = page_rect.width - target_w - margin
            y = margin
        elif p == "left_bottom":
            x = margin
            y = page_rect.height - target_h - margin
        elif p == "center":
            x = (page_rect.width - target_w) / 2
            y = (page_rect.height - target_h) / 2
        else:
            x = page_rect.width - target_w - margin
            y = page_rect.height - target_h - margin

        return fitz.Rect(x, y, x + target_w, y + target_h)

    @staticmethod
    def _aligned_y(page_h, target_h, align):
        a = (align or "center").lower()
        if a == "top":
            return 6
        if a == "bottom":
            return max(0, page_h - target_h - 6)
        return max(0, (page_h - target_h) / 2)

    @staticmethod
    def _aligned_x(page_w, target_w, align):
        a = (align or "center").lower()
        if a == "left":
            return 6
        if a == "right":
            return max(0, page_w - target_w - 6)
        return max(0, (page_w - target_w) / 2)

    @staticmethod
    def _make_qr_png_bytes(text, opacity=1.0):
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=8,
            border=2,
        )
        qr.add_data(text)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
        if opacity < 0.999:
            alpha = img.getchannel("A")
            alpha = ImageEnhance.Brightness(alpha).enhance(max(0.05, min(1.0, float(opacity))))
            img.putalpha(alpha)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

