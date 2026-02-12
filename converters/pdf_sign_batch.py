"""Batch PDF signature converter."""

import logging
import os
from datetime import datetime

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    from PIL import Image  # noqa: F401
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

from converters.pdf_stamp_batch import PDFBatchStampConverter


class PDFBatchSignConverter:
    """Batch PDF sign converter (UI-decoupled)."""

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)
        self._stamp_helper = PDFBatchStampConverter(on_progress=None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(self, files, signature_items, remove_white_bg=False):
        result = {
            "success": False,
            "message": "",
            "output_files": [],
            "file_count": 0,
            "page_count": 0,
            "error_count": 0,
            "errors": [],
            "skipped_page_filtered": 0,
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

        normalized_items = self._normalize_items(signature_items)
        if not normalized_items:
            result["message"] = "No valid signature placements"
            return result

        # image cache for performance
        img_bytes_cache = {}
        img_size_cache = {}

        readable_files = []
        for f in files:
            try:
                d = fitz.open(f)
                if d.is_encrypted and not d.authenticate(""):
                    d.close()
                    result["errors"].append(f"Encrypted PDF skipped: {f}")
                    continue
                readable_files.append(f)
                d.close()
            except Exception as e:
                result["errors"].append(f"Open failed: {f} ({e})")

        if not readable_files:
            result["message"] = "No readable PDF files"
            result["error_count"] = len(result["errors"])
            return result

        for file_idx, pdf_path in enumerate(readable_files, 1):
            doc = None
            try:
                doc = fitz.open(pdf_path)
                page_count = len(doc)
                applied_pages = set()

                for item in normalized_items:
                    target_idx = item["page"] - 1
                    if target_idx < 0 or target_idx >= page_count:
                        continue

                    img_key = (item["image_path"], int(item["opacity"] * 1000), bool(remove_white_bg))
                    if img_key not in img_bytes_cache:
                        img_bytes_cache[img_key] = self._stamp_helper._image_with_opacity(
                            item["image_path"],
                            opacity=item["opacity"],
                            remove_white_bg=bool(remove_white_bg),
                        )
                    img_bytes = img_bytes_cache[img_key]

                    if img_key not in img_size_cache:
                        img_size_cache[img_key] = self._stamp_helper._image_size_from_bytes(img_bytes)
                    iw, ih = img_size_cache[img_key]

                    page = doc[target_idx]
                    pr = page.rect
                    rw = max(8.0, pr.width * item["size_ratio"])
                    rh = max(8.0, rw * ih / max(1, iw))

                    cx = pr.width * item["x_ratio"]
                    cy = pr.height * item["y_ratio"]
                    x = max(0.0, min(cx - rw / 2.0, max(0.0, pr.width - rw)))
                    y = max(0.0, min(cy - rh / 2.0, max(0.0, pr.height - rh)))
                    rect = fitz.Rect(x, y, x + rw, y + rh)
                    page.insert_image(rect, stream=img_bytes, keep_proportion=True, overlay=True)
                    applied_pages.add(target_idx)

                if not applied_pages:
                    result["skipped_page_filtered"] += 1
                    result["errors"].append(
                        f"Skipped (no configured pages in file): {os.path.basename(pdf_path)}"
                    )
                    continue

                out_path = self._make_output_path(pdf_path, suffix="签名")
                doc.save(out_path, garbage=3, deflate=True)
                result["output_files"].append(out_path)
                result["file_count"] += 1
                result["page_count"] += len(applied_pages)

            except Exception as e:
                logging.error("Sign failed: %s: %s", pdf_path, e, exc_info=True)
                result["errors"].append(f"Sign failed: {os.path.basename(pdf_path)} ({e})")
            finally:
                if doc is not None:
                    doc.close()
                pct = int((file_idx / max(1, len(readable_files))) * 100)
                self._report(
                    pct,
                    progress_text=f"Signing {file_idx}/{len(readable_files)}: {os.path.basename(pdf_path)}",
                    status_text=f"Processed {file_idx}/{len(readable_files)} files",
                )

        result["error_count"] = len(result["errors"])
        result["success"] = result["file_count"] > 0
        if result["success"]:
            result["message"] = (
                "Batch signing completed\n"
                f"Success files: {result['file_count']}\n"
                f"Signed pages: {result['page_count']}\n"
                f"Skipped by page config: {result['skipped_page_filtered']}\n"
                f"Warnings: {result['error_count']}"
            )
        else:
            result["message"] = "Batch signing failed"
        self._report(100, progress_text="Batch signing completed")
        return result

    @staticmethod
    def _normalize_items(signature_items):
        out = []
        items = signature_items if isinstance(signature_items, list) else []
        for it in items:
            if not isinstance(it, dict):
                continue
            image_path = os.path.abspath(str(it.get("image_path", "")).strip())
            if not image_path or not os.path.exists(image_path):
                continue
            try:
                page = int(it.get("page", 0))
            except Exception:
                page = 0
            if page < 1:
                continue
            x_ratio = PDFBatchStampConverter._clamp(it.get("x_ratio", 0.85), 0.0, 1.0, 0.85)
            y_ratio = PDFBatchStampConverter._clamp(it.get("y_ratio", 0.85), 0.0, 1.0, 0.85)
            size_ratio = PDFBatchStampConverter._clamp(it.get("size_ratio", 0.18), 0.03, 0.7, 0.18)
            opacity = PDFBatchStampConverter._clamp(it.get("opacity", 0.85), 0.05, 1.0, 0.85)
            out.append({
                "page": page,
                "image_path": image_path,
                "x_ratio": x_ratio,
                "y_ratio": y_ratio,
                "size_ratio": size_ratio,
                "opacity": opacity,
            })
        return out

    @staticmethod
    def _make_output_path(pdf_path, suffix="签名"):
        base, _ = os.path.splitext(pdf_path)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base}_{suffix}_{ts}.pdf"
