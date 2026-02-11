"""
PDF page reorder / rotate / reverse utility.

Supported modes:
- reorder: reorder pages by a full page sequence (e.g. "3,1,2,4-6")
- rotate: rotate selected pages by 90/180/270 degrees
- reverse: reverse all pages
"""

import logging
import os
from datetime import datetime

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False


class PDFReorderConverter:
    """PDF page reorder / rotate / reverse converter (UI-decoupled)."""

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(
        self,
        input_file,
        mode="reorder",
        reorder_pages="",
        rotate_pages="",
        rotate_angle=90,
        output_path=None,
    ):
        result = {
            "success": False,
            "message": "",
            "output_file": "",
            "page_count": 0,
        }

        if not FITZ_AVAILABLE:
            result["message"] = "PyMuPDF (fitz) is not installed. Run: pip install PyMuPDF"
            return result

        if not input_file or not os.path.exists(input_file):
            result["message"] = f"Input file not found: {input_file}"
            return result

        try:
            doc = fitz.open(input_file)
        except Exception as e:
            result["message"] = f"Failed to open PDF: {e}"
            return result

        try:
            if doc.is_encrypted and not doc.authenticate(""):
                result["message"] = "Encrypted PDF cannot be opened with empty password"
                doc.close()
                return result

            total_pages = len(doc)
            if total_pages <= 0:
                result["message"] = "PDF has no pages"
                doc.close()
                return result

            if not output_path:
                dirname = os.path.dirname(input_file)
                basename = os.path.splitext(os.path.basename(input_file))[0]
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                suffix = {
                    "reorder": "页面重排",
                    "rotate": "页面旋转",
                    "reverse": "页面倒序",
                }.get(mode, "页面处理")
                output_path = os.path.join(dirname, f"{basename}_{suffix}_{ts}.pdf")

            if mode == "reorder":
                seq, err = self._parse_reorder_sequence(reorder_pages, total_pages)
                if err:
                    doc.close()
                    result["message"] = err
                    return result
                self._report(10, "Preparing reordered pages...", "Reordering pages")
                out_doc = fitz.open()
                for idx, page_no in enumerate(seq, 1):
                    out_doc.insert_pdf(doc, from_page=page_no, to_page=page_no)
                    if idx % 5 == 0 or idx == len(seq):
                        pct = 10 + int(idx / len(seq) * 80)
                        self._report(pct, f"Reordering {idx}/{len(seq)}", "Reordering pages")
                out_doc.save(output_path, garbage=3, deflate=True)
                out_doc.close()
                result["message"] = "Page reorder completed"

            elif mode == "rotate":
                angle = int(rotate_angle)
                if angle not in (90, 180, 270):
                    doc.close()
                    result["message"] = "Rotate angle must be 90 / 180 / 270"
                    return result
                pages, err = self._parse_pages_str(rotate_pages, total_pages)
                if err:
                    doc.close()
                    result["message"] = err
                    return result
                if not pages:
                    pages = list(range(total_pages))

                self._report(10, "Preparing rotation...", "Rotating pages")
                for idx, p in enumerate(pages, 1):
                    page = doc[p]
                    old_rot = page.rotation or 0
                    page.set_rotation((old_rot + angle) % 360)
                    if idx % 5 == 0 or idx == len(pages):
                        pct = 10 + int(idx / len(pages) * 80)
                        self._report(pct, f"Rotating {idx}/{len(pages)}", "Rotating pages")
                doc.save(output_path, garbage=3, deflate=True)
                result["message"] = "Page rotation completed"

            elif mode == "reverse":
                self._report(10, "Preparing reverse order...", "Reversing pages")
                out_doc = fitz.open()
                for idx, p in enumerate(range(total_pages - 1, -1, -1), 1):
                    out_doc.insert_pdf(doc, from_page=p, to_page=p)
                    if idx % 5 == 0 or idx == total_pages:
                        pct = 10 + int(idx / total_pages * 80)
                        self._report(pct, f"Reversing {idx}/{total_pages}", "Reversing pages")
                out_doc.save(output_path, garbage=3, deflate=True)
                out_doc.close()
                result["message"] = "Page reverse completed"

            else:
                doc.close()
                result["message"] = f"Unsupported mode: {mode}"
                return result

            result["success"] = True
            result["output_file"] = output_path
            result["page_count"] = total_pages
            self._report(100, "PDF page processing completed")
            doc.close()
            return result

        except Exception as e:
            logging.error(f"PDF page process failed: {e}", exc_info=True)
            try:
                doc.close()
            except Exception:
                pass
            result["message"] = f"Page process failed: {e}"
            return result

    @staticmethod
    def _parse_pages_str(pages_str, total_pages):
        if not pages_str or not pages_str.strip():
            return [], None
        pages = set()
        text = (pages_str or "").strip()
        text = text.replace("，", ",").replace("；", ",").replace("、", ",").replace(";", ",")
        text = text.replace("～", "-").replace("~", "-").replace("—", "-").replace("–", "-")
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
                    return None, f"Invalid range: {part}"
                if start < 1 or end < 1 or start > end:
                    return None, f"Invalid range: {part}"
                if end > total_pages:
                    return None, f"Page out of range (max {total_pages}): {part}"
                for p in range(start, end + 1):
                    pages.add(p - 1)
            else:
                try:
                    p = int(part)
                except ValueError:
                    return None, f"Invalid page number: {part}"
                if p < 1 or p > total_pages:
                    return None, f"Page out of range (max {total_pages}): {part}"
                pages.add(p - 1)
        return sorted(pages), None

    def _parse_reorder_sequence(self, pages_str, total_pages):
        if not pages_str or not pages_str.strip():
            return None, "Please enter full page order, e.g. 3,1,2,4-6"
        text = (pages_str or "").strip()
        text = text.replace("，", ",").replace("；", ",").replace("、", ",").replace(";", ",")
        text = text.replace("～", "-").replace("~", "-").replace("—", "-").replace("–", "-")
        sequence = []
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
                    return None, f"Invalid range: {part}"
                if start < 1 or end < 1 or start > end:
                    return None, f"Invalid range: {part}"
                if end > total_pages:
                    return None, f"Page out of range (max {total_pages}): {part}"
                for p in range(start, end + 1):
                    sequence.append(p - 1)
            else:
                try:
                    p = int(part)
                except ValueError:
                    return None, f"Invalid page number: {part}"
                if p < 1 or p > total_pages:
                    return None, f"Page out of range (max {total_pages}): {part}"
                sequence.append(p - 1)

        if len(sequence) != total_pages:
            return None, f"Page order must contain exactly {total_pages} pages"
        if len(set(sequence)) != total_pages:
            return None, "Page order contains duplicates"
        return sequence, None
