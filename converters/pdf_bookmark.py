"""
PDF bookmark utility.

Supported modes:
- add: add a single bookmark
- remove: remove bookmarks by level and/or keyword
- import_json: import bookmarks from JSON
- export_json: export bookmarks to JSON
- clear: remove all bookmarks
- auto: auto-generate bookmarks by heading pattern
"""

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


class PDFBookmarkConverter:
    """PDF bookmark converter (UI-decoupled)."""

    MODES = ("add", "remove", "import_json", "export_json", "clear", "auto")

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def convert(
        self,
        input_file,
        mode="add",
        output_path=None,
        json_path="",
        title="",
        page=1,
        level=1,
        remove_levels="",
        remove_keyword="",
        auto_pattern="",
        merge_existing=False,
    ):
        result = {
            "success": False,
            "message": "",
            "output_file": "",
            "output_json": "",
            "bookmark_count": 0,
        }

        if not FITZ_AVAILABLE:
            result["message"] = "PyMuPDF (fitz) 未安装！\n请运行: pip install PyMuPDF"
            return result

        if mode not in self.MODES:
            result["message"] = f"不支持的书签模式: {mode}"
            return result

        if not input_file or not os.path.exists(input_file):
            result["message"] = "输入PDF不存在"
            return result

        try:
            doc = fitz.open(input_file)
        except Exception as e:
            result["message"] = f"无法打开PDF: {e}"
            return result

        try:
            if doc.is_encrypted and not doc.authenticate(""):
                result["message"] = "加密PDF无法使用空密码打开"
                doc.close()
                return result

            total_pages = len(doc)
            if total_pages <= 0:
                result["message"] = "PDF没有页面"
                doc.close()
                return result

            existing_toc = doc.get_toc() or []
            new_toc = None

            if mode == "export_json":
                out_json = self._resolve_json_export_path(input_file, json_path)
                export_items = self._toc_to_dicts(existing_toc)
                with open(out_json, "w", encoding="utf-8") as f:
                    json.dump(export_items, f, ensure_ascii=False, indent=2)
                result["success"] = True
                result["output_json"] = out_json
                result["bookmark_count"] = len(export_items)
                result["message"] = f"书签已导出：{len(export_items)} 条"
                self._report(100, "书签导出完成")
                doc.close()
                return result

            if mode == "add":
                title_clean = (title or "").strip()
                if not title_clean:
                    result["message"] = "书签标题不能为空"
                    doc.close()
                    return result
                try:
                    page_i = int(page)
                    level_i = int(level)
                except Exception:
                    result["message"] = "页码和级别必须是数字"
                    doc.close()
                    return result
                if page_i < 1 or page_i > total_pages:
                    result["message"] = f"页码超出范围：{page_i}（共{total_pages}页）"
                    doc.close()
                    return result
                if level_i < 1:
                    result["message"] = "书签级别必须 >= 1"
                    doc.close()
                    return result

                new_toc = list(existing_toc)
                new_toc.append([level_i, title_clean, page_i])
                new_toc = self._sort_toc_by_page(new_toc)

            elif mode == "remove":
                levels, err = self._parse_levels(remove_levels)
                if err:
                    result["message"] = err
                    doc.close()
                    return result
                kw = (remove_keyword or "").strip()
                if not levels and not kw:
                    result["message"] = "请填写移除级别或关键词"
                    doc.close()
                    return result

                new_toc = []
                removed = 0
                for item in existing_toc:
                    lvl = int(item[0]) if len(item) > 0 else 1
                    txt = str(item[1]) if len(item) > 1 else ""
                    hit_level = bool(levels and lvl in levels)
                    hit_kw = bool(kw and kw in txt)
                    if hit_level or hit_kw:
                        removed += 1
                        continue
                    new_toc.append(item)
                result["message"] = f"已移除 {removed} 条书签"

            elif mode == "clear":
                new_toc = []
                result["message"] = "已清空全部书签"

            elif mode == "import_json":
                in_json = (json_path or "").strip()
                if not in_json or not os.path.exists(in_json):
                    result["message"] = "请先选择有效的JSON文件"
                    doc.close()
                    return result
                try:
                    with open(in_json, "r", encoding="utf-8") as f:
                        data = json.load(f)
                except Exception as e:
                    result["message"] = f"读取JSON失败: {e}"
                    doc.close()
                    return result

                imported, err = self._parse_import_toc(data, total_pages)
                if err:
                    result["message"] = err
                    doc.close()
                    return result

                if merge_existing:
                    new_toc = self._sort_toc_by_page(list(existing_toc) + imported)
                else:
                    new_toc = imported
                result["message"] = f"已导入 {len(imported)} 条书签"

            elif mode == "auto":
                pattern = (auto_pattern or "").strip()
                auto_items = self._auto_generate_toc(doc, pattern)
                if merge_existing:
                    new_toc = self._sort_toc_by_page(list(existing_toc) + auto_items)
                else:
                    new_toc = auto_items
                result["message"] = f"自动生成 {len(auto_items)} 条书签"

            out_pdf = output_path or self._resolve_pdf_output_path(input_file, mode)
            self._report(65, "正在写入书签...", "保存PDF中")
            doc.set_toc(new_toc or [])
            doc.save(out_pdf, garbage=3, deflate=True)
            doc.close()

            result["success"] = True
            result["output_file"] = out_pdf
            result["bookmark_count"] = len(new_toc or [])
            if mode not in ("remove", "clear", "import_json", "auto"):
                result["message"] = f"书签操作完成，当前共 {result['bookmark_count']} 条书签"
            else:
                result["message"] = f"{result['message']}\n当前共 {result['bookmark_count']} 条书签"
            self._report(100, "书签处理完成")
            return result

        except Exception as e:
            logging.error(f"PDF书签处理失败: {e}", exc_info=True)
            try:
                doc.close()
            except Exception:
                pass
            result["message"] = f"书签处理失败: {e}"
            return result

    @staticmethod
    def _resolve_pdf_output_path(input_file, mode):
        dirname = os.path.dirname(input_file)
        basename = os.path.splitext(os.path.basename(input_file))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        mode_map = {
            "add": "书签添加",
            "remove": "书签移除",
            "import_json": "书签导入",
            "clear": "书签清空",
            "auto": "书签自动生成",
        }
        suffix = mode_map.get(mode, "书签处理")
        return os.path.join(dirname, f"{basename}_{suffix}_{ts}.pdf")

    @staticmethod
    def _resolve_json_export_path(input_file, json_path):
        if json_path and json_path.strip():
            return json_path.strip()
        dirname = os.path.dirname(input_file)
        basename = os.path.splitext(os.path.basename(input_file))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(dirname, f"{basename}_书签_{ts}.json")

    @staticmethod
    def _sort_toc_by_page(toc_items):
        normalized = []
        for idx, item in enumerate(toc_items or []):
            if len(item) < 3:
                continue
            try:
                lvl = max(1, int(item[0]))
                title = str(item[1]).strip()
                page = int(item[2])
            except Exception:
                continue
            if not title or page < 1:
                continue
            normalized.append((page, lvl, idx, [lvl, title, page]))
        normalized.sort(key=lambda x: (x[0], x[1], x[2]))
        return [x[3] for x in normalized]

    @staticmethod
    def _toc_to_dicts(toc_items):
        out = []
        for item in toc_items or []:
            if len(item) < 3:
                continue
            try:
                out.append({
                    "level": int(item[0]),
                    "title": str(item[1]),
                    "page": int(item[2]),
                })
            except Exception:
                continue
        return out

    @staticmethod
    def _parse_import_toc(data, max_pages):
        entries = data
        if isinstance(data, dict):
            if isinstance(data.get("toc"), list):
                entries = data.get("toc")
            elif isinstance(data.get("bookmarks"), list):
                entries = data.get("bookmarks")
        if not isinstance(entries, list):
            return None, "JSON格式错误：需要数组或包含 toc/bookmarks 数组"

        out = []
        for idx, item in enumerate(entries, 1):
            try:
                if isinstance(item, dict):
                    lvl = int(item.get("level", 1))
                    title = str(item.get("title", "")).strip()
                    page = int(item.get("page", 1))
                elif isinstance(item, (list, tuple)) and len(item) >= 3:
                    lvl = int(item[0])
                    title = str(item[1]).strip()
                    page = int(item[2])
                else:
                    return None, f"第{idx}条书签格式错误"
            except Exception:
                return None, f"第{idx}条书签格式错误"

            if lvl < 1:
                return None, f"第{idx}条书签级别必须 >= 1"
            if not title:
                return None, f"第{idx}条书签标题为空"
            if page < 1 or page > max_pages:
                return None, f"第{idx}条书签页码超出范围（共{max_pages}页）"

            out.append([lvl, title, page])

        return out, None

    @staticmethod
    def _parse_levels(levels_text):
        text = (levels_text or "").strip()
        if not text:
            return set(), None
        text = text.replace("，", ",").replace("；", ",").replace(";", ",")
        levels = set()
        for part in text.split(","):
            part = part.strip()
            if not part:
                continue
            try:
                v = int(part)
            except Exception:
                return None, f"无效级别: {part}（示例: 1,2）"
            if v < 1:
                return None, f"级别必须 >= 1: {part}"
            levels.add(v)
        return levels, None

    def _auto_generate_toc(self, doc, pattern):
        default_pattern = (
            r"^(第[一二三四五六七八九十百千万0-9]+[编卷篇章节]|"
            r"\d+(?:\.\d+){0,3}\s+.+)"
        )
        pat = pattern.strip() if pattern else default_pattern
        try:
            reg = re.compile(pat)
        except re.error:
            reg = re.compile(default_pattern)

        seen = set()
        out = []
        total_pages = len(doc)
        for i in range(total_pages):
            page_num = i + 1
            page = doc[i]
            page_items = 0
            lines = self._collect_candidate_lines(page)
            for line in lines:
                text = self._normalize_heading_line(line)
                if not text:
                    continue
                if len(text) < 2 or len(text) > 80:
                    continue
                if not reg.search(text):
                    continue
                dedupe_key = (text, page_num)
                if dedupe_key in seen:
                    continue
                seen.add(dedupe_key)
                out.append([self._guess_level(text), text, page_num])
                page_items += 1
                if page_items >= 3:
                    break
            if i % 5 == 0 or i == total_pages - 1:
                pct = min(95, int((i + 1) / max(total_pages, 1) * 90))
                self._report(pct, f"自动分析页 {i + 1}/{total_pages}...", "正在识别标题")
        return self._sort_toc_by_page(out)

    @staticmethod
    def _collect_candidate_lines(page):
        try:
            blocks = page.get_text("blocks") or []
        except Exception:
            blocks = []
        lines = []
        for b in blocks:
            if len(b) < 5:
                continue
            text = str(b[4] or "").strip()
            if not text:
                continue
            for line in text.splitlines():
                s = line.strip()
                if s:
                    lines.append(s)
        if not lines:
            try:
                raw = page.get_text("text") or ""
                lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
            except Exception:
                lines = []
        return lines

    @staticmethod
    def _normalize_heading_line(line):
        text = re.sub(r"\s+", " ", str(line or "")).strip()
        text = text.strip("•·-—_")
        return text

    @staticmethod
    def _guess_level(text):
        m_num = re.match(r"^(\d+(?:\.\d+){0,3})\b", text)
        if m_num:
            return min(3, m_num.group(1).count(".") + 1)
        m_cn = re.match(r"^第[一二三四五六七八九十百千万0-9]+([编卷篇章节])", text)
        if m_cn:
            tail = m_cn.group(1)
            if tail in ("节",):
                return 2
            if tail in ("章",):
                return 1
            return 1
        return 1
