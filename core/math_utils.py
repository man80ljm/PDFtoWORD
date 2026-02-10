"""
数学字体检测、Unicode规范化、LaTeX→OMML公式转换工具。

本模块不依赖任何UI库，可独立使用和测试。
"""

import logging
import os

# ============================================================
# 可选依赖
# ============================================================
try:
    import latex2mathml.converter
    from lxml import etree
    LATEX2OMML_AVAILABLE = True
except ImportError:
    LATEX2OMML_AVAILABLE = False


# ============================================================
# 数学字体模式
# ============================================================

MATH_FONT_PATTERNS = [
    # Computer Modern (LaTeX经典)
    'CMMI', 'CMSY', 'CMEX', 'CMR', 'CMTI', 'CMBX', 'CMSS',
    # AMS 字体
    'MSAM', 'MSBM', 'EUSM', 'EUFM', 'EURM', 'EUBM',
    # Latin Modern (现代LaTeX默认)
    'LatinModernMath', 'LMMath', 'LatinModern-Math', 'LMRoman', 'LMSans',
    # STIX / XITS
    'STIX', 'XITS', 'STIXMath', 'XITSMath',
    # Cambria Math (Office常用)
    'CambriaMath', 'Cambria Math', 'Cambria-Math',
    # Libertinus / Linux Libertine
    'LibertinusMath', 'Libertinus Math', 'LinuxLibertine',
    # TeX Gyre
    'TeXGyre', 'TeX Gyre', 'TexGyreMath',
    # Fira Math
    'FiraMath', 'Fira Math',
    # Asana Math
    'AsanaMath', 'Asana-Math', 'Asana Math',
    # DejaVu Math
    'DejaVuMath', 'DejaVu Math',
    # Garamond Math
    'GaramondMath', 'Garamond-Math',
    # Symbol / MathType
    'Symbol', 'MT Extra', 'MT Symbol',
    'Mathematica', 'MathematicalPi',
    'Euclid',
    # 其他
    'RSFS', 'WASY', 'LASY',
]

MATH_FONT_KEYWORDS = ['math', 'symbol', 'cmmi', 'cmsy', 'cmex']


# ============================================================
# 字体和公式检测
# ============================================================

def detect_math_pages(fitz_doc, start=0, end=None):
    """检测包含数学公式的页面（通过分析字体、CID字体、Type3字体）

    Args:
        fitz_doc: PyMuPDF document对象
        start: 起始页索引
        end: 结束页索引（不含）

    Returns:
        set: 包含数学内容的页面索引集合
    """
    if end is None:
        end = len(fitz_doc)
    math_pages = set()
    for page_idx in range(start, end):
        page = fitz_doc[page_idx]
        fonts = page.get_fonts()
        has_math_font = False
        for font in fonts:
            font_type = font[2] if len(font) > 2 else ""
            font_basefont = font[3] if len(font) > 3 else ""
            clean_name = font_basefont
            if '+' in clean_name:
                clean_name = clean_name.split('+', 1)[1]
            clean_lower = clean_name.lower().replace('-', '').replace(' ', '')
            # 1. 精确匹配已知数学字体模式
            for pat in MATH_FONT_PATTERNS:
                if pat.lower().replace('-', '').replace(' ', '') in clean_lower:
                    has_math_font = True
                    break
            if has_math_font:
                break
            # 2. 关键词匹配
            for kw in MATH_FONT_KEYWORDS:
                if kw in clean_lower:
                    has_math_font = True
                    break
            if has_math_font:
                break
            # 3. Type3 字体常用于嵌入的数学符号
            if font_type == 'Type3':
                has_math_font = True
                break
        if has_math_font:
            math_pages.add(page_idx)
    return math_pages


def is_math_font(font_name):
    """判断字体名是否为数学字体"""
    if not font_name:
        return False
    clean = font_name
    if '+' in clean:
        clean = clean.split('+', 1)[1]
    clean_lower = clean.lower().replace('-', '').replace(' ', '')
    for pat in MATH_FONT_PATTERNS:
        if pat.lower().replace('-', '').replace(' ', '') in clean_lower:
            return True
    for kw in MATH_FONT_KEYWORDS:
        if kw in clean_lower:
            return True
    return False


def has_math_unicode(text):
    """检查文本是否包含需要规范化的数学Unicode字符"""
    for c in text:
        cp = ord(c)
        if 0x1D400 <= cp <= 0x1D7FF:  # Mathematical Alphanumeric Symbols
            return True
        if cp == 0x210E:  # PLANCK CONSTANT
            return True
    return False


def is_display_equation(block):
    """判断一个块是否为独立的行间公式（大部分为数学字体，不含CJK字符）"""
    if block.get("type") != 0:
        return False
    total_chars = 0
    math_chars = 0
    cjk_chars = 0
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            text = span.get("text", "").strip()
            font = span.get("font", "")
            is_math = is_math_font(font)
            for c in text:
                if c.isspace():
                    continue
                total_chars += 1
                if is_math:
                    math_chars += 1
                if 0x4E00 <= ord(c) <= 0x9FFF:
                    cjk_chars += 1
    if total_chars < 2:
        return False
    return math_chars / total_chars > 0.5 and cjk_chars == 0


def get_block_text(block):
    """提取块中所有span的文本"""
    parts = []
    for line in block.get("lines", []):
        line_parts = []
        for span in line.get("spans", []):
            line_parts.append(span.get("text", ""))
        parts.append("".join(line_parts))
    return " ".join(parts).strip()


# ============================================================
# Unicode数学字符规范化
# ============================================================

def normalize_math_unicode(text):
    """将 Unicode 数学字母数字符号转为普通字符，使 Word 能正确显示。
    例如: 𝑓(U+1D453) → f, 𝑥(U+1D465) → x, 𝜋(U+1D70B) → π"""
    if not text:
        return text
    result = []
    for c in text:
        cp = ord(c)
        mapped = _map_math_char(cp)
        result.append(mapped)
    return ''.join(result)


# --- 数学字符映射常量（模块级，避免每次调用重建） ---
_GREEK_LOWER = 'αβγδεζηθικλμνξοπρςστυφχψω'
_GREEK_UPPER = 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡϴΣΤΥΦΧΨΩ'
_MATH_OPERATORS = {
    0x2212: '-',   # MINUS SIGN → -
    0x2032: "'",   # PRIME → '
    0x2033: "''",  # DOUBLE PRIME
    0x2190: '←', 0x2192: '→', 0x21D2: '⇒', 0x21D0: '⇐',
    0x2260: '≠', 0x2264: '≤', 0x2265: '≥',
    0x222B: '∫', 0x2211: '∑', 0x220F: '∏',
    0x221A: '√', 0x221E: '∞', 0x2202: '∂',
    0x210E: 'h',  # PLANCK CONSTANT → h
}


def _map_math_char(cp):
    """将数学Unicode码点映射为普通可显示字符"""
    # Mathematical Italic Small (U+1D44E - U+1D467) → a-z
    if 0x1D44E <= cp <= 0x1D467:
        return chr(ord('a') + cp - 0x1D44E)
    # Mathematical Italic Capital (U+1D434 - U+1D44D) → A-Z
    if 0x1D434 <= cp <= 0x1D44D:
        return chr(ord('A') + cp - 0x1D434)
    # Mathematical Bold Small (U+1D41A - U+1D433) → a-z
    if 0x1D41A <= cp <= 0x1D433:
        return chr(ord('a') + cp - 0x1D41A)
    # Mathematical Bold Capital (U+1D400 - U+1D419) → A-Z
    if 0x1D400 <= cp <= 0x1D419:
        return chr(ord('A') + cp - 0x1D400)
    # Mathematical Bold Italic Small (U+1D482 - U+1D49B) → a-z
    if 0x1D482 <= cp <= 0x1D49B:
        return chr(ord('a') + cp - 0x1D482)
    # Mathematical Bold Italic Capital (U+1D468 - U+1D481) → A-Z
    if 0x1D468 <= cp <= 0x1D481:
        return chr(ord('A') + cp - 0x1D468)
    # Mathematical Sans-Serif variants
    if 0x1D5A0 <= cp <= 0x1D5B9:  # sans capital
        return chr(ord('A') + cp - 0x1D5A0)
    if 0x1D5BA <= cp <= 0x1D5D3:  # sans small
        return chr(ord('a') + cp - 0x1D5BA)
    # Mathematical Italic Greek Small (U+1D6FC - U+1D714) → α-ω
    if 0x1D6FC <= cp <= 0x1D714:
        idx = cp - 0x1D6FC
        if idx < len(_GREEK_LOWER):
            return _GREEK_LOWER[idx]
    # Mathematical Italic Greek Capital (U+1D6E2 - U+1D6FA) → Α-Ω
    if 0x1D6E2 <= cp <= 0x1D6FA:
        idx = cp - 0x1D6E2
        if idx < len(_GREEK_UPPER):
            return _GREEK_UPPER[idx]
    # Mathematical Bold Greek Small (U+1D736 - U+1D74E)
    if 0x1D736 <= cp <= 0x1D74E:
        idx = cp - 0x1D736
        if idx < len(_GREEK_LOWER):
            return _GREEK_LOWER[idx]
    # Mathematical Bold Greek Capital (U+1D71C - U+1D734)
    if 0x1D71C <= cp <= 0x1D734:
        idx = cp - 0x1D71C
        if idx < len(_GREEK_UPPER):
            return _GREEK_UPPER[idx]
    # 数学运算符映射
    if cp in _MATH_OPERATORS:
        return _MATH_OPERATORS[cp]
    return chr(cp)


# ============================================================
# LaTeX → OMML 公式转换
# ============================================================

def latex_to_omml(latex_str, xslt_path=None):
    """将LaTeX公式转为Word OMML XML元素。

    需要 latex2mathml 和 lxml，以及 MML2OMML.XSL（Office自带或自动检测）。

    Args:
        latex_str: LaTeX公式字符串
        xslt_path: MML2OMML.XSL文件路径，None则自动查找

    Returns:
        lxml Element 或 None
    """
    if not LATEX2OMML_AVAILABLE:
        return None

    # 清理LaTeX（去掉可能的$包裹）
    latex_clean = latex_str.strip()
    for prefix in ['$$', '$', '\\[', '\\(']:
        if latex_clean.startswith(prefix):
            latex_clean = latex_clean[len(prefix):]
    for suffix in ['$$', '$', '\\]', '\\)']:
        if latex_clean.endswith(suffix):
            latex_clean = latex_clean[:-len(suffix)]
    latex_clean = latex_clean.strip()
    if not latex_clean:
        return None

    # LaTeX → MathML
    try:
        mathml_str = latex2mathml.converter.convert(latex_clean)
    except Exception as e:
        logging.warning(f"LaTeX→MathML转换失败: {e}, 原始: {latex_clean}")
        return None

    # MathML → OMML via XSLT
    if xslt_path is None:
        candidate_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
            r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
            r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
            r"C:\Program Files\Microsoft Office\root\Office15\MML2OMML.XSL",
            r"C:\Program Files (x86)\Microsoft Office\root\Office15\MML2OMML.XSL",
        ]
        for p in candidate_paths:
            if os.path.exists(p):
                xslt_path = p
                break

    if xslt_path is None or not os.path.exists(xslt_path):
        logging.warning("未找到 MML2OMML.XSL，无法将MathML转为OMML")
        return None

    try:
        with open(xslt_path, 'rb') as f:
            xslt_doc = etree.parse(f)
        transform = etree.XSLT(xslt_doc)
        mathml_doc = etree.fromstring(mathml_str.encode())
        omml_result = transform(mathml_doc)
        omml_element = omml_result.getroot()
        return omml_element
    except Exception as e:
        logging.warning(f"MathML→OMML转换失败: {e}")
        return None


def insert_omml_to_paragraph(paragraph, omml_element):
    """将OMML公式元素插入到Word段落中"""
    from lxml import etree
    omml_str = etree.tostring(omml_element)
    omml_parsed = etree.fromstring(omml_str)
    paragraph._element.append(omml_parsed)
