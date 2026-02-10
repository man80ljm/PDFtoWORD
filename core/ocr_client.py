"""
百度OCR API客户端 — 支持通用文字识别和公式识别。

本模块不依赖任何UI库，可独立使用和测试。
"""

import base64
import io
import logging
import time

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


# ============================================================
# 简单加解密（用于设置文件中的API Key存储）
# ============================================================

def simple_encrypt(text):
    """简单混淆存储（非安全加密，仅避免明文）"""
    if not text:
        return ""
    return base64.b64encode(text.encode('utf-8')).decode('utf-8')


def simple_decrypt(encoded):
    """解码简单混淆"""
    if not encoded:
        return ""
    try:
        return base64.b64decode(encoded.encode('utf-8')).decode('utf-8')
    except Exception:
        return encoded  # 兼容旧版明文


# ============================================================
# 百度OCR客户端
# ============================================================

class BaiduOCRClient:
    """百度OCR API客户端 - 支持通用文字识别和公式识别"""

    TOKEN_URL = "https://aip.baidubce.com/oauth/2.0/token"
    OCR_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic"
    FORMULA_URL = "https://aip.baidubce.com/rest/2.0/ocr/v1/formula"

    def __init__(self, api_key, secret_key):
        self.api_key = api_key
        self.secret_key = secret_key
        self._access_token = None
        self._token_time = 0

    def _get_access_token(self):
        """获取百度API access_token（有效期30天，自动缓存）"""
        if self._access_token and (time.time() - self._token_time) < 86400 * 25:
            return self._access_token
        params = {
            "grant_type": "client_credentials",
            "client_id": self.api_key,
            "client_secret": self.secret_key,
        }
        resp = requests.post(self.TOKEN_URL, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if "access_token" not in data:
            raise RuntimeError(f"百度API认证失败: {data.get('error_description', data)}")
        self._access_token = data["access_token"]
        self._token_time = time.time()
        return self._access_token

    def test_connection(self):
        """测试API连接是否可用。返回 (bool, str)。"""
        try:
            self._get_access_token()
            return True, "连接成功"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _compress_image(image_bytes, max_size_bytes=3 * 1024 * 1024):
        """将图片压缩为JPEG格式，确保不超过百度API的大小限制"""
        img = Image.open(io.BytesIO(image_bytes))
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        # 先尝试高质量JPEG
        for quality in [90, 80, 65, 50]:
            buf = io.BytesIO()
            img.save(buf, 'JPEG', quality=quality)
            jpg_bytes = buf.getvalue()
            b64_len = len(base64.b64encode(jpg_bytes))
            if b64_len <= max_size_bytes:
                logging.info(f'Image compressed: {len(image_bytes)//1024}KB→{len(jpg_bytes)//1024}KB '
                             f'(q={quality}, b64={b64_len//1024}KB)')
                return jpg_bytes
        # 如果还是太大，缩小尺寸
        w, h = img.size
        img = img.resize((w // 2, h // 2), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, 'JPEG', quality=70)
        return buf.getvalue()

    def recognize_text(self, image_bytes):
        """通用文字识别（高精度版），返回文字行列表"""
        token = self._get_access_token()
        compressed = self._compress_image(image_bytes)
        img_b64 = base64.b64encode(compressed).decode()
        logging.info(f'OCR text request: image base64 size = {len(img_b64)//1024} KB')
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "image": img_b64,
            "language_type": "CHN_ENG",
            "detect_direction": "true",
            "paragraph": "true",
        }
        resp = requests.post(
            f"{self.OCR_URL}?access_token={token}",
            headers=headers, data=data, timeout=60
        )
        resp.raise_for_status()
        result = resp.json()
        logging.info(f'OCR text response keys: {list(result.keys())}, '
                     f'words_num: {result.get("words_result_num", 0)}')
        if "error_code" in result:
            raise RuntimeError(f"OCR识别失败[{result.get('error_code')}]: "
                               f"{result.get('error_msg', result)}")
        words = []
        for item in result.get("words_result", []):
            words.append(item.get("words", ""))
        return words

    def recognize_formula(self, image_bytes):
        """公式识别，返回 LaTeX 字符串列表"""
        token = self._get_access_token()
        compressed = self._compress_image(image_bytes)
        img_b64 = base64.b64encode(compressed).decode()
        logging.info(f'Formula request: image base64 size = {len(img_b64)//1024} KB')
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "image": img_b64,
            "recognize_granularity": "big",
        }
        resp = requests.post(
            f"{self.FORMULA_URL}?access_token={token}",
            headers=headers, data=data, timeout=60
        )
        resp.raise_for_status()
        result = resp.json()
        logging.info(f'Formula response keys: {list(result.keys())}')
        if "error_code" in result:
            raise RuntimeError(f"公式识别失败[{result.get('error_code')}]: "
                               f"{result.get('error_msg', result)}")
        formulas = []
        # 百度API可能返回 words_result 或 formulas_result，两个都尝试
        formula_items = result.get("formulas_result", result.get("words_result", []))
        for item in formula_items:
            text = item.get("words", "")
            if text:
                formulas.append(text)
                logging.info(f'  Formula detected: {text[:80]}')
        if not formulas:
            logging.info(f'  No formulas found in response: {str(result)[:200]}')
        return formulas
