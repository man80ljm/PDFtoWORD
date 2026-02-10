"""
PDF 加密/解密工具

加密：设置打开密码和/或权限密码
解密：移除已知密码的保护
通过 on_progress 回调报告进度，不直接操作UI。
"""

import logging
import os
from datetime import datetime

try:
    import fitz
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

# fitz 权限位常量
PERM_PRINT = fitz.PDF_PERM_PRINT           # 打印
PERM_MODIFY = fitz.PDF_PERM_MODIFY         # 修改内容
PERM_COPY = fitz.PDF_PERM_COPY             # 复制内容
PERM_ANNOTATE = fitz.PDF_PERM_ANNOTATE     # 添加注释
PERM_ALL = PERM_PRINT | PERM_MODIFY | PERM_COPY | PERM_ANNOTATE


class PDFEncryptConverter:
    """PDF加密/解密转换器，与 UI 完全解耦。

    用法::

        converter = PDFEncryptConverter(on_progress=my_callback)
        # 加密
        result = converter.encrypt(input_file, user_password="123")
        # 解密
        result = converter.decrypt(input_file, password="123")
    """

    def __init__(self, on_progress=None):
        self.on_progress = on_progress or (lambda *a: None)

    def _report(self, percent=-1, progress_text="", status_text=""):
        self.on_progress(percent, progress_text, status_text)

    def encrypt(self, input_file, output_path=None,
                user_password="", owner_password="",
                allow_print=True, allow_copy=True,
                allow_modify=False, allow_annotate=True):
        """加密PDF文件。

        Args:
            input_file: 输入PDF路径
            output_path: 输出路径, None则自动生成
            user_password: 打开密码（用户需输入此密码才能打开）
            owner_password: 权限密码（控制打印/复制等权限）
            allow_print: 允许打印
            allow_copy: 允许复制
            allow_modify: 允许修改
            allow_annotate: 允许添加注释

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

        if not user_password and not owner_password:
            result['message'] = "请至少设置一个密码（打开密码或权限密码）！"
            return result

        if not output_path:
            dir_path = os.path.dirname(input_file)
            basename = os.path.splitext(os.path.basename(input_file))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(dir_path, f"{basename}_加密_{timestamp}.pdf")

        try:
            self._report(percent=20, progress_text="正在打开PDF...",
                         status_text="准备加密...")

            doc = fitz.open(input_file)
            page_count = len(doc)

            if page_count == 0:
                doc.close()
                result['message'] = "PDF文件无内容"
                return result

            self._report(percent=50, progress_text="正在加密...",
                         status_text="设置密码和权限中...")

            # 计算权限
            perm = 0
            if allow_print:
                perm |= PERM_PRINT
            if allow_copy:
                perm |= PERM_COPY
            if allow_modify:
                perm |= PERM_MODIFY
            if allow_annotate:
                perm |= PERM_ANNOTATE

            # 如果只设打开密码不设权限密码，则默认权限密码 = 打开密码
            effective_owner = owner_password or user_password
            effective_user = user_password

            encrypt_meth = fitz.PDF_ENCRYPT_AES_256

            doc.save(
                output_path,
                encryption=encrypt_meth,
                user_pw=effective_user,
                owner_pw=effective_owner,
                permissions=perm,
            )
            doc.close()

            self._report(percent=100, progress_text="加密完成！")

            # 构建权限描述
            perms = []
            if allow_print:
                perms.append("打印")
            if allow_copy:
                perms.append("复制")
            if allow_modify:
                perms.append("修改")
            if allow_annotate:
                perms.append("注释")
            perm_text = "、".join(perms) if perms else "无"

            result['success'] = True
            result['output_file'] = output_path
            result['page_count'] = page_count
            result['message'] = (
                f"加密成功！共 {page_count} 页\n"
                f"打开密码: {'已设置' if user_password else '未设置'}\n"
                f"权限密码: {'已设置' if owner_password else '未设置'}\n"
                f"允许操作: {perm_text}"
            )

        except Exception as e:
            logging.error(f"PDF加密失败: {e}", exc_info=True)
            result['message'] = f"加密失败：{str(e)}"

        return result

    def decrypt(self, input_file, password="", output_path=None):
        """解密PDF文件（移除密码保护）。

        Args:
            input_file: 加密的PDF路径
            password: PDF的密码（打开密码或权限密码）
            output_path: 输出路径, None则自动生成

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

        if not output_path:
            dir_path = os.path.dirname(input_file)
            basename = os.path.splitext(os.path.basename(input_file))[0]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(dir_path, f"{basename}_解密_{timestamp}.pdf")

        try:
            self._report(percent=20, progress_text="正在打开PDF...",
                         status_text="尝试解密...")

            doc = fitz.open(input_file)

            # 检查是否加密
            if not doc.is_encrypted:
                doc.close()
                result['message'] = "此PDF未加密，无需解密"
                return result

            # 尝试用密码解锁
            if password:
                auth_ok = doc.authenticate(password)
                if not auth_ok:
                    doc.close()
                    result['message'] = "密码错误，无法解密！"
                    return result
            else:
                # 尝试空密码
                auth_ok = doc.authenticate("")
                if not auth_ok:
                    doc.close()
                    result['message'] = "此PDF需要密码才能解密，请输入密码"
                    return result

            page_count = len(doc)

            self._report(percent=60, progress_text="正在移除加密...",
                         status_text="保存解密文件...")

            # 保存为无加密版本
            doc.save(output_path)
            doc.close()

            self._report(percent=100, progress_text="解密完成！")

            result['success'] = True
            result['output_file'] = output_path
            result['page_count'] = page_count
            result['message'] = f"解密成功！共 {page_count} 页，已移除密码保护"

        except Exception as e:
            logging.error(f"PDF解密失败: {e}", exc_info=True)
            result['message'] = f"解密失败：{str(e)}"

        return result
