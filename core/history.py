"""
转换历史记录管理

保存在 conversion_history.json 中，最多保留100条。
"""

import json
import logging
import os
from datetime import datetime

from core import get_app_dir


class ConversionHistory:
    """管理转换历史记录"""

    MAX_RECORDS = 100

    def __init__(self):
        self.history_file = os.path.join(get_app_dir(), "conversion_history.json")
        self._records = []
        self.load()

    def load(self):
        """从文件加载历史记录"""
        if not os.path.exists(self.history_file):
            return
        try:
            with open(self.history_file, 'r', encoding='utf-8') as f:
                self._records = json.load(f)
        except Exception:
            self._records = []

    def save(self):
        """保存历史记录到文件"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self._records, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"保存历史记录失败: {e}")

    def add(self, record):
        """添加一条记录

        Args:
            record: dict with keys:
                function (str): 功能名称
                input_files (list[str]): 输入文件
                output (str): 输出文件/目录
                success (bool): 是否成功
                message (str): 结果消息
                page_count (int): 处理页数
                timestamp (str): 时间戳（可选，自动填充）
        """
        if 'timestamp' not in record:
            record['timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._records.insert(0, record)
        # 限制最大记录数
        if len(self._records) > self.MAX_RECORDS:
            self._records = self._records[:self.MAX_RECORDS]
        self.save()

    def get_all(self):
        """获取所有记录"""
        return list(self._records)

    def clear(self):
        """清空历史记录"""
        self._records = []
        self.save()

    @property
    def count(self):
        return len(self._records)
