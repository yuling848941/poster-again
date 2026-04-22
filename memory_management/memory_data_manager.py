"""
内存数据管理器

提供统一的内存数据管理功能。
"""

import pandas as pd
import os
import hashlib
from typing import Dict, Optional, Any
try:
    from .data_formatter import DataFormatter
    from .memory_optimizer import MemoryOptimizer
except ImportError:
    from data_formatter import DataFormatter
    from memory_optimizer import MemoryOptimizer


class MemoryDataManager:
    """内存数据管理器 - 所有数据在内存中的统一管理和处理"""

    def __init__(self):
        """初始化内存数据管理器"""
        self.data_cache: Dict[str, pd.DataFrame] = {}  # 原始数据缓存
        self.formatted_cache: Dict[str, pd.DataFrame] = {}  # 格式化数据缓存
        self.formatter = DataFormatter()  # 数据格式化工具
        self.optimizer = MemoryOptimizer()  # 内存优化管理器

    def _get_file_hash(self, file_path: str) -> str:
        """生成文件的唯一哈希值用于缓存键"""
        try:
            # 使用文件路径和修改时间生成哈希
            stat = os.stat(file_path)
            content = f"{file_path}_{stat.st_mtime}_{stat.st_size}"
            return hashlib.md5(content.encode()).hexdigest()[:16]
        except (OSError, IOError):
            # 如果无法获取文件状态，使用文件路径的哈希
            return hashlib.md5(file_path.encode()).hexdigest()[:16]

    def load_excel_data(self, file_path: str, sheet_name: str = None) -> pd.DataFrame:
        """
        直接从Excel文件读取数据到内存

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称，None表示第一个工作表

        Returns:
            包含Excel数据的DataFrame

        Raises:
            FileNotFoundError: 文件不存在
            ValueError: 文件格式错误
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel文件不存在: {file_path}")

        # 检查缓存
        cache_key = f"{self._get_file_hash(file_path)}_{sheet_name or 'default'}"
        if cache_key in self.data_cache:
            return self.data_cache[cache_key].copy()

        try:
            # 直接读取Excel文件到内存
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)

            # 缓存数据
            self.data_cache[cache_key] = df.copy()

            # 检查内存使用情况
            self.optimizer.check_memory_usage()

            return df

        except Exception as e:
            raise ValueError(f"读取Excel文件失败: {file_path}, 错误: {str(e)}")

    def get_formatted_data(self, file_path: str, format_type: str, **kwargs) -> pd.DataFrame:
        """
        获取指定格式的数据，在内存中完成格式化

        Args:
            file_path: Excel文件路径
            format_type: 格式化类型 ('thousands_separator', 'term_data')
            **kwargs: 格式化参数

        Returns:
            格式化后的DataFrame
        """
        # 生成缓存键
        cache_key = f"{self._get_file_hash(file_path)}_{format_type}_{str(sorted(kwargs.items()))}"

        # 检查格式化缓存
        if cache_key in self.formatted_cache:
            return self.formatted_cache[cache_key].copy()

        # 先加载原始数据
        base_data = self.load_excel_data(file_path)

        # 在内存中进行格式化
        if format_type == 'thousands_separator':
            formatted_data = self.formatter.add_thousands_separator(base_data, **kwargs)
        elif format_type == 'term_data':
            formatted_data = self.formatter.calculate_term_data(base_data, **kwargs)
        else:
            raise ValueError(f"不支持的格式化类型: {format_type}")

        # 缓存格式化结果
        self.formatted_cache[cache_key] = formatted_data.copy()

        return formatted_data

    def calculate_term_data(self, base_data: pd.DataFrame, years: int = None,
                          term_column: str = None, **kwargs) -> pd.DataFrame:
        """
        在内存中计算期趸数据

        Args:
            base_data: 基础数据
            years: 缴费年期
            term_column: 期趸标识列名
            **kwargs: 其他参数

        Returns:
            包含期趸数据的DataFrame
        """
        return self.formatter.calculate_term_data(
            base_data, years, term_column, **kwargs
        )

    def get_formatted_data_with_term_and_separator(self, file_path: str,
                                                 use_thousand_separator: bool = True,
                                                 add_term_data: bool = True) -> pd.DataFrame:
        """
        获取包含千位分隔符和期趸数据的完整格式化数据（替代原有临时文件功能）

        Args:
            file_path: Excel文件路径
            use_thousand_separator: 是否添加千位分隔符
            add_term_data: 是否添加期趸数据

        Returns:
            完整格式化后的DataFrame
        """
        # 先加载原始数据
        result_df = self.load_excel_data(file_path)

        # 添加期趸数据（如果需要且存在源列）
        if add_term_data and "缴费年期（主险） ID" in result_df.columns:
            result_df = self.formatter.calculate_term_data(
                result_df,
                source_column="缴费年期（主险） ID",
                target_column="期趸数据"
            )

        # 添加千位分隔符（如果需要）
        if use_thousand_separator:
            result_df = self.formatter.add_thousands_separator(result_df)

        return result_df

    def clear_cache(self, file_path: str = None):
        """
        清理内存缓存

        Args:
            file_path: 特定文件的缓存，None表示清理所有缓存
        """
        if file_path is None:
            # 清理所有缓存
            self.data_cache.clear()
            self.formatted_cache.clear()
        else:
            # 清理特定文件的缓存
            file_hash = self._get_file_hash(file_path)
            keys_to_remove = [key for key in self.data_cache.keys()
                           if key.startswith(file_hash)]
            for key in keys_to_remove:
                del self.data_cache[key]

            keys_to_remove = [key for key in self.formatted_cache.keys()
                           if key.startswith(file_hash)]
            for key in keys_to_remove:
                del self.formatted_cache[key]

    def get_memory_usage(self) -> Dict[str, Any]:
        """
        获取内存使用统计

        Returns:
            内存使用信息字典
        """
        return self.optimizer.get_memory_stats()

    def invalidate_cache(self, file_path: str):
        """
        使特定文件的缓存失效（当文件被修改时调用）

        Args:
            file_path: 文件路径
        """
        self.clear_cache(file_path)

    def __del__(self):
        """析构函数，清理资源"""
        try:
            self.clear_cache()
        except:
            pass