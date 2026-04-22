"""
内存优化管理器

确保内存使用效率和防止内存泄漏。
"""

import gc
import os
import psutil
import pandas as pd
from typing import Dict, Any, List, Optional
from collections import OrderedDict
import time


class MemoryOptimizer:
    """内存优化管理器 - 监控和优化内存使用"""

    def __init__(self, max_memory_usage: float = 0.8, cache_size_limit: int = 100):
        """
        初始化内存优化管理器

        Args:
            max_memory_usage: 最大内存使用比例（0.1-1.0）
            cache_size_limit: 缓存项数量限制
        """
        self.max_memory_usage = max_memory_usage
        self.cache_size_limit = cache_size_limit
        self.process = psutil.Process(os.getpid())
        self.memory_history = OrderedDict()
        self.last_cleanup_time = time.time()
        self.cleanup_interval = 300  # 5分钟清理一次

    def get_memory_stats(self) -> Dict[str, Any]:
        """
        获取内存使用统计

        Returns:
            内存使用信息字典
        """
        memory_info = self.process.memory_info()
        memory_percent = self.process.memory_percent()

        stats = {
            'current_memory_mb': memory_info.rss / 1024 / 1024,
            'memory_percent': memory_percent,
            'max_memory_usage': self.max_memory_usage * 100,
            'is_near_limit': memory_percent > (self.max_memory_usage * 100 * 0.9),
            'cache_size': len(self.memory_history),
            'last_cleanup': self.last_cleanup_time
        }

        return stats

    def check_memory_usage(self) -> bool:
        """
        检查内存使用情况，必要时执行清理

        Returns:
            True如果内存使用正常，False如果需要清理
        """
        current_time = time.time()
        memory_percent = self.process.memory_percent()

        # 记录内存使用历史
        self.memory_history[current_time] = memory_percent

        # 保持历史记录不超过1000条
        if len(self.memory_history) > 1000:
            self.memory_history.popitem(last=False)

        # 检查是否需要清理
        need_cleanup = False
        if memory_percent > (self.max_memory_usage * 100):
            need_cleanup = True
        elif (current_time - self.last_cleanup_time) > self.cleanup_interval:
            need_cleanup = True

        if need_cleanup:
            self.perform_cleanup()

        return not need_cleanup

    def perform_cleanup(self) -> Dict[str, Any]:
        """
        执行内存清理

        Returns:
            清理结果统计
        """
        cleanup_stats = {
            'garbage_collected': 0,
            'memory_before': self.process.memory_info().rss / 1024 / 1024,
            'memory_after': 0,
            'memory_freed': 0
        }

        # 强制垃圾回收
        collected = gc.collect()
        cleanup_stats['garbage_collected'] = collected

        # 清理旧的内存历史记录
        current_time = time.time()
        keys_to_remove = []
        for timestamp in self.memory_history.keys():
            if current_time - timestamp > 3600:  # 保留最近1小时的历史
                keys_to_remove.append(timestamp)

        for key in keys_to_remove:
            del self.memory_history[key]

        # 获取清理后的内存使用情况
        memory_after = self.process.memory_info().rss / 1024 / 1024
        cleanup_stats['memory_after'] = memory_after
        cleanup_stats['memory_freed'] = cleanup_stats['memory_before'] - memory_after

        self.last_cleanup_time = current_time

        return cleanup_stats

    def suggest_dataframe_optimization(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        为DataFrame提供内存优化建议

        Args:
            df: 要分析的DataFrame

        Returns:
            优化建议字典
        """
        suggestions = []
        memory_usage = df.memory_usage(deep=True).sum() / 1024 / 1024  # MB

        # 检查数据类型优化
        for col in df.columns:
            series = df[col]
            col_memory = series.memory_usage(deep=True) / 1024 / 1024  # MB

            # 整数类型优化
            if pd.api.types.is_integer_dtype(series):
                min_val, max_val = series.min(), series.max()
                if min_val >= 0:
                    if max_val < 255:
                        suggestions.append(f"列 '{col}' 可转换为 uint8")
                    elif max_val < 65535:
                        suggestions.append(f"列 '{col}' 可转换为 uint16")
                    elif max_val < 4294967295:
                        suggestions.append(f"列 '{col}' 可转换为 uint32")
                else:
                    if min_val > -128 and max_val < 127:
                        suggestions.append(f"列 '{col}' 可转换为 int8")
                    elif min_val > -32768 and max_val < 32767:
                        suggestions.append(f"列 '{col}' 可转换为 int16")
                    elif min_val > -2147483648 and max_val < 2147483647:
                        suggestions.append(f"列 '{col}' 可转换为 int32")

            # 浮点数类型优化
            elif pd.api.types.is_float_dtype(series):
                suggestions.append(f"列 '{col}' 可考虑转换为 float32")

            # 字符串类型优化
            elif pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series):
                unique_ratio = series.nunique() / len(series)
                if unique_ratio < 0.5:  # 如果重复值较多，建议转换为category
                    suggestions.append(f"列 '{col}' 可转换为 category 类型")

        return {
            'total_memory_mb': memory_usage,
            'suggestions': suggestions,
            'optimization_potential': len(suggestions) > 0
        }

    def optimize_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        优化DataFrame的内存使用

        Args:
            df: 要优化的DataFrame

        Returns:
            优化后的DataFrame
        """
        optimized_df = df.copy()

        for col in optimized_df.columns:
            series = optimized_df[col]
            original_dtype = series.dtype

            # 跳过已经是优化过的类型
            if pd.api.types.is_categorical_dtype(series):
                continue

            # 整数类型优化
            if pd.api.types.is_integer_dtype(series):
                min_val, max_val = series.min(), series.max()
                if min_val >= 0:
                    if max_val < 255:
                        optimized_df[col] = series.astype('uint8')
                    elif max_val < 65535:
                        optimized_df[col] = series.astype('uint16')
                    elif max_val < 4294967295:
                        optimized_df[col] = series.astype('uint32')
                else:
                    if min_val > -128 and max_val < 127:
                        optimized_df[col] = series.astype('int8')
                    elif min_val > -32768 and max_val < 32767:
                        optimized_df[col] = series.astype('int16')
                    elif min_val > -2147483648 and max_val < 2147483647:
                        optimized_df[col] = series.astype('int32')

            # 浮点数类型优化
            elif pd.api.types.is_float_dtype(series):
                optimized_df[col] = series.astype('float32')

            # 字符串类型优化
            elif pd.api.types.is_string_dtype(series) or pd.api.types.is_object_dtype(series):
                unique_ratio = series.nunique() / len(series)
                if unique_ratio < 0.5 and series.nunique() < 1000:
                    try:
                        optimized_df[col] = series.astype('category')
                    except:
                        pass

        return optimized_df

    def get_memory_trend(self, minutes: int = 60) -> Dict[str, Any]:
        """
        获取内存使用趋势

        Args:
            minutes: 要分析的分钟数

        Returns:
            内存使用趋势信息
        """
        current_time = time.time()
        cutoff_time = current_time - (minutes * 60)

        # 筛选指定时间范围内的数据
        recent_history = {
            timestamp: memory_percent
            for timestamp, memory_percent in self.memory_history.items()
            if timestamp >= cutoff_time
        }

        if not recent_history:
            return {
                'trend': 'insufficient_data',
                'avg_memory': 0,
                'max_memory': 0,
                'min_memory': 0,
                'data_points': 0
            }

        memory_values = list(recent_history.values())
        avg_memory = sum(memory_values) / len(memory_values)
        max_memory = max(memory_values)
        min_memory = min(memory_values)

        # 判断趋势
        if len(memory_values) >= 2:
            recent_avg = sum(memory_values[-5:]) / min(5, len(memory_values))
            early_avg = sum(memory_values[:5]) / min(5, len(memory_values))

            if recent_avg > early_avg + 5:
                trend = 'increasing'
            elif recent_avg < early_avg - 5:
                trend = 'decreasing'
            else:
                trend = 'stable'
        else:
            trend = 'insufficient_data'

        return {
            'trend': trend,
            'avg_memory': avg_memory,
            'max_memory': max_memory,
            'min_memory': min_memory,
            'data_points': len(memory_values)
        }

    def set_cleanup_interval(self, seconds: int):
        """
        设置清理间隔

        Args:
            seconds: 清理间隔秒数
        """
        self.cleanup_interval = max(60, seconds)  # 最少1分钟

    def set_max_memory_usage(self, ratio: float):
        """
        设置最大内存使用比例

        Args:
            ratio: 内存使用比例（0.1-1.0）
        """
        self.max_memory_usage = max(0.1, min(1.0, ratio))