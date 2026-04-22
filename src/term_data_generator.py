"""
期趸数据自动生成器
负责根据缴费年期信息自动生成期趸标识数据
"""

import pandas as pd
import numpy as np
from typing import List, Dict, Optional, Tuple, Any
import logging
import os
import time

# 设置日志
logger = logging.getLogger(__name__)


class TermDataGenerator:
    """期趸数据生成器"""

    def __init__(self):
        self.generation_history = []  # 记录生成历史

    def generate_term_data_column(self, excel_path: str, source_column: str, target_column: str,
                                 term_prefix: str = "趸交", annual_prefix: str = "年交",
                                 sheet_name: Optional[str] = None,
                                 suffix: str = "") -> bool:
        """
        在Excel文件中生成期趸数据列

        Args:
            excel_path: Excel文件路径
            source_column: 源列名（缴费年期）
            target_column: 目标列名（期趸数据）
            term_prefix: 趸交前缀
            annual_prefix: 年交前缀
            sheet_name: 工作表名称
            suffix: 后缀（如FYP, SFYP等）

        Returns:
            bool: 是否生成成功
        """
        try:
            logger.info(f"开始生成期趸数据列: {source_column} -> {target_column}")

            # 验证源文件
            if not os.path.exists(excel_path):
                logger.error(f"Excel文件不存在: {excel_path}")
                return False

            # 读取Excel文件
            try:
                if sheet_name:
                    df = pd.read_excel(excel_path, sheet_name=sheet_name)
                else:
                    df = pd.read_excel(excel_path)
            except Exception as e:
                logger.error(f"读取Excel文件失败: {str(e)}")
                return False

            # 检查源列是否存在
            if source_column not in df.columns:
                logger.error(f"源列 '{source_column}' 不存在")
                return False

            # 检查目标列是否已存在
            if target_column in df.columns:
                logger.warning(f"目标列 '{target_column}' 已存在，将覆盖")

            # 生成期趸数据
            term_data = self._convert_payment_years_to_term_data(
                df[source_column],
                term_prefix=term_prefix,
                annual_prefix=annual_prefix,
                suffix=suffix
            )

            # 添加新列到DataFrame
            df[target_column] = term_data

            # 保存修改后的Excel文件
            self._save_excel_with_retry(df, excel_path, sheet_name)

            # 记录生成历史
            self._record_generation_history(
                excel_path, source_column, target_column,
                len(term_data), term_prefix, annual_prefix, suffix
            )

            logger.info(f"成功生成期趸数据列: {target_column}")
            return True

        except Exception as e:
            logger.error(f"生成期趸数据列失败: {str(e)}")
            return False

    def _convert_payment_years_to_term_data(self, payment_series: pd.Series,
                                          term_prefix: str = "趸交",
                                          annual_prefix: str = "年交",
                                          suffix: str = "") -> List[str]:
        """
        将缴费年期转换为期趸数据

        Args:
            payment_series: 缴费年期数据列
            term_prefix: 趸交前缀
            annual_prefix: 年交前缀
            suffix: 后缀

        Returns:
            List[str]: 期趸数据列表
        """
        term_data = []

        for value in payment_series:
            try:
                # 处理空值
                if pd.isna(value) or value == "" or value is None:
                    term_data.append("")
                    continue

                # 转换为数值
                if isinstance(value, str):
                    # 清理字符串中的非数字字符
                    cleaned_value = ''.join(c for c in value if c.isdigit() or c == '.')
                    if cleaned_value:
                        payment_years = float(cleaned_value)
                    else:
                        term_data.append("")
                        continue
                else:
                    payment_years = float(value)

                # 根据缴费年期生成期趸数据
                if payment_years == 0:
                    term_data.append(f"{term_prefix}{suffix}")
                elif payment_years == 1:
                    term_data.append(f"{term_prefix}{suffix}")
                else:
                    # 检查是否为整数
                    if payment_years.is_integer():
                        term_data.append(f"{int(payment_years)}{annual_prefix}{suffix}")
                    else:
                        term_data.append(f"{payment_years}{annual_prefix}{suffix}")

            except (ValueError, TypeError, AttributeError):
                # 处理无法转换的值
                term_data.append("")

        return term_data

    def _save_excel_with_retry(self, df: pd.DataFrame, excel_path: str,
                              sheet_name: Optional[str] = None, max_retries: int = 3):
        """
        带重试机制的Excel文件保存

        Args:
            df: 要保存的DataFrame
            excel_path: 文件路径
            sheet_name: 工作表名称
            max_retries: 最大重试次数
        """
        for attempt in range(max_retries):
            try:
                # 等待文件可用
                if not self._wait_for_file_available(excel_path, timeout=2):
                    raise OSError(f"文件被占用: {excel_path}")

                # 保存文件
                if sheet_name:
                    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    df.to_excel(excel_path, index=False)

                # 等待保存完成
                time.sleep(0.2)

                # 验证文件是否成功保存
                if self._wait_for_file_available(excel_path, timeout=2):
                    logger.debug(f"Excel文件保存成功: {excel_path}")
                    return
                else:
                    raise OSError("文件保存后无法访问")

            except Exception as e:
                if attempt < max_retries - 1:
                    logger.warning(f"保存Excel文件失败，重试中... ({attempt + 1}/{max_retries}): {str(e)}")
                    time.sleep(0.5)
                else:
                    raise e

    def _wait_for_file_available(self, file_path: str, timeout: float = 5) -> bool:
        """
        等待文件变为可访问状态

        Args:
            file_path: 文件路径
            timeout: 超时时间（秒）

        Returns:
            bool: 文件是否可用
        """
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # 尝试以只读模式打开文件
                with open(file_path, 'rb') as f:
                    pass
                return True
            except (PermissionError, OSError):
                time.sleep(0.1)
        return False

    def _record_generation_history(self, excel_path: str, source_column: str,
                                  target_column: str, row_count: int,
                                  term_prefix: str, annual_prefix: str, suffix: str):
        """记录生成历史"""
        history_record = {
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "excel_file": os.path.basename(excel_path),
            "source_column": source_column,
            "target_column": target_column,
            "row_count": row_count,
            "term_prefix": term_prefix,
            "annual_prefix": annual_prefix,
            "suffix": suffix
        }

        self.generation_history.append(history_record)

        # 只保留最近50条记录
        if len(self.generation_history) > 50:
            self.generation_history = self.generation_history[-50:]

    def preview_term_data_generation(self, excel_path: str, source_column: str,
                                   term_prefix: str = "趸交",
                                   annual_prefix: str = "年交",
                                   suffix: str = "",
                                   max_preview_rows: int = 10) -> Dict[str, Any]:
        """
        预览期趸数据生成结果

        Args:
            excel_path: Excel文件路径
            source_column: 源列名
            term_prefix: 趸交前缀
            annual_prefix: 年交前缀
            suffix: 后缀
            max_preview_rows: 最大预览行数

        Returns:
            Dict[str, Any]: 预览结果
        """
        try:
            if not os.path.exists(excel_path):
                return {"error": "文件不存在"}

            # 读取Excel文件（只读取前几行用于预览）
            df = pd.read_excel(excel_path, nrows=max_preview_rows)

            if source_column not in df.columns:
                return {"error": f"源列 '{source_column}' 不存在"}

            # 生成预览数据
            source_data = df[source_column].tolist()
            preview_data = self._convert_payment_years_to_term_data(
                df[source_column],
                term_prefix=term_prefix,
                annual_prefix=annual_prefix,
                suffix=suffix
            )

            # 分析数据分布
            distribution = self._analyze_term_data_distribution(source_data)

            return {
                "success": True,
                "source_column": source_column,
                "preview_rows": min(len(source_data), max_preview_rows),
                "total_rows": len(source_data),
                "source_data": source_data[:max_preview_rows],
                "preview_data": preview_data[:max_preview_rows],
                "distribution": distribution,
                "term_prefix": term_prefix,
                "annual_prefix": annual_prefix,
                "suffix": suffix
            }

        except Exception as e:
            return {"error": f"预览生成失败: {str(e)}"}

    def _analyze_term_data_distribution(self, payment_data: List[Any]) -> Dict[str, Any]:
        """分析缴费年期的分布情况"""
        try:
            distribution = {
                "total_count": len(payment_data),
                "valid_count": 0,
                "invalid_count": 0,
                "term_count": 0,  # 趸交数量
                "annual_counts": {},  # 年交数量统计
                "unique_values": set()
            }

            for value in payment_data:
                if pd.isna(value) or value == "" or value is None:
                    distribution["invalid_count"] += 1
                    continue

                try:
                    # 转换为数值
                    if isinstance(value, str):
                        cleaned_value = ''.join(c for c in value if c.isdigit() or c == '.')
                        if cleaned_value:
                            years = float(cleaned_value)
                        else:
                            distribution["invalid_count"] += 1
                            continue
                    else:
                        years = float(value)

                    distribution["valid_count"] += 1
                    distribution["unique_values"].add(years)

                    if years == 0 or years == 1:
                        distribution["term_count"] += 1
                    else:
                        years_key = int(years) if years.is_integer() else years
                        distribution["annual_counts"][years_key] = distribution["annual_counts"].get(years_key, 0) + 1

                except (ValueError, TypeError):
                    distribution["invalid_count"] += 1

            # 转换set为list以便JSON序列化
            distribution["unique_values"] = sorted(list(distribution["unique_values"]))

            return distribution

        except Exception as e:
            logger.error(f"分析数据分布失败: {str(e)}")
            return {"error": str(e)}

    def validate_source_data(self, excel_path: str, source_column: str,
                           sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """
        验证源数据是否适合生成期趸数据

        Args:
            excel_path: Excel文件路径
            source_column: 源列名
            sheet_name: 工作表名称

        Returns:
            Dict[str, Any]: 验证结果
        """
        try:
            if not os.path.exists(excel_path):
                return {"valid": False, "error": "文件不存在"}

            # 读取Excel文件
            if sheet_name:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(excel_path)

            if source_column not in df.columns:
                return {"valid": False, "error": f"源列 '{source_column}' 不存在"}

            # 分析源数据
            source_data = df[source_column].tolist()
            distribution = self._analyze_term_data_distribution(source_data)

            # 判断是否适合生成期趸数据
            valid_ratio = distribution["valid_count"] / distribution["total_count"] if distribution["total_count"] > 0 else 0

            result = {
                "valid": True,
                "source_column": source_column,
                "total_rows": distribution["total_count"],
                "valid_rows": distribution["valid_count"],
                "invalid_rows": distribution["invalid_count"],
                "valid_ratio": valid_ratio,
                "term_count": distribution["term_count"],
                "annual_variety": len(distribution["annual_counts"]),
                "recommendations": []
            }

            # 生成建议
            if valid_ratio < 0.8:
                result["recommendations"].append(f"有效数据比例较低({valid_ratio:.1%})，建议检查数据质量")

            if distribution["term_count"] == 0:
                result["recommendations"].append("数据中没有检测到趸交情况")

            if len(distribution["annual_counts"]) == 0:
                result["recommendations"].append("数据中没有检测到年交情况")

            if result["valid_ratio"] >= 0.8 and (distribution["term_count"] > 0 or len(distribution["annual_counts"]) > 0):
                result["recommendations"].append("数据适合生成期趸标识")

            return result

        except Exception as e:
            return {"valid": False, "error": f"验证失败: {str(e)}"}

    def get_generation_history(self, limit: int = 20) -> List[Dict[str, Any]]:
        """
        获取生成历史记录

        Args:
            limit: 返回记录数限制

        Returns:
            List[Dict]: 历史记录列表
        """
        return self.generation_history[-limit:] if limit > 0 else self.generation_history

    def clear_generation_history(self):
        """清空生成历史"""
        self.generation_history.clear()
        logger.info("已清空期趸数据生成历史")

    def get_supported_patterns(self) -> Dict[str, List[str]]:
        """
        获取支持的缴费年期模式

        Returns:
            Dict[str, List[str]]: 支持的模式
        """
        return {
            "term_payment": [0, 1],  # 趸交年数
            "annual_payment": [3, 5, 10, 15, 20, 25, 30],  # 常见年交年数
            "output_formats": {
                "default": {
                    "term": "趸交",
                    "annual": "{年数}年交"
                },
                "with_suffix": {
                    "term": "趸交FYP",
                    "annual": "{年数}年交SFYP"
                },
                "custom": {
                    "term_prefix": "趸交",
                    "annual_prefix": "年交",
                    "suffix": ""
                }
            }
        }