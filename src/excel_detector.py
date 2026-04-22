"""
Excel数据检测器
负责检测Excel文件的结构和列信息，为智能匹配提供数据支持
"""

import pandas as pd
import os
from typing import List, Dict, Optional, Set, Tuple, Any
import logging

# 设置日志
logger = logging.getLogger(__name__)


class ExcelColumnDetector:
    """Excel列检测器"""

    def __init__(self):
        self.column_cache = {}  # 列信息缓存

    def get_excel_columns(self, excel_path: str, sheet_name: Optional[str] = None) -> List[str]:
        """
        获取Excel文件的所有列名

        Args:
            excel_path: Excel文件路径
            sheet_name: 工作表名称，如果为None则使用第一个工作表

        Returns:
            List[str]: 列名列表
        """
        try:
            # 检查缓存
            cache_key = f"{excel_path}_{sheet_name or 'first'}"
            if cache_key in self.column_cache:
                logger.debug(f"从缓存获取Excel列信息: {excel_path}")
                return self.column_cache[cache_key]

            # 验证文件存在
            if not os.path.exists(excel_path):
                logger.error(f"Excel文件不存在: {excel_path}")
                return []

            # 读取Excel文件
            try:
                if sheet_name:
                    df = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=0)  # 只读取列名
                else:
                    df = pd.read_excel(excel_path, nrows=0)  # 只读取列名

                columns = df.columns.tolist()
                logger.info(f"检测到Excel列: {columns}")

                # 缓存结果
                self.column_cache[cache_key] = columns

                return columns

            except Exception as e:
                logger.error(f"读取Excel文件失败 {excel_path}: {str(e)}")
                return []

        except Exception as e:
            logger.error(f"获取Excel列名失败: {str(e)}")
            return []

    def get_sheet_names(self, excel_path: str) -> List[str]:
        """
        获取Excel文件的所有工作表名称

        Args:
            excel_path: Excel文件路径

        Returns:
            List[str]: 工作表名称列表
        """
        try:
            if not os.path.exists(excel_path):
                logger.error(f"Excel文件不存在: {excel_path}")
                return []

            with pd.ExcelFile(excel_path) as xls:
                sheet_names = xls.sheet_names
                logger.info(f"检测到工作表: {sheet_names}")
                return sheet_names

        except Exception as e:
            logger.error(f"获取工作表名称失败: {str(e)}")
            return []

    def has_required_columns(self, excel_path: str, required_columns: List[str],
                           sheet_name: Optional[str] = None) -> Dict[str, bool]:
        """
        检查Excel是否包含所需的列

        Args:
            excel_path: Excel文件路径
            required_columns: 需要检查的列名列表
            sheet_name: 工作表名称

        Returns:
            Dict[str, bool]: 列名 -> 是否存在的映射
        """
        try:
            available_columns = self.get_excel_columns(excel_path, sheet_name)

            result = {}
            for col in required_columns:
                result[col] = col in available_columns

            logger.info(f"列检查结果: {result}")
            return result

        except Exception as e:
            logger.error(f"检查必需列失败: {str(e)}")
            return {col: False for col in required_columns}

    def find_similar_columns(self, excel_path: str, target_columns: List[str],
                           similarity_threshold: float = 0.6,
                           sheet_name: Optional[str] = None) -> Dict[str, List[str]]:
        """
        查找相似的列名

        Args:
            excel_path: Excel文件路径
            target_columns: 目标列名列表
            similarity_threshold: 相似度阈值
            sheet_name: 工作表名称

        Returns:
            Dict[str, List[str]]: 目标列 -> 相似列列表的映射
        """
        try:
            available_columns = self.get_excel_columns(excel_path, sheet_name)
            result = {}

            for target in target_columns:
                similar_columns = []
                target_lower = target.lower()

                for available in available_columns:
                    available_lower = available.lower()

                    # 精确匹配
                    if target_lower == available_lower:
                        similar_columns.insert(0, available)  # 精确匹配排在最前面
                    # 包含关系
                    elif target_lower in available_lower or available_lower in target_lower:
                        similar_columns.append(available)
                    # 编辑距离相似度（简单实现）
                    elif self._calculate_similarity(target_lower, available_lower) >= similarity_threshold:
                        similar_columns.append(available)

                result[target] = similar_columns

            return result

        except Exception as e:
            logger.error(f"查找相似列失败: {str(e)}")
            return {col: [] for col in target_columns}

    def _calculate_similarity(self, str1: str, str2: str) -> float:
        """
        计算两个字符串的相似度（简单实现）

        Args:
            str1: 字符串1
            str2: 字符串2

        Returns:
            float: 相似度 (0-1)
        """
        try:
            # 使用简单的字符匹配相似度
            len1, len2 = len(str1), len(str2)
            if len1 == 0 or len2 == 0:
                return 0.0

            # 计算公共字符数量
            common_chars = set(str1) & set(str2)
            if not common_chars:
                return 0.0

            # 简单的相似度计算
            similarity = len(common_chars) / max(len1, len2)
            return similarity

        except Exception as e:
            logger.debug(f"字符串相似度计算失败：{e}")
            return 0.0

    def detect_term_data_columns(self, excel_path: str, sheet_name: Optional[str] = None) -> Dict[str, str]:
        """
        检测期趸数据相关的列

        Args:
            excel_path: Excel文件路径
            sheet_name: 工作表名称

        Returns:
            Dict[str, str]: 检测结果 {列类型: 列名}
        """
        try:
            columns = self.get_excel_columns(excel_path, sheet_name)
            result = {}

            # 定义常见的期趸数据列名模式
            term_data_patterns = {
                "payment_years": [
                    "缴费年期", "缴费年期（主险）", "缴费年期（主险） ID", "缴费年期ID",
                    "保费年期", "缴费期限", "保险年期", "年期", "缴费年数"
                ],
                "term_data": [
                    "期趸数据", "缴费方式", "期交趸交", "缴费类型", "年期标识"
                ]
            }

            # 检测各类列
            for col_type, patterns in term_data_patterns.items():
                found_columns = self._find_columns_by_patterns(columns, patterns)
                if found_columns:
                    result[col_type] = found_columns[0]  # 取第一个匹配的

            logger.info(f"期趸数据列检测结果: {result}")
            return result

        except Exception as e:
            logger.error(f"检测期趸数据列失败: {str(e)}")
            return {}

    def _find_columns_by_patterns(self, available_columns: List[str], patterns: List[str]) -> List[str]:
        """
        根据模式查找匹配的列

        Args:
            available_columns: 可用列名列表
            patterns: 模式列表

        Returns:
            List[str]: 匹配的列名列表（按优先级排序）
        """
        matches = []

        for pattern in patterns:
            pattern_lower = pattern.lower()

            for column in available_columns:
                column_lower = column.lower()

                # 精确匹配
                if pattern_lower == column_lower:
                    matches.insert(0, column)  # 精确匹配优先
                # 包含匹配
                elif pattern_lower in column_lower or column_lower in pattern_lower:
                    if column not in matches:
                        matches.append(column)

        return matches

    def analyze_excel_structure(self, excel_path: str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
        """
        分析Excel文件结构

        Args:
            excel_path: Excel文件路径
            sheet_name: 工作表名称

        Returns:
            Dict[str, Any]: 分析结果
        """
        try:
            if not os.path.exists(excel_path):
                return {"error": "文件不存在"}

            result = {
                "file_path": excel_path,
                "file_size": os.path.getsize(excel_path),
                "sheet_names": [],
                "current_sheet": sheet_name,
                "columns": [],
                "row_count": 0,
                "column_types": {},
                "term_data_info": {},
                "numeric_columns": [],
                "text_columns": []
            }

            # 获取工作表信息
            result["sheet_names"] = self.get_sheet_names(excel_path)

            # 分析指定工作表
            if sheet_name and sheet_name in result["sheet_names"]:
                target_sheet = sheet_name
            elif result["sheet_names"]:
                target_sheet = result["sheet_names"][0]  # 使用第一个工作表
            else:
                return {"error": "没有找到工作表"}

            result["current_sheet"] = target_sheet

            # 获取列信息
            columns = self.get_excel_columns(excel_path, target_sheet)
            result["columns"] = columns

            # 分析列类型和行数
            try:
                df = pd.read_excel(excel_path, sheet_name=target_sheet, nrows=100)  # 读取前100行分析

                result["row_count"] = len(df)

                for col in columns:
                    if col in df.columns:
                        # 检测列的数据类型
                        series = df[col].dropna()

                        if len(series) > 0:
                            # 检测是否为数字列
                            if pd.api.types.is_numeric_dtype(series):
                                result["numeric_columns"].append(col)
                                result["column_types"][col] = "numeric"
                            else:
                                result["text_columns"].append(col)
                                result["column_types"][col] = "text"

                        # 检测是否有空值
                        null_count = df[col].isnull().sum()
                        if null_count > 0:
                            result["column_types"][col] += f" (含{null_count}个空值)"

            except Exception as e:
                logger.warning(f"分析Excel数据类型失败: {str(e)}")

            # 检测期趸数据相关列
            result["term_data_info"] = self.detect_term_data_columns(excel_path, target_sheet)

            logger.info(f"Excel结构分析完成: {result['file_path']}")
            return result

        except Exception as e:
            error_msg = f"分析Excel结构失败: {str(e)}"
            logger.error(error_msg)
            return {"error": error_msg}

    def clear_cache(self):
        """清空缓存"""
        self.column_cache.clear()
        logger.debug("已清空Excel列检测缓存")

    def is_excel_file(self, file_path: str) -> bool:
        """
        检查文件是否为Excel文件

        Args:
            file_path: 文件路径

        Returns:
            bool: 是否为Excel文件
        """
        try:
            if not os.path.exists(file_path):
                return False

            # 检查文件扩展名
            excel_extensions = ['.xlsx', '.xls', '.xlsm']
            file_ext = os.path.splitext(file_path)[1].lower()

            return file_ext in excel_extensions

        except Exception as e:
            logger.error(f"检查Excel文件失败: {str(e)}")
            return False

    def get_excel_summary(self, excel_path: str, sheet_name: Optional[str] = None) -> str:
        """
        获取Excel文件的简要摘要信息

        Args:
            excel_path: Excel文件路径
            sheet_name: 工作表名称

        Returns:
            str: 摘要信息
        """
        try:
            structure = self.analyze_excel_structure(excel_path, sheet_name)

            if "error" in structure:
                return f"❌ 分析失败: {structure['error']}"

            summary_parts = []

            # 基本信息
            summary_parts.append(f"📁 文件: {os.path.basename(excel_path)}")
            summary_parts.append(f"📊 工作表: {structure['current_sheet']}")

            # 列信息
            col_count = len(structure['columns'])
            row_count = structure['row_count']
            summary_parts.append(f"📋 数据规模: {row_count}行 × {col_count}列")

            # 列类型信息
            numeric_count = len(structure['numeric_columns'])
            text_count = len(structure['text_columns'])
            summary_parts.append(f"🔢 数字列: {numeric_count}个")
            summary_parts.append(f"📝 文本列: {text_count}个")

            # 期趸数据信息
            term_info = structure['term_data_info']
            if term_info:
                term_parts = []
                if "payment_years" in term_info:
                    term_parts.append(f"缴费年期列: {term_info['payment_years']}")
                if "term_data" in term_info:
                    term_parts.append(f"期趸数据列: {term_info['term_data']}")

                if term_parts:
                    summary_parts.append("💰 期趸数据相关:")
                    for part in term_parts:
                        summary_parts.append(f"   • {part}")

            return "\n".join(summary_parts)

        except Exception as e:
            error_msg = f"生成Excel摘要失败: {str(e)}"
            logger.error(error_msg)
            return f"❌ {error_msg}"