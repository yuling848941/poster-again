"""
数据格式化工具

在内存中完成所有数据格式化操作。
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional
import logging

logger = logging.getLogger(__name__)


class DataFormatter:
    """数据格式化工具 - 在内存中完成所有数据格式化操作"""

    def __init__(self):
        """初始化数据格式化工具"""
        self.term_mapping = {
            1: "趸交",
            10: "10年交",
            15: "15年交",
            20: "20年交",
            25: "25年交",
            30: "30年交"
        }

    def add_thousands_separator(self, df: pd.DataFrame,
                              columns: List[str] = None,
                              precision: int = 2) -> pd.DataFrame:
        """
        为指定列添加千位分隔符格式

        Args:
            df: 原始DataFrame
            columns: 需要格式化的列名列表，None表示自动检测数值列
            precision: 小数位数

        Returns:
            格式化后的DataFrame
        """
        result_df = df.copy()

        if columns is None:
            # 自动检测数值类型的列
            numeric_columns = result_df.select_dtypes(include=[np.number]).columns.tolist()

            # ✅ 新增：排除不应该添加千位分隔符的列（ID、编号等）
            excluded_columns = [
                '保单号', 'policy_number', 'PolicyNumber', 'POLICY_NUMBER',
                '客户号', '客户编号', '保单号(主险)', '投保单号',
                '编号', 'ID', 'id', 'Id',
                '序号', '行号', '索引'
            ]

            # 从数值列中排除这些列
            numeric_columns = [col for col in numeric_columns if col not in excluded_columns]

            logger.debug(f"自动检测到 {len(numeric_columns)} 个需要添加千位分隔符的数值列")
            logger.debug(f"已排除 {len(result_df.select_dtypes(include=[np.number]).columns.tolist()) - len(numeric_columns)} 个列：{excluded_columns}")

        else:
            numeric_columns = []

            # 验证指定的列是否存在
            for col in columns:
                if col in result_df.columns:
                    if pd.api.types.is_numeric_dtype(result_df[col]):
                        numeric_columns.append(col)
                    else:
                        # 尝试转换为数值
                        try:
                            result_df[col] = pd.to_numeric(result_df[col], errors='coerce')
                            numeric_columns.append(col)
                        except:
                            continue

        # 为数值列添加千位分隔符
        for col in numeric_columns:
            result_df[col] = result_df[col].apply(
                lambda x: self._format_number(x, precision) if pd.notna(x) else x
            )

        return result_df

    def calculate_term_data(self, base_data: pd.DataFrame, years: int = None,
                          term_column: str = None,
                          source_column: str = "缴费年期（主险） ID",
                          target_column: str = "期趸数据",
                          preserve_original: bool = True) -> pd.DataFrame:
        """
        在内存中计算期趸数据

        Args:
            base_data: 基础数据
            years: 缴费年期（如果提供，将使用此值而非从source_column读取）
            term_column: 期趸标识列名（已弃用，使用target_column）
            source_column: 源数据列名，默认为"缴费年期（主险） ID"
            target_column: 目标期趸数据列名，默认为"期趸数据"
            preserve_original: 是否保留原始列

        Returns:
            包含期趸数据的DataFrame
        """
        result_df = base_data.copy()

        # 使用新的参数名，保持向后兼容
        if term_column is not None:
            target_column = term_column

        # 创建期趸数据列
        if source_column in result_df.columns:
            term_data = []
            for value in result_df[source_column]:
                if pd.isna(value):
                    term_data.append("")
                elif years is not None:
                    # 使用提供的年数
                    if years == 1:
                        term_data.append("趸交FYP")
                    else:
                        term_data.append(f"{int(years)}年交SFYP")
                else:
                    # 从数据中读取年数
                    if value == 1:
                        term_data.append("趸交FYP")
                    else:
                        term_data.append(f"{int(value)}年交SFYP")

            # 将期趸数据列添加到DataFrame
            result_df[target_column] = term_data
        else:
            # 如果没有找到源列，使用传统的期趸标识逻辑
            if years is None:
                # 如果没有提供年数，无法处理
                raise ValueError(f"未找到源列 '{source_column}' 且未提供年数参数")

            # 确定期趸标识列名
            if target_column not in result_df.columns:
                target_column = '期趸标识'

            # 获取期趸标识
            term_text = self.term_mapping.get(years, f"{years}年交")

            if preserve_original and target_column in result_df.columns:
                # 保留原始列，创建新的期趸标识列
                original_column = f"{target_column}_原始"
                result_df[original_column] = result_df[target_column].copy()
                result_df[target_column] = term_text
            else:
                # 直接替换或创建期趸标识列
                result_df[target_column] = term_text

        return result_df

    def format_currency_columns(self, df: pd.DataFrame,
                               currency_columns: List[str] = None,
                               symbol: str = "¥") -> pd.DataFrame:
        """
        格式化货币列

        Args:
            df: 原始DataFrame
            currency_columns: 货币列名列表
            symbol: 货币符号

        Returns:
            格式化后的DataFrame
        """
        result_df = df.copy()

        if currency_columns is None:
            # 根据列名推断货币列
            currency_columns = []
            for col in result_df.columns:
                if any(keyword in col for keyword in ['保费', '金额', '费用', '价格', 'Price', 'Amount']):
                    currency_columns.append(col)

        for col in currency_columns:
            if col in result_df.columns:
                result_df[col] = result_df[col].apply(
                    lambda x: f"{symbol}{self._format_number(x, 2)}" if pd.notna(x) else x
                )

        return result_df

    def standardize_column_names(self, df: pd.DataFrame,
                                column_mapping: Dict[str, str] = None) -> pd.DataFrame:
        """
        标准化列名

        Args:
            df: 原始DataFrame
            column_mapping: 列名映射字典

        Returns:
            标准化列名后的DataFrame
        """
        result_df = df.copy()

        if column_mapping:
            result_df = result_df.rename(columns=column_mapping)
        else:
            # 应用默认的列名标准化
            default_mapping = {
                'FYP': '标保',
                'fyp': '标保',
                'FYC': '佣金',
                'fyc': '佣金',
                'Premium': '保费',
                'premium': '保费'
            }
            result_df = result_df.rename(columns=default_mapping)

        return result_df

    def validate_data_integrity(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        验证数据完整性

        Args:
            df: 要验证的DataFrame

        Returns:
            验证结果字典
        """
        validation_result = {
            'is_valid': True,
            'errors': [],
            'warnings': [],
            'stats': {
                'total_rows': len(df),
                'total_columns': len(df.columns),
                'null_values': df.isnull().sum().sum(),
                'duplicate_rows': df.duplicated().sum()
            }
        }

        # 检查空DataFrame
        if df.empty:
            validation_result['is_valid'] = False
            validation_result['errors'].append("数据为空")
            return validation_result

        # 检查是否有重复的列名
        if df.columns.duplicated().any():
            validation_result['is_valid'] = False
            validation_result['errors'].append("存在重复的列名")

        # 检查数值列是否包含异常值
        numeric_columns = df.select_dtypes(include=[np.number]).columns
        for col in numeric_columns:
            if df[col].isnull().all():
                validation_result['warnings'].append(f"数值列 '{col}' 全为空值")
            elif (df[col] < 0).any() and '保费' in col:
                validation_result['warnings'].append(f"保费列 '{col}' 包含负值")

        return validation_result

    def _format_number(self, value: Any, precision: int = 2) -> str:
        """
        格式化数字为千位分隔符形式

        Args:
            value: 要格式化的数值
            precision: 小数位数

        Returns:
            格式化后的字符串
        """
        try:
            if pd.isna(value):
                return str(value)

            # 转换为数值
            numeric_value = float(value)

            # 智能检测是否需要小数位
            if precision == 0:
                # 精度设为0，总是格式化为整数
                formatted_value = f"{int(round(numeric_value)):,}"
            elif numeric_value == int(numeric_value):
                # 如果数值本身是整数，不显示小数位
                formatted_value = f"{int(numeric_value):,}"
            else:
                # 如果数值有小数部分，按指定精度格式化
                formatted_value = f"{numeric_value:,.{precision}f}"

            return formatted_value

        except (ValueError, TypeError):
            return str(value)

    def calculate_chengbao_term_data(self, base_data: pd.DataFrame,
                                     sfyp2_column: str = "SFYP2(不含短险续保)",
                                     premium_column: str = "首年保费",
                                     target_column: str = "承保趸期数据",
                                     policy_number_column: str = "保单号") -> tuple[pd.DataFrame, List[Dict]]:
        """
        在内存中计算承保趸期数据

        根据SFYP2(不含短险续保)和首年保费的比例计算趸期数据：
        - R = 0.1：输出"趸交FYP"
        - R = 1：需要用户输入，收集行号和保单号
        - R = 0.2~0.9：计算X=R×10，输出"X年交SFYP"

        Args:
            base_data: 基础数据
            sfyp2_column: SFYP2列名，默认为"SFYP2(不含短险续保)"
            premium_column: 首年保费列名，默认为"首年保费"
            target_column: 目标承保趸期数据列名，默认为"承保趸期数据"
            policy_number_column: 保单号列名，默认为"保单号"

        Returns:
            tuple: (处理后的DataFrame, 需要用户输入的行列表)
                   行列表元素格式：{'row_index': 行号, 'policy_number': 保单号}
        """
        result_df = base_data.copy()
        pending_rows = []  # 存储需要用户输入的行

        # 检查必要列是否存在
        if sfyp2_column not in result_df.columns:
            # 记录警告但不影响其他功能
            import logging
            logger = logging.getLogger(__name__)
            logger.warning(f"未找到列'{sfyp2_column}'，跳过承保趸期数据计算")
            return result_df, pending_rows

        if premium_column not in result_df.columns:
            import logging
            logger = logging.getLogger(__name__)
            logger.warning(f"未找到列'{premium_column}'，跳过承保趸期数据计算")
            return result_df, pending_rows

        # 计算承保趸期数据
        chengbao_term_data = []
        for index, row in result_df.iterrows():
            sfyp2_value = row[sfyp2_column]
            premium_value = row[premium_column]

            # 处理空值或无效值
            if pd.isna(sfyp2_value) or pd.isna(premium_value):
                chengbao_term_data.append("")
                continue

            # 转换为数值类型（处理可能包含千位分隔符的字符串）
            try:
                # 如果是字符串类型，先移除千位分隔符（逗号、空格等）
                if isinstance(sfyp2_value, str):
                    sfyp2_clean = sfyp2_value.replace(',', '').replace(' ', '')
                    sfyp2_num = float(sfyp2_clean)
                else:
                    sfyp2_num = float(sfyp2_value)

                if isinstance(premium_value, str):
                    premium_clean = premium_value.replace(',', '').replace(' ', '')
                    premium_num = float(premium_clean)
                else:
                    premium_num = float(premium_value)
            except (ValueError, TypeError) as e:
                import logging
                logger = logging.getLogger(__name__)
                logger.warning(f"第{index+1}行数据类型转换失败: {str(e)}")
                chengbao_term_data.append("")
                continue

            # 检查除零错误
            if premium_num == 0:
                chengbao_term_data.append("")
                import logging
                logger = logging.getLogger(__name__)
                logger.warning(f"第{index+1}行首年保费为0，跳过计算")
                continue

            # 计算比例 R = SFYP2 / 首年保费
            ratio = sfyp2_num / premium_num

            # 根据比例值进行分类处理
            if abs(ratio - 0.1) < 0.001:  # R = 0.1
                chengbao_term_data.append("趸交FYP")
            elif abs(ratio - 1.0) < 0.001:  # R = 1
                chengbao_term_data.append("")  # 暂不填值，等待用户输入

                # 收集需要用户输入的行信息
                policy_number = ""
                if policy_number_column in result_df.columns:
                    policy_number = str(row[policy_number_column]) if not pd.isna(row[policy_number_column]) else ""

                pending_rows.append({
                    'row_index': index,
                    'policy_number': policy_number
                })
            elif 0.2 <= ratio <= 0.9:  # R在0.2~0.9范围内
                # 计算X = R × 10
                years = int(round(ratio * 10))
                chengbao_term_data.append(f"{years}年交SFYP")
            else:  # 其他情况
                # 比例不在预期范围内，返回空值
                chengbao_term_data.append("")

        # 将承保趸期数据列添加到DataFrame
        result_df[target_column] = chengbao_term_data

        # 记录处理结果
        import logging
        logger = logging.getLogger(__name__)
        logger.info(f"承保趸期数据计算完成，处理{len(result_df)}行数据")
        logger.info(f"需要用户输入的行数：{len(pending_rows)}")

        return result_df, pending_rows