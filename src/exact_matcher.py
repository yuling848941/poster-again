"""
精确匹配模块
用于将数据与PPT模板中的占位符进行精确匹配
"""

import re
import os
import logging
from typing import Dict, List, Tuple, Optional, Any
import pandas as pd

# 设置日志 - 使用模块级 logger，避免 basicConfig 全局配置冲突
logger = logging.getLogger(__name__)


class ExactMatcher:
    """精确匹配器，负责将数据与PPT模板中的占位符进行精确匹配"""
    
    def __init__(self):
        """初始化精确匹配器"""
        self.placeholder_pattern = re.compile(r'\{\{([^}]+)\}\}')
        self.data = None
        self.template_placeholders = []
        self.matching_rules = {}
        self.text_addition_rules = {}  # 新增：文本增加规则
        
    def set_data(self, data: pd.DataFrame):
        """
        设置数据源
        
        Args:
            data: 数据DataFrame
        """
        self.data = data
        logger.info(f"设置数据源，行数: {len(data)}, 列数: {len(data.columns)}")
        
    def extract_placeholders(self, template_text: str) -> List[str]:
        """
        从模板文本中提取占位符
        
        Args:
            template_text: 模板文本
            
        Returns:
            List[str]: 占位符列表
        """
        placeholders = self.placeholder_pattern.findall(template_text)
        unique_placeholders = list(set(placeholders))
        logger.info(f"从模板中提取到 {len(unique_placeholders)} 个唯一占位符")
        return unique_placeholders
    
    def set_template_placeholders(self, placeholders: List[str]):
        """
        设置模板占位符列表
        
        Args:
            placeholders: 占位符列表
        """
        self.template_placeholders = placeholders
        logger.info(f"设置模板占位符: {placeholders}")
    
    def find_column_for_placeholder(self, placeholder: str) -> Optional[str]:
        """
        为占位符找到最匹配的列名
        匹配策略：1.精确匹配 2.忽略大小写匹配 3.特殊字符清理匹配

        Args:
            placeholder: 占位符

        Returns:
            Optional[str]: 匹配的列名，如果没有找到返回None
        """
        if self.data is None:
            logger.error("数据源未设置")
            return None

        # 1. 首先尝试精确匹配
        if placeholder in self.data.columns:
            logger.info(f"占位符 '{placeholder}' 精确匹配列名 '{placeholder}' [匹配类型: 精确匹配]")
            return placeholder

        # 2. 尝试忽略大小写匹配
        for col in self.data.columns:
            if placeholder.lower() == col.lower():
                logger.info(f"占位符 '{placeholder}' 忽略大小写匹配列名 '{col}' [匹配类型: 忽略大小写]")
                return col

        # 3. 尝试特殊字符清理匹配（移除占位符前缀和特殊字符）
        # 移除常见占位符前缀
        placeholder_without_prefix = placeholder
        for prefix in ['ph_', 'PH_', 'ph', 'PH']:
            if placeholder.startswith(prefix):
                placeholder_without_prefix = placeholder[len(prefix):]
                break

        # 清理特殊字符
        clean_placeholder = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', placeholder_without_prefix)
        if clean_placeholder:
            for col in self.data.columns:
                clean_col = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]', '', col)
                if clean_placeholder.lower() == clean_col.lower():
                    logger.info(f"占位符 '{placeholder}' 清理后匹配列名 '{col}' [匹配类型: 特殊字符清理]")
                    return col

        logger.warning(f"占位符 '{placeholder}' 所有匹配策略均失败，未找到匹配的列")
        return None
    
    def auto_match_placeholders(self) -> Dict[str, str]:
        """
        自动匹配所有占位符到列
        
        Returns:
            Dict[str, str]: 占位符到列名的映射
        """
        if self.data is None:
            logger.error("数据源未设置")
            return {}
            
        if not self.template_placeholders:
            logger.error("模板占位符未设置")
            return {}
            
        matching_rules = {}
        for placeholder in self.template_placeholders:
            column = self.find_column_for_placeholder(placeholder)
            if column:
                matching_rules[placeholder] = column
                
        self.matching_rules = matching_rules
        logger.info(f"自动匹配完成，匹配了 {len(matching_rules)} 个占位符")
        return matching_rules
    
    def set_matching_rule(self, placeholder: str, column: str):
        """
        手动设置占位符与列的匹配规则
        
        Args:
            placeholder: 占位符
            column: 列名
        """
        if self.data is None:
            logger.error("数据源未设置")
            return
            
        if not column or column.strip() == "":
            logger.error(f"列名不能为空")
            return
            
        if column not in self.data.columns:
            logger.error(f"列名 '{column}' 不存在于数据中")
            return
            
        self.matching_rules[placeholder] = column
        logger.info(f"设置匹配规则: '{placeholder}' -> '{column}'")
    
    def set_matching_rules(self, matching_rules: Dict[str, str]):
        """
        批量设置匹配规则
        
        Args:
            matching_rules: 匹配规则字典，键为占位符，值为列名
        """
        if self.data is None:
            logger.error("数据源未设置")
            return
        
        for placeholder, column in matching_rules.items():
            if column in self.data.columns:
                self.matching_rules[placeholder] = column
            else:
                logger.warning(f"列名 '{column}' 不存在于数据中，跳过")
        
        logger.info(f"批量设置匹配规则，共 {len(matching_rules)} 条")
    
    def remove_matching_rule(self, placeholder: str):
        """
        移除占位符的匹配规则

        Args:
            placeholder: 占位符
        """
        if placeholder in self.matching_rules:
            del self.matching_rules[placeholder]
            logger.info(f"移除匹配规则: '{placeholder}'")

    def clear_matching_rules(self):
        """
        清除所有匹配规则
        """
        self.matching_rules = {}
        logger.info("清除所有匹配规则")

    def get_matching_rules(self) -> Dict[str, str]:
        """
        获取所有匹配规则
        
        Returns:
            Dict[str, str]: 占位符到列名的映射
        """
        return self.matching_rules.copy()
    
    def get_unmatched_placeholders(self) -> List[str]:
        """
        获取未匹配的占位符
        
        Returns:
            List[str]: 未匹配的占位符列表
        """
        return [p for p in self.template_placeholders if p not in self.matching_rules]
    
    def get_unmatched_columns(self) -> List[str]:
        """
        获取未匹配的列
        
        Returns:
            List[str]: 未匹配的列名列表
        """
        if self.data is None:
            return []
            
        matched_columns = set(self.matching_rules.values())
        return [col for col in self.data.columns if col not in matched_columns]
    
    def get_data_for_placeholder(self, placeholder: str, row_index: int = 0) -> Any:
        """
        获取指定占位符对应的数据值
        
        Args:
            placeholder: 占位符
            row_index: 行索引
            
        Returns:
            Any: 数据值
        """
        if placeholder not in self.matching_rules:
            logger.warning(f"占位符 '{placeholder}' 没有匹配的列")
            return None
            
        column = self.matching_rules[placeholder]
        if row_index >= len(self.data):
            logger.warning(f"行索引 {row_index} 超出数据范围")
            return None
            
        try:
            value = self.data.iloc[row_index][column]
            # 处理NaN值
            if pd.isna(value):
                return None
            return value
        except Exception as e:
            logger.error(f"获取数据失败: {str(e)}")
            return None
    
    def get_all_data_for_row(self, row_index: int = 0) -> Dict[str, Any]:
        """
        获取指定行的所有匹配数据
        
        Args:
            row_index: 行索引
            
        Returns:
            Dict[str, Any]: 占位符到数据值的映射
        """
        if row_index >= len(self.data):
            logger.warning(f"行索引 {row_index} 超出数据范围")
            return {}
            
        result = {}
        for placeholder in self.matching_rules:
            # 获取原始数据值
            original_value = self.get_data_for_placeholder(placeholder, row_index)
            # 应用文本增加规则
            processed_value = self.apply_text_addition(placeholder, original_value)
            result[placeholder] = processed_value
            
        return result
    
    def validate_matching(self) -> Dict[str, Any]:
        """
        验证匹配结果
        
        Returns:
            Dict[str, Any]: 验证结果
        """
        if self.data is None:
            return {
                "valid": False,
                "message": "数据源未设置"
            }
            
        if not self.template_placeholders:
            return {
                "valid": False,
                "message": "模板占位符未设置"
            }
            
        unmatched_placeholders = self.get_unmatched_placeholders()
        
        return {
            "valid": len(unmatched_placeholders) == 0,
            "total_placeholders": len(self.template_placeholders),
            "matched_placeholders": len(self.matching_rules),
            "unmatched_placeholders": unmatched_placeholders,
            "message": f"匹配完成，{len(unmatched_placeholders)} 个占位符未匹配" if unmatched_placeholders else "所有占位符已匹配"
        }
    
    def export_matching_rules(self) -> List[Dict[str, str]]:
        """
        导出匹配规则为列表格式
        
        Returns:
            List[Dict[str, str]]: 匹配规则列表
        """
        return [
            {"placeholder": placeholder, "column": column}
            for placeholder, column in self.matching_rules.items()
        ]
    
    def import_matching_rules(self, rules: List[Dict[str, str]]):
        """
        从列表格式导入匹配规则
        
        Args:
            rules: 匹配规则列表
        """
        self.matching_rules = {}
        for rule in rules:
            placeholder = rule.get("placeholder")
            column = rule.get("column")
            if placeholder and column:
                self.matching_rules[placeholder] = column
                
        logger.info(f"导入 {len(self.matching_rules)} 条匹配规则")
    
    def set_text_addition_rule(self, placeholder: str, prefix: str = "", suffix: str = ""):
        """
        设置占位符的文本增加规则
        
        Args:
            placeholder: 占位符
            prefix: 前缀文本
            suffix: 后缀文本
        """
        if not placeholder:
            logger.error("占位符不能为空")
            return
            
        self.text_addition_rules[placeholder] = {
            "prefix": prefix,
            "suffix": suffix
        }
        logger.info(f"设置文本增加规则: '{placeholder}' -> 前缀:'{prefix}', 后缀:'{suffix}'")
    
    def remove_text_addition_rule(self, placeholder: str):
        """
        移除占位符的文本增加规则
        
        Args:
            placeholder: 占位符
        """
        if placeholder in self.text_addition_rules:
            del self.text_addition_rules[placeholder]
            logger.info(f"移除文本增加规则: '{placeholder}'")
    
    def get_text_addition_rule(self, placeholder: str) -> Dict[str, str]:
        """
        获取占位符的文本增加规则
        
        Args:
            placeholder: 占位符
            
        Returns:
            Dict[str, str]: 包含prefix和suffix的字典，如果没有规则返回空字典
        """
        return self.text_addition_rules.get(placeholder, {"prefix": "", "suffix": ""})
    
    def get_all_text_addition_rules(self) -> Dict[str, Dict[str, str]]:
        """
        获取所有文本增加规则
        
        Returns:
            Dict[str, Dict[str, str]]: 占位符到文本增加规则的映射
        """
        return self.text_addition_rules.copy()
    
    def apply_text_addition(self, placeholder: str, value: Any) -> str:
        """
        对值应用文本增加规则
        
        Args:
            placeholder: 占位符
            value: 原始值
            
        Returns:
            str: 应用文本增加规则后的值
        """
        if value is None:
            return ""
            
        # 将值转换为字符串
        str_value = str(value)
        
        # 获取文本增加规则
        rule = self.get_text_addition_rule(placeholder)
        prefix = rule.get("prefix", "")
        suffix = rule.get("suffix", "")
        
        # 应用前缀和后缀
        result = f"{prefix}{str_value}{suffix}"

        return result

    # ==================== 配置导出和导入方法 ====================

    def export_matching_config(self, template_file_path: str = "", data_file_path: str = "") -> List[Dict[str, Any]]:
        """
        导出当前的匹配配置为标准格式

        Args:
            template_file_path: PPT模板文件路径
            data_file_path: Excel数据文件路径

        Returns:
            List[Dict[str, Any]]: 匹配规则列表
        """
        try:
            matching_rules = []

            # 遍历所有匹配规则
            for placeholder, column in self.matching_rules.items():
                rule = {
                    "placeholder": placeholder,
                    "column": column
                }

                # 添加文本增加规则
                text_addition_rule = self.get_text_addition_rule(placeholder)
                if text_addition_rule.get("prefix") or text_addition_rule.get("suffix"):
                    rule["text_addition"] = {
                        "prefix": text_addition_rule.get("prefix", ""),
                        "suffix": text_addition_rule.get("suffix", "")
                    }

                matching_rules.append(rule)

            logger.info(f"成功导出 {len(matching_rules)} 条匹配规则")
            return matching_rules

        except Exception as e:
            logger.error(f"导出匹配配置失败: {str(e)}")
            return []

    def import_matching_config(self, matching_rules: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        导入匹配配置

        Args:
            matching_rules: 匹配规则列表

        Returns:
            Dict[str, Any]: 导入结果统计
        """
        try:
            imported_count = 0
            skipped_count = 0
            error_count = 0
            errors = []

            # 清除现有匹配规则
            self.clear_matching_rules()
            self.text_addition_rules = {}

            # 导入新的匹配规则
            for rule in matching_rules:
                try:
                    placeholder = rule.get("placeholder", "")
                    column = rule.get("column", "")

                    if not placeholder or not column:
                        skipped_count += 1
                        errors.append(f"跳过无效规则: 缺少占位符或字段名")
                        continue

                    # 验证数据字段是否存在
                    if self.data is not None and column not in self.data.columns:
                        skipped_count += 1
                        errors.append(f"跳过规则: 字段 '{column}' 不在当前数据中")
                        continue

                    # 设置匹配规则
                    self.set_matching_rule(placeholder, column)

                    # 设置文本增加规则
                    text_addition = rule.get("text_addition", {})
                    if isinstance(text_addition, dict):
                        prefix = text_addition.get("prefix", "")
                        suffix = text_addition.get("suffix", "")
                        if prefix or suffix:
                            self.set_text_addition_rule(placeholder, prefix, suffix)

                    imported_count += 1

                except Exception as rule_error:
                    error_count += 1
                    errors.append(f"规则导入错误: {str(rule_error)}")

            result = {
                "success": True,
                "imported_count": imported_count,
                "skipped_count": skipped_count,
                "error_count": error_count,
                "errors": errors
            }

            logger.info(f"配置导入完成: 导入 {imported_count} 条，跳过 {skipped_count} 条，错误 {error_count} 条")
            return result

        except Exception as e:
            logger.error(f"导入匹配配置失败: {str(e)}")
            return {
                "success": False,
                "error": str(e)
            }

    def get_template_info(self, template_file_path: str = "") -> Dict[str, Any]:
        """
        获取模板信息

        Args:
            template_file_path: 模板文件路径

        Returns:
            Dict[str, Any]: 模板信息
        """
        return {
            "file_name": os.path.basename(template_file_path) if template_file_path else "",
            "placeholders": self.template_placeholders.copy(),
            "placeholder_count": len(self.template_placeholders)
        }

    def get_data_info(self, data_file_path: str = "") -> Dict[str, Any]:
        """
        获取数据信息

        Args:
            data_file_path: 数据文件路径

        Returns:
            Dict[str, Any]: 数据信息
        """
        if self.data is None:
            return {
                "file_name": os.path.basename(data_file_path) if data_file_path else "",
                "columns": [],
                "column_count": 0,
                "row_count": 0
            }

        return {
            "file_name": os.path.basename(data_file_path) if data_file_path else "",
            "columns": list(self.data.columns),
            "column_count": len(self.data.columns),
            "row_count": len(self.data)
        }

    def validate_config_compatibility_quick(self, matching_rules: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        快速验证配置兼容性

        Args:
            matching_rules: 匹配规则列表

        Returns:
            Dict[str, Any]: 兼容性验证结果
        """
        try:
            compatible_rules = []
            incompatible_rules = []

            for rule in matching_rules:
                placeholder = rule.get("placeholder", "")
                column = rule.get("column", "")

                if placeholder in self.template_placeholders and (
                    self.data is None or column in self.data.columns
                ):
                    compatible_rules.append(rule)
                else:
                    incompatible_rules.append({
                        "rule": rule,
                        "reason": self._get_quick_incompatibility_reason(placeholder, column)
                    })

            total_rules = len(matching_rules)
            compatibility_rate = len(compatible_rules) / total_rules if total_rules > 0 else 0

            return {
                "compatible": compatibility_rate > 0,
                "compatibility_rate": compatibility_rate,
                "compatible_rules": compatible_rules,
                "incompatible_rules": incompatible_rules,
                "statistics": {
                    "total_rules": total_rules,
                    "compatible_count": len(compatible_rules),
                    "incompatible_count": len(incompatible_rules)
                }
            }

        except Exception as e:
            logger.error(f"快速兼容性验证失败: {str(e)}")
            return {
                "compatible": False,
                "error": str(e),
                "compatibility_rate": 0.0
            }

    def _get_quick_incompatibility_reason(self, placeholder: str, column: str) -> str:
        """
        获取快速不兼容原因

        Args:
            placeholder: 占位符
            column: 数据字段

        Returns:
            str: 不兼容原因
        """
        reasons = []

        if placeholder and placeholder not in self.template_placeholders:
            reasons.append(f"占位符 '{placeholder}' 不在当前模板中")

        if self.data is not None and column and column not in self.data.columns:
            reasons.append(f"数据字段 '{column}' 不在当前Excel文件中")

        if not reasons:
            reasons.append("未知原因")

        return "；".join(reasons)