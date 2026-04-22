"""
占位符配置模块
负责管理占位符匹配规则的配置
"""

import os
import json
import logging
from typing import Dict, List, Any, Set
from datetime import datetime

logger = logging.getLogger(__name__)


class PlaceholderConfigMixin:
    """占位符配置混合类，提供占位符匹配规则的配置管理功能"""

    def save_placeholder_config(self, file_path: str, matching_rules: List[Dict[str, Any]],
                               template_info: Dict[str, Any], data_info: Dict[str, Any]) -> bool:
        """
        保存占位符配置到文件

        Args:
            file_path: 配置文件保存路径
            matching_rules: 匹配规则列表，格式：[{"placeholder": "ph_name", "column": "姓名", "text_addition": {"prefix": "", "suffix": "先生"}}]
            template_info: 模板信息，包含文件名和占位符列表
            data_info: 数据信息，包含文件名和字段列表

        Returns:
            bool: 保存成功返回 True，失败返回 False
        """
        try:
            if not matching_rules:
                logger.error("没有有效的匹配规则可以保存")
                return False

            # 构建配置数据结构
            config_data = {
                "version": "1.0.0",
                "created_at": datetime.now().isoformat(),
                "metadata": {
                    "template_info": template_info,
                    "data_info": data_info,
                    "rule_count": len(matching_rules)
                },
                "matching_rules": matching_rules
            }

            # 确保文件扩展名为.pptcfg
            if not file_path.endswith('.pptcfg'):
                file_path += '.pptcfg'

            # 保存配置文件
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)

            logger.info(f"成功保存占位符配置到：{file_path}")
            return True

        except Exception as e:
            logger.error(f"保存占位符配置失败：{str(e)}")
            return False

    def load_placeholder_config(self, file_path: str) -> Dict[str, Any]:
        """
        从文件加载占位符配置

        Args:
            file_path: 配置文件路径

        Returns:
            Dict[str, Any]: 包含配置数据和验证结果的字典
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(file_path):
                return {
                    "success": False,
                    "error": f"配置文件不存在：{file_path}"
                }

            # 读取配置文件
            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)

            # 验证配置文件格式
            validation_result = self._validate_placeholder_config_format(config_data)
            if not validation_result["valid"]:
                return {
                    "success": False,
                    "error": f"配置文件格式验证失败：{validation_result['error']}"
                }

            logger.info(f"成功加载占位符配置：{file_path}")
            return {
                "success": True,
                "config_data": config_data
            }

        except json.JSONDecodeError as e:
            logger.error(f"JSON 格式错误：{str(e)}")
            return {
                "success": False,
                "error": f"JSON 格式错误：{str(e)}"
            }
        except Exception as e:
            logger.error(f"加载占位符配置失败：{str(e)}")
            return {
                "success": False,
                "error": f"加载失败：{str(e)}"
            }

    def validate_placeholder_compatibility(self, config_data: Dict[str, Any],
                                         current_template_placeholders: List[str],
                                         current_data_columns: List[str]) -> Dict[str, Any]:
        """
        验证配置与当前模板和数据的兼容性

        Args:
            config_data: 配置数据
            current_template_placeholders: 当前 PPT 模板中的占位符列表
            current_data_columns: 当前 Excel 数据中的列名列表

        Returns:
            Dict[str, Any]: 兼容性验证结果
        """
        try:
            config_placeholders = set()
            config_columns = set()

            # 提取配置中的占位符和数据字段
            for rule in config_data.get("matching_rules", []):
                placeholder = rule.get("placeholder", "")
                column = rule.get("column", "")
                if placeholder:
                    config_placeholders.add(placeholder)
                if column:
                    config_columns.add(column)

            current_placeholders = set(current_template_placeholders)
            current_columns = set(current_data_columns)

            # 计算兼容性
            matched_placeholders = current_placeholders & config_placeholders
            matched_columns = current_columns & config_columns

            placeholder_compatibility_rate = (
                len(matched_placeholders) / len(current_placeholders)
                if current_placeholders else 1.0
            )
            data_compatibility_rate = (
                len(matched_columns) / len(current_columns)
                if current_columns else 1.0
            )
            overall_compatibility_rate = min(placeholder_compatibility_rate, data_compatibility_rate)

            # 确定兼容的匹配规则
            compatible_rules = []
            incompatible_rules = []

            for rule in config_data.get("matching_rules", []):
                placeholder = rule.get("placeholder", "")
                column = rule.get("column", "")

                if placeholder in current_placeholders and column in current_columns:
                    compatible_rules.append(rule)
                else:
                    incompatible_rules.append({
                        "rule": rule,
                        "reason": self._get_incompatibility_reason(
                            placeholder, column, current_placeholders, current_columns
                        )
                    })

            return {
                "compatible": overall_compatibility_rate >= 0.5,
                "overall_compatibility_rate": overall_compatibility_rate,
                "placeholder_compatibility_rate": placeholder_compatibility_rate,
                "data_compatibility_rate": data_compatibility_rate,
                "compatible_rules": compatible_rules,
                "incompatible_rules": incompatible_rules,
                "statistics": {
                    "total_placeholders": len(current_placeholders),
                    "matched_placeholders": len(matched_placeholders),
                    "total_columns": len(current_columns),
                    "matched_columns": len(matched_columns),
                    "total_rules": len(config_data.get("matching_rules", [])),
                    "compatible_rules_count": len(compatible_rules),
                    "incompatible_rules_count": len(incompatible_rules)
                }
            }

        except Exception as e:
            logger.error(f"兼容性验证失败：{str(e)}")
            return {
                "compatible": False,
                "error": f"兼容性验证失败：{str(e)}",
                "overall_compatibility_rate": 0.0
            }

    def _validate_placeholder_config_format(self, config_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        验证配置文件格式的正确性

        Args:
            config_data: 配置数据

        Returns:
            Dict[str, Any]: 验证结果
        """
        try:
            # 检查必需的顶级字段
            required_fields = ["version", "metadata", "matching_rules"]
            for field in required_fields:
                if field not in config_data:
                    return {
                        "valid": False,
                        "error": f"缺少必需字段：{field}"
                    }

            # 检查版本兼容性
            version = config_data.get("version", "")
            if version != "1.0.0":
                return {
                    "valid": False,
                    "error": f"不支持的配置文件版本：{version}，仅支持 1.0.0"
                }

            # 验证匹配规则格式
            matching_rules = config_data.get("matching_rules", [])
            if not isinstance(matching_rules, list):
                return {
                    "valid": False,
                    "error": "matching_rules 必须是列表格式"
                }

            for i, rule in enumerate(matching_rules):
                if not isinstance(rule, dict):
                    return {
                        "valid": False,
                        "error": f"第{i+1}个匹配规则必须是字典格式"
                    }

                # 检查必需字段
                if "placeholder" not in rule or "column" not in rule:
                    return {
                        "valid": False,
                        "error": f"第{i+1}个匹配规则缺少 placeholder 或 column 字段"
                    }

                # 验证占位符格式
                placeholder = rule["placeholder"]
                if not placeholder or not isinstance(placeholder, str):
                    return {
                        "valid": False,
                        "error": f"第{i+1}个匹配规则的占位符无效"
                    }

                # 验证文本增加规则格式
                text_addition = rule.get("text_addition", {})
                if text_addition and not isinstance(text_addition, dict):
                    return {
                        "valid": False,
                        "error": f"第{i+1}个匹配规则的 text_addition 必须是字典格式"
                    }

            return {
                "valid": True
            }

        except Exception as e:
            return {
                "valid": False,
                "error": f"格式验证异常：{str(e)}"
            }

    def _get_incompatibility_reason(self, placeholder: str, column: str,
                                 current_placeholders: Set[str], current_columns: Set[str]) -> str:
        """
        获取不兼容规则的具体原因

        Args:
            placeholder: 占位符
            column: 数据字段
            current_placeholders: 当前占位符集合
            current_columns: 当前数据字段集合

        Returns:
            str: 不兼容的原因描述
        """
        reasons = []

        if placeholder and placeholder not in current_placeholders:
            reasons.append(f"占位符 '{placeholder}' 不在当前模板中")

        if column and column not in current_columns:
            reasons.append(f"数据字段 '{column}' 不在当前 Excel 文件中")

        if not reasons:
            reasons.append("未知原因")

        return "；".join(reasons)

    def get_matching_config(self) -> Dict[str, Any]:
        """
        获取匹配配置

        Returns:
            Dict[str, Any]: 匹配配置
        """
        return self.get("matching", {})

    def update_matching_rules(self, rules: List[Dict[str, str]]):
        """
        更新匹配规则

        Args:
            rules: 匹配规则列表
        """
        self.set("matching.rules", rules)
        self.save_config()

    def get_matching_rules(self) -> List[Dict[str, str]]:
        """
        获取匹配规则

        Returns:
            List[Dict[str, str]]: 匹配规则列表
        """
        return self.get("matching.rules", [])
