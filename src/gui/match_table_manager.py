"""
MatchTableManager — 管理匹配结果表格的填充、同步和操作

从 main_window.py 中拆分出来的匹配表格管理模块，负责:
- 填充匹配结果表格
- 同步匹配状态与 ExactMatcher
- 添加匹配行到表格
"""

import logging
from PySide6.QtWidgets import QTableWidgetItem

logger = logging.getLogger(__name__)


class MatchTableManager:
    """管理匹配结果表格的填充和同步"""

    PLACEHOLDER_PREFIXES = ['ph_', 'PH_', 'ph', 'PH']

    def __init__(self, match_table, exact_matcher):
        self.match_table = match_table
        self.exact_matcher = exact_matcher

    def populate_table(self, placeholders, data_columns, matching_rules=None):
        """
        填充匹配结果表格

        Args:
            placeholders: PPT 占位符列表
            data_columns: 数据列名列表
            matching_rules: 匹配规则字典 (可选)
        """
        self.match_table.setRowCount(len(placeholders))

        for i, placeholder in enumerate(placeholders):
            self.match_table.setItem(i, 0, QTableWidgetItem(placeholder))

            matched_column = self._match_column(placeholder, data_columns, matching_rules)
            self.match_table.setItem(i, 1, QTableWidgetItem(matched_column))

    def _match_column(self, placeholder, data_columns, matching_rules=None):
        """为占位符匹配数据列"""
        if matching_rules and placeholder in matching_rules:
            return matching_rules[placeholder]

        # 移除占位符前缀
        placeholder_clean = placeholder
        for prefix in self.PLACEHOLDER_PREFIXES:
            if placeholder.startswith(prefix):
                placeholder_clean = placeholder[len(prefix):]
                break

        # 精确匹配
        if placeholder_clean in data_columns:
            return placeholder_clean

        # 大小写不敏感匹配
        for column in data_columns:
            if column.lower() == placeholder_clean.lower():
                return column

        return ""

    def sync_to_exact_matcher(self, text_addition_rules=None):
        """
        将表格中的匹配状态同步到 ExactMatcher

        Args:
            text_addition_rules: 文本增加规则字典 (可选)
        """
        try:
            self.exact_matcher.clear_matching_rules()
            matched_count = 0

            for row in range(self.match_table.rowCount()):
                placeholder_item = self.match_table.item(row, 0)
                column_item = self.match_table.item(row, 1)
                if not placeholder_item or not column_item:
                    continue

                placeholder = placeholder_item.text()
                column_text = column_item.text()

                if not placeholder or not column_text or column_text == "未匹配":
                    continue

                # 解析自定义文本格式: [原始列名] 前缀:xxx 后缀:xxx
                data_column = column_text
                if data_column.startswith("[") and "]" in data_column:
                    bracket_end = data_column.find("]")
                    data_column = data_column[1:bracket_end]

                if data_column:
                    self.exact_matcher.set_matching_rule(placeholder, data_column)
                    matched_count += 1

            logger.info(f"已同步匹配规则: {matched_count} 条")

            # 同步文本增加规则
            if text_addition_rules:
                self.exact_matcher.text_addition_rules = {}
                for placeholder, rule in text_addition_rules.items():
                    if isinstance(rule, dict):
                        prefix = rule.get('prefix', '')
                        suffix = rule.get('suffix', '')
                        self.exact_matcher.set_text_addition_rule(placeholder, prefix, suffix)

        except Exception as e:
            logger.error(f"同步到 ExactMatcher 失败: {e}", exc_info=True)

    def sync_from_exact_matcher(self):
        """
        从 ExactMatcher 同步匹配状态回表格

        Returns:
            dict: 文本增加规则字典
        """
        try:
            matching_rules = self.exact_matcher.get_matching_rules()
            if matching_rules:
                self.match_table.setRowCount(0)
                for placeholder, column in matching_rules.items():
                    self._add_match_row(placeholder, column)

                text_rules = self.exact_matcher.get_all_text_addition_rules()
                logger.info(f"已从 ExactMatcher 同步 {len(matching_rules)} 条匹配规则")
                return text_rules
            else:
                logger.info("ExactMatcher 中没有匹配规则")
                return {}

        except Exception as e:
            logger.error(f"从 ExactMatcher 同步失败: {e}", exc_info=True)
            return {}

    def _add_match_row(self, placeholder, column):
        """添加匹配行到表格"""
        try:
            row = self.match_table.rowCount()
            self.match_table.insertRow(row)

            self.match_table.setItem(row, 0, QTableWidgetItem(placeholder))
            self.match_table.setItem(row, 1, QTableWidgetItem(column))

            text_rule = self.exact_matcher.get_text_addition_rule(placeholder)
            prefix = text_rule.get('prefix', '')
            suffix = text_rule.get('suffix', '')
            text_preview = f"'{prefix}'内容'{suffix}'" if prefix or suffix else ""
            self.match_table.setItem(row, 2, QTableWidgetItem(text_preview))

        except Exception as e:
            logger.error(f"添加匹配行失败: {e}")
