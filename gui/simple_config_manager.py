"""
简化的配置管理器
直接基于GUI状态进行配置保存和加载
"""

import os
import json
import logging
from datetime import datetime
from typing import Dict, Any, Optional
from PySide6.QtWidgets import QMessageBox, QTableWidgetItem

logger = logging.getLogger(__name__)


class SimpleConfigManager:
    """简化的配置管理器，直接操作GUI状态"""

    def __init__(self):
        """初始化简化配置管理器"""
        self.config_version = "1.0.0"

    def save_gui_config(self, main_window, file_path: str) -> bool:
        """
        从主窗口直接保存GUI状态配置

        Args:
            main_window: 主窗口实例
            file_path: 保存文件路径

        Returns:
            bool: 保存成功返回True，失败返回False
        """
        try:
            # 收集GUI状态数据
            gui_config = self._collect_gui_state(main_window)

            if not self._has_valid_settings(gui_config):
                QMessageBox.warning(
                    None, "保存配置",
                    "当前没有有效的设置可以保存！\n请先添加文本增加设置或配置递交趸期数据。"
                )
                return False

            # 构建配置文件数据
            config_data = {
                "version": self.config_version,
                "created_at": datetime.now().isoformat(),
                "gui_settings": gui_config
            }

            # 确保文件扩展名为.pptcfg
            if not file_path.endswith('.pptcfg'):
                file_path += '.pptcfg'

            # 保存配置文件
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)

            logger.info(f"成功保存GUI配置到: {file_path}")
            return True

        except Exception as e:
            logger.error(f"保存GUI配置失败: {str(e)}")
            QMessageBox.critical(None, "保存失败", f"保存配置时发生错误:\n{str(e)}")
            return False

    def load_gui_config(self, main_window, file_path: str) -> bool:
        """
        从配置文件加载并恢复GUI状态

        Args:
            main_window: 主窗口实例
            file_path: 配置文件路径

        Returns:
            bool: 加载成功返回True，失败返回False
        """
        try:
            if not os.path.exists(file_path):
                QMessageBox.warning(None, "文件不存在", f"配置文件不存在:\n{file_path}")
                return False

            # 读取配置文件
            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)

            # 验证配置文件格式
            if not self._validate_config_format(config_data):
                QMessageBox.error(None, "格式错误", "配置文件格式不正确或版本不兼容！")
                return False

            # 提取GUI设置
            gui_settings = config_data.get('gui_settings', {})

            # 恢复GUI状态
            success = self._restore_gui_state(main_window, gui_settings)

            if success:
                logger.info(f"成功加载GUI配置: {file_path}")
                return True
            else:
                return False

        except json.JSONDecodeError as e:
            QMessageBox.error(None, "文件错误", f"配置文件格式错误:\n{str(e)}")
            return False
        except Exception as e:
            logger.error(f"加载GUI配置失败: {str(e)}")
            QMessageBox.critical(None, "加载失败", f"加载配置时发生错误:\n{str(e)}")
            return False

    def _collect_gui_state(self, main_window) -> Dict[str, Any]:
        """收集主窗口的GUI状态"""
        gui_config = {
            "text_additions": {},
            "dundian_mappings": {},
            "chengbao_mappings": {}
        }

        try:
            # 收集文本增加设置
            if hasattr(main_window, 'text_addition_rules'):
                text_additions = main_window.text_addition_rules
                if text_additions:
                    for placeholder, rule in text_additions.items():
                        if isinstance(rule, dict):
                            gui_config["text_additions"][placeholder] = {
                                "prefix": rule.get('prefix', ''),
                                "suffix": rule.get('suffix', '')
                            }

            # 收集递交趸期映射关系
            if hasattr(main_window, 'dundian_mappings'):
                gui_config["dundian_mappings"] = main_window.dundian_mappings.copy()

            # 收集承保趸期映射关系
            if hasattr(main_window, 'chengbao_mappings'):
                gui_config["chengbao_mappings"] = main_window.chengbao_mappings.copy()

            logger.info(f"收集GUI状态: 文本规则 {len(gui_config['text_additions'])}, 趸期映射 {len(gui_config['dundian_mappings'])}, 承保映射 {len(gui_config['chengbao_mappings'])}")

        except Exception as e:
            logger.error(f"收集GUI状态失败: {str(e)}")

        return gui_config

    def _collect_dundian_settings(self, main_window) -> Dict[str, Any]:
        """收集递交趸期设置"""
        dundian_settings = {}

        try:
            # 检查是否有递交趸期相关的控件
            if hasattr(main_window, 'dundian_checkbox'):
                dundian_settings["enabled"] = main_window.dundian_checkbox.isChecked()
            else:
                dundian_settings["enabled"] = False

            if hasattr(main_window, 'dundian_type_combo'):
                dundian_settings["default_type"] = main_window.dundian_type_combo.currentText()
            else:
                dundian_settings["default_type"] = "期缴"

        except Exception as e:
            logger.error(f"收集递交趸期设置失败: {str(e)}")
            dundian_settings = {"enabled": False, "default_type": "期缴"}

        return dundian_settings

    def _has_valid_settings(self, gui_config: Dict[str, Any]) -> bool:
        """检查是否有有效的设置"""
        text_additions = gui_config.get('text_additions', {})
        dundian_mappings = gui_config.get('dundian_mappings', {})
        chengbao_mappings = gui_config.get('chengbao_mappings', {})

        # 有文本增加规则或有递交趸期映射或有承保趸期映射都算有效
        has_text_rules = len(text_additions) > 0
        has_dundian_mappings = len(dundian_mappings) > 0
        has_chengbao_mappings = len(chengbao_mappings) > 0

        return has_text_rules or has_dundian_mappings or has_chengbao_mappings

    def _validate_config_format(self, config_data: Dict[str, Any]) -> bool:
        """验证配置文件格式"""
        try:
            # 检查必需字段
            if 'version' not in config_data:
                return False

            # 检查版本兼容性
            version = config_data.get('version', '')
            if version != self.config_version:
                return False

            # 检查gui_settings字段
            if 'gui_settings' not in config_data:
                return False

            return True

        except Exception:
            return False

    def _restore_gui_state(self, main_window, gui_settings: Dict[str, Any]) -> bool:
        """恢复GUI状态"""
        try:
            restored_count = 0

            # 恢复文本增加设置
            text_additions = gui_settings.get('text_additions', {})
            if text_additions:
                if not hasattr(main_window, 'text_addition_rules'):
                    main_window.text_addition_rules = {}

                for placeholder, rule in text_additions.items():
                    if isinstance(rule, dict):
                        main_window.text_addition_rules[placeholder] = {
                            'prefix': rule.get('prefix', ''),
                            'suffix': rule.get('suffix', '')
                        }
                        restored_count += 1
                        logger.info(f"恢复文本增加设置: {placeholder} -> {rule}")

            # 恢复递交趸期映射关系
            dundian_mappings = gui_settings.get('dundian_mappings', {})
            if dundian_mappings:
                if not hasattr(main_window, 'dundian_mappings'):
                    main_window.dundian_mappings = {}

                for placeholder, data_column in dundian_mappings.items():
                    # 只有在占位符名称匹配时才恢复
                    if self._placeholder_exists(main_window, placeholder):
                        main_window.dundian_mappings[placeholder] = data_column
                        # 更新匹配表格显示
                        self._update_match_table_for_dundian(main_window, placeholder, data_column)
                        restored_count += 1
                        logger.info(f"恢复递交趸期映射: {placeholder} -> {data_column}")
                    else:
                        logger.info(f"跳过不匹配的占位符: {placeholder}")

            # 恢复承保趸期映射关系
            chengbao_mappings = gui_settings.get('chengbao_mappings', {})
            if chengbao_mappings:
                if not hasattr(main_window, 'chengbao_mappings'):
                    main_window.chengbao_mappings = {}

                for placeholder, data_column in chengbao_mappings.items():
                    # 只有在占位符名称匹配时才恢复
                    if self._placeholder_exists(main_window, placeholder):
                        main_window.chengbao_mappings[placeholder] = data_column
                        # 更新匹配表格显示
                        self._update_match_table_for_chengbao(main_window, placeholder, data_column)
                        restored_count += 1
                        logger.info(f"恢复承保趸期映射: {placeholder} -> {data_column}")
                    else:
                        logger.info(f"跳过不匹配的占位符: {placeholder}")

                # ✅ 触发承保趸期数据计算（用于用户输入）
                # 注意：这里使用防重复触发机制，避免重复计算
                if restored_count > 0 and len(main_window.chengbao_mappings) > 0:
                    logger.info("检测到承保趸期映射，开始自动计算承保趸期数据...")
                    self._trigger_chengbao_calculation(main_window)

            # 更新GUI显示和同步状态
            if hasattr(main_window, 'log_text') and main_window.log_text:
                main_window.log_text.append(f"已加载配置，恢复了 {restored_count} 项设置")

            # 重要：更新占位符匹配界面的表格显示
            matched_text_additions = self._update_match_table_display(main_window, text_additions)

            # 重要：只同步匹配成功的文本增加规则到PPT生成器
            if matched_text_additions:
                self._sync_to_ppt_generator(main_window, matched_text_additions)
            else:
                logger.info("没有匹配的文本增加规则，跳过PPT生成器同步")

            logger.info(f"GUI状态恢复完成，恢复了 {restored_count} 项设置")
            return True

        except Exception as e:
            logger.error(f"恢复GUI状态失败: {str(e)}")
            return False

    def _restore_dundian_settings(self, main_window, dundian_settings: Dict[str, Any]):
        """恢复递交趸期设置"""
        try:
            # 恢复复选框状态
            if hasattr(main_window, 'dundian_checkbox'):
                enabled = dundian_settings.get('enabled', False)
                main_window.dundian_checkbox.setChecked(enabled)
                logger.info(f"恢复递交趸期设置: {enabled}")

            # 恢复默认类型
            if hasattr(main_window, 'dundian_type_combo'):
                default_type = dundian_settings.get('default_type', '期缴')
                # 查找并设置对应的选项
                for i in range(main_window.dundian_type_combo.count()):
                    if main_window.dundian_type_combo.itemText(i) == default_type:
                        main_window.dundian_type_combo.setCurrentIndex(i)
                        logger.info(f"恢复默认递交类型: {default_type}")
                        break

        except Exception as e:
            logger.error(f"恢复递交趸期设置失败: {str(e)}")

    def get_config_preview(self, file_path: str) -> Optional[str]:
        """获取配置文件预览信息"""
        try:
            if not os.path.exists(file_path):
                return "配置文件不存在"

            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)

            if not self._validate_config_format(config_data):
                return "配置文件格式不正确"

            gui_settings = config_data.get('gui_settings', {})
            preview_info = []

            # 文本增加规则
            text_additions = gui_settings.get('text_additions', {})
            if text_additions:
                preview_info.append(f"文本增加规则: {len(text_additions)} 条")
                for placeholder, rule in text_additions.items():
                    if isinstance(rule, dict):
                        prefix = rule.get('prefix', '')
                        suffix = rule.get('suffix', '')
                        if prefix or suffix:
                            preview_info.append(f"  {placeholder}: 前缀'{prefix}' 后缀'{suffix}'")
            else:
                preview_info.append("文本增加规则: 无")

            # 递交趸期映射
            dundian_mappings = gui_settings.get('dundian_mappings', {})
            if dundian_mappings:
                preview_info.append(f"递交趸期设置: {len(dundian_mappings)} 条映射")
                for placeholder, data_column in dundian_mappings.items():
                    preview_info.append(f"  {placeholder} -> {data_column}")
            else:
                preview_info.append("递交趸期设置: 未配置")

            # 承保趸期映射
            chengbao_mappings = gui_settings.get('chengbao_mappings', {})
            if chengbao_mappings:
                preview_info.append(f"承保趸期设置: {len(chengbao_mappings)} 条映射")
                for placeholder, data_column in chengbao_mappings.items():
                    preview_info.append(f"  {placeholder} -> {data_column}")
            else:
                preview_info.append("承保趸期设置: 未配置")

            return '\n'.join(preview_info) if preview_info else "没有找到有效设置"

        except Exception as e:
            logger.error(f"获取配置预览失败: {str(e)}")
            return f"读取配置文件失败: {str(e)}"

    def _update_match_table_display(self, main_window, text_additions: Dict[str, Any]) -> Dict[str, Any]:
        """更新占位符匹配界面的表格显示
        返回匹配成功的文本增加规则字典
        """
        matched_rules = {}

        try:
            if not hasattr(main_window, 'match_table'):
                return matched_rules

            # 遍历表格的每一行，更新有文本增加规则的行
            for row in range(main_window.match_table.rowCount()):
                placeholder_item = main_window.match_table.item(row, 0)
                if not placeholder_item:
                    continue

                placeholder = placeholder_item.text()

                # 检查这个占位符是否有文本增加规则
                if placeholder in text_additions:
                    rule = text_additions[placeholder]
                    prefix = rule.get('prefix', '')
                    suffix = rule.get('suffix', '')

                    # 首先记录匹配成功的规则（不管是否更新表格）
                    matched_rules[placeholder] = rule

                    # 获取当前表格中显示的匹配列
                    current_item = main_window.match_table.item(row, 1)
                    if current_item:
                        current_text = current_item.text()

                        # 如果已经包含了文本增加信息，跳过（避免重复添加）
                        if "前缀:" in current_text or "后缀:" in current_text:
                            continue

                        # 获取原始的Excel列名
                        original_column = current_text.replace('[', '').replace(']', '').strip()
                        if not original_column:
                            original_column = "自定义"

                        # 构建显示文本
                        display_text = ""
                        if prefix:
                            display_text += f"前缀:{prefix} "
                        if suffix:
                            display_text += f"后缀:{suffix}"

                        # 更新表格显示
                        new_display = f"[{original_column}] {display_text}"
                        main_window.match_table.setItem(row, 1, QTableWidgetItem(new_display))

                        logger.info(f"更新表格显示: {placeholder} -> {new_display}")

        except Exception as e:
            logger.error(f"更新表格显示失败: {str(e)}")

        return matched_rules

    def _sync_to_ppt_generator(self, main_window, text_additions: Dict[str, Any]):
        """同步文本增加规则到PPT生成器"""
        try:
            # 同步到ExactMatcher
            if hasattr(main_window, 'exact_matcher') and main_window.exact_matcher:
                for placeholder, rule in text_additions.items():
                    prefix = rule.get('prefix', '')
                    suffix = rule.get('suffix', '')
                    main_window.exact_matcher.set_text_addition_rule(placeholder, prefix, suffix)
                    logger.info(f"同步到ExactMatcher: {placeholder} -> 前缀:'{prefix}' 后缀:'{suffix}'")

            # 同步到工作线程
            if hasattr(main_window, 'worker_thread') and main_window.worker_thread:
                # 更新工作线程的文本增加规则
                main_window.worker_thread.set_text_addition_rules(text_additions)
                logger.info(f"同步到工作线程: {len(text_additions)} 条文本增加规则")

            # 同步到PPT生成器
            if (hasattr(main_window, 'worker_thread') and
                hasattr(main_window.worker_thread, 'ppt_generator') and
                main_window.worker_thread.ppt_generator):

                for placeholder, rule in text_additions.items():
                    prefix = rule.get('prefix', '')
                    suffix = rule.get('suffix', '')
                    main_window.worker_thread.ppt_generator.exact_matcher.set_text_addition_rule(
                        placeholder, prefix, suffix
                    )

                logger.info(f"同步到PPT生成器完成: {len(text_additions)} 条规则")

        except Exception as e:
            logger.error(f"同步到PPT生成器失败: {str(e)}")

    def _placeholder_exists(self, main_window, placeholder: str) -> bool:
        """检查占位符是否在当前模板中存在"""
        try:
            if not hasattr(main_window, 'match_table'):
                logger.warning(f"占位符检查失败: match_table不存在")
                return False

            # 遍历匹配表格，查找占位符
            for row in range(main_window.match_table.rowCount()):
                placeholder_item = main_window.match_table.item(row, 0)
                if placeholder_item:
                    try:
                        item_text = placeholder_item.text()
                        if item_text == placeholder:
                            logger.info(f"找到匹配的占位符: {placeholder}")
                            return True
                    except Exception as e:
                        logger.error(f"获取占位符文本失败: {str(e)}, row={row}")
                        continue

            logger.info(f"未找到匹配的占位符: {placeholder}")
            return False

        except Exception as e:
            logger.error(f"检查占位符存在性失败: {str(e)}")
            return False

    def _update_match_table_for_dundian(self, main_window, placeholder: str, data_column: str):
        """更新匹配表格中递交趸期数据的显示"""
        try:
            if not hasattr(main_window, 'match_table'):
                return

            # 遍历匹配表格，找到对应的占位符行
            for row in range(main_window.match_table.rowCount()):
                placeholder_item = main_window.match_table.item(row, 0)
                if placeholder_item and placeholder_item.text() == placeholder:
                    # 更新数据列显示
                    match_item = QTableWidgetItem(data_column)
                    main_window.match_table.setItem(row, 1, match_item)

                    # 更新匹配规则
                    if hasattr(main_window, 'worker_thread') and main_window.worker_thread:
                        current_rules = main_window.worker_thread.matching_rules.copy()
                        current_rules[placeholder] = data_column
                        main_window.worker_thread.matching_rules = current_rules

                    logger.info(f"更新匹配表格: {placeholder} -> {data_column}")
                    return

        except Exception as e:
            logger.error(f"更新匹配表格失败: {str(e)}")

    def _update_match_table_for_chengbao(self, main_window, placeholder: str, data_column: str):
        """更新匹配表格中承保趸期数据的显示"""
        try:
            if not hasattr(main_window, 'match_table'):
                return

            # 遍历匹配表格，找到对应的占位符行
            for row in range(main_window.match_table.rowCount()):
                placeholder_item = main_window.match_table.item(row, 0)
                if placeholder_item and placeholder_item.text() == placeholder:
                    # 更新数据列显示
                    match_item = QTableWidgetItem(data_column)
                    main_window.match_table.setItem(row, 1, match_item)

                    # 更新匹配规则到工作线程和PPT生成器
                    if hasattr(main_window, 'worker_thread') and main_window.worker_thread:
                        # 获取PPT生成器（延迟初始化）
                        if hasattr(main_window.worker_thread, 'get_ppt_generator'):
                            ppt_generator = main_window.worker_thread.get_ppt_generator()
                            if ppt_generator:
                                # 添加匹配规则
                                ppt_generator.add_matching_rule(placeholder, data_column)
                                logger.info(f"更新PPT生成器匹配规则: {placeholder} -> {data_column}")

                        # 更新工作线程的匹配规则
                        current_rules = main_window.worker_thread.matching_rules.copy()
                        current_rules[placeholder] = data_column
                        main_window.worker_thread.matching_rules = current_rules
                        logger.info(f"更新工作线程匹配规则: {placeholder} -> {data_column}")

                    logger.info(f"更新匹配表格: {placeholder} -> {data_column}")
                    return

        except Exception as e:
            logger.error(f"更新承保趸期匹配表格失败: {str(e)}")

    def _trigger_chengbao_calculation(self, main_window):
        """自动触发承保趸期数据计算"""
        # 检查是否已经触发过（避免重复触发）
        if hasattr(main_window, '_chengbao_calculation_triggered') and main_window._chengbao_calculation_triggered:
            logger.info("承保趸期数据计算已触发过，跳过重复计算")
            return

        # 添加调用计数
        if not hasattr(self, '_calculation_call_count'):
            self._calculation_call_count = 0
        self._calculation_call_count += 1

        logger.info(f"承保趸期数据自动计算触发 (第 {self._calculation_call_count} 次)")

        # 标记为已触发
        main_window._chengbao_calculation_triggered = True

        try:
            # 检查是否有Excel数据文件路径
            if not hasattr(main_window, 'data_path_edit') or not main_window.data_path_edit.text():
                logger.info("未找到Excel数据文件，跳过承保趸期数据计算")
                return

            excel_file = main_window.data_path_edit.text()
            if not os.path.exists(excel_file):
                logger.info(f"Excel文件不存在: {excel_file}，跳过承保趸期数据计算")
                return

            # 加载Excel数据
            try:
                from src.data_reader import DataReader
            except ImportError:
                from data_reader import DataReader
            data_reader = DataReader()

            # 加载Excel文件并自动处理承保趸期数据
            # 重要：不传递parent_widget，避免DataReader弹出对话框
            # 配置管理器自己处理对话框显示
            if not data_reader.load_excel(excel_file, use_thousand_separator=True, parent_widget=None):
                logger.warning("加载Excel数据失败，跳过承保趸期数据计算")
                return

            # 获取承保趸期数据列
            chengbao_data_column = data_reader.get_column("承保趸期数据")

            if not chengbao_data_column:
                logger.info("未找到承保趸期数据列，跳过计算")
                return

            # 确保PPT生成器包含承保趸期数据列
            if hasattr(main_window, 'worker_thread') and main_window.worker_thread:
                if hasattr(main_window.worker_thread, 'get_ppt_generator'):
                    try:
                        ppt_generator = main_window.worker_thread.get_ppt_generator()
                        if ppt_generator:
                            # 直接使用DataReader已计算的数据，避免重复加载和重复计算
                            logger.info("为PPT生成器同步承保趸期数据...")
                            ppt_generator.data = data_reader.data.copy()
                            ppt_generator.data_loaded = True

                            # ⚠️ 关键：同时同步到ppt_generator的data_reader实例
                            # 这样在PPTWorkerThread中检查时，数据才一致
                            if hasattr(ppt_generator, 'data_reader') and ppt_generator.data_reader:
                                logger.info("为PPT生成器的data_reader同步承保趸期数据...")
                                ppt_generator.data_reader.data = data_reader.data.copy()
                                logger.info("PPT生成器data_reader数据同步完成")

                            logger.info("PPT生成器数据同步完成")
                    except Exception as e:
                        logger.warning(f"PPT生成器数据同步警告: {str(e)}")

            # 检查是否需要用户输入
            user_input_needed = False
            user_input_rows = []

            for i, value in enumerate(chengbao_data_column):
                # 修复：只检测空值行（R=1的情况，需要用户输入）
                # 不应该检测包含"年交SFYP"的行（这些是自动计算的）
                if not value or (isinstance(value, str) and value.strip() == ""):
                    user_input_needed = True
                    user_input_rows.append(i)

            # 如果需要用户输入，弹出对话框
            if user_input_needed and hasattr(main_window, 'show_chengbao_input_dialog'):
                logger.info(f"检测到 {len(user_input_rows)} 行需要用户输入，弹出承保趸期数据输入对话框...")
                # 使用主窗口的方法显示对话框
                main_window.show_chengbao_input_dialog(data_reader, user_input_rows)
            else:
                logger.info("承保趸期数据计算完成，无需用户输入")

            logger.info("承保趸期数据自动计算完成")

        except Exception as e:
            logger.error(f"承保趸期数据自动计算失败: {str(e)}")
            if hasattr(main_window, 'log_text') and main_window.log_text:
                main_window.log_text.append(f"[配置管理] 承保趸期数据自动计算失败: {str(e)}")