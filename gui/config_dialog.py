"""
配置保存加载对话框
提供用户友好的配置管理界面
"""

import os
import json
from typing import Dict, Any, List, Optional, Callable
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QTextEdit, QProgressBar,
    QGroupBox, QRadioButton, QButtonGroup,
    QMessageBox, QFileDialog, QCheckBox,
    QListWidget, QListWidgetItem, QSplitter
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QColor

try:
    from src.config import ConfigManager
    from src.exact_matcher import ExactMatcher
except ImportError:
    from config_manager import ConfigManager
    from exact_matcher import ExactMatcher


class ConfigLoadingThread(QThread):
    """配置加载线程，避免界面冻结"""
    progress_updated = Signal(int)
    status_updated = Signal(str)
    result_ready = Signal(dict)
    error_occurred = Signal(str)

    def __init__(self, config_manager: ConfigManager, file_path: str):
        super().__init__()
        self.config_manager = config_manager
        self.file_path = file_path

    def run(self):
        try:
            self.status_updated.emit("正在加载配置文件...")
            self.progress_updated.emit(20)

            # 加载配置
            load_result = self.config_manager.load_placeholder_config(self.file_path)
            self.progress_updated.emit(60)

            if not load_result["success"]:
                self.error_occurred.emit(load_result.get("error", "加载失败"))
                return

            self.status_updated.emit("配置文件加载成功")
            self.progress_updated.emit(80)

            # 发送结果
            self.result_ready.emit({
                "success": True,
                "config_data": load_result["config_data"]
            })

            self.progress_updated.emit(100)

        except Exception as e:
            self.error_occurred.emit(f"加载过程中发生错误: {str(e)}")


class ConfigCompatibilityDialog(QDialog):
    """配置兼容性检查结果对话框"""

    def __init__(self, compatibility_result: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.compatibility_result = compatibility_result
        self.setWindowTitle("配置兼容性检查")
        self.setMinimumSize(600, 500)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        # 总体兼容性
        compat_group = QGroupBox("兼容性分析结果")
        compat_layout = QVBoxLayout()

        overall_rate = self.compatibility_result.get('overall_compatibility_rate', 0)
        placeholder_rate = self.compatibility_result.get('placeholder_compatibility_rate', 0)
        data_rate = self.compatibility_result.get('data_compatibility_rate', 0)

        # 兼容性进度条
        compat_layout.addWidget(QLabel("整体兼容性:"))
        self.overall_progress = QProgressBar()
        self.overall_progress.setRange(0, 100)
        self.overall_progress.setValue(int(overall_rate * 100))
        self._set_progress_color(self.overall_progress, overall_rate)
        compat_layout.addWidget(self.overall_progress)

        compat_layout.addWidget(QLabel(f"占位符兼容性: {placeholder_rate:.1%}"))
        compat_layout.addWidget(QLabel(f"数据字段兼容性: {data_rate:.1%}"))

        # 统计信息
        stats = self.compatibility_result.get('statistics', {})
        stats_text = f"""
总规则数: {stats.get('total_rules', 0)}
兼容规则数: {stats.get('compatible_rules_count', 0)}
不兼容规则数: {stats.get('incompatible_rules_count', 0)}
        """.strip()
        compat_layout.addWidget(QLabel("统计信息:"))
        compat_layout.addWidget(QLabel(stats_text))

        compat_group.setLayout(compat_layout)
        layout.addWidget(compat_group)

        # 详细规则列表
        details_group = QGroupBox("规则详情")
        details_layout = QVBoxLayout()

        # 兼容的规则
        compatible_rules = self.compatibility_result.get('compatible_rules', [])
        if compatible_rules:
            compatible_label = QLabel(f"兼容的规则 ({len(compatible_rules)} 条):")
            compatible_label.setStyleSheet("font-weight: bold; color: green;")
            details_layout.addWidget(compatible_label)

            compatible_list = QListWidget()
            for rule in compatible_rules:
                item_text = f"{rule['placeholder']} → {rule['column']}"
                if 'text_addition' in rule:
                    ta = rule['text_addition']
                    if ta.get('prefix') or ta.get('suffix'):
                        item_text += f" (前缀:'{ta.get('prefix', '')}' 后缀:'{ta.get('suffix', '')}')"
                compatible_list.addItem(item_text)
            details_layout.addWidget(compatible_list)

        # 不兼容的规则
        incompatible_rules = self.compatibility_result.get('incompatible_rules', [])
        if incompatible_rules:
            incompatible_label = QLabel(f"不兼容的规则 ({len(incompatible_rules)} 条):")
            incompatible_label.setStyleSheet("font-weight: bold; color: red;")
            details_layout.addWidget(incompatible_label)

            incompatible_list = QListWidget()
            for item in incompatible_rules:
                rule = item['rule']
                reason = item['reason']
                item_text = f"{rule['placeholder']} → {rule['column']} ({reason})"
                incompatible_list.addItem(item_text)
            details_layout.addWidget(incompatible_list)

        details_group.setLayout(details_layout)
        layout.addWidget(details_group)

        # 按钮
        button_layout = QHBoxLayout()

        self.load_compatible_btn = QPushButton("仅加载兼容规则")
        self.load_compatible_btn.clicked.connect(self.accept)
        if not compatible_rules:
            self.load_compatible_btn.setEnabled(False)

        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)

        button_layout.addWidget(self.load_compatible_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def _set_progress_color(self, progress_bar: QProgressBar, rate: float):
        """根据兼容性设置进度条颜色"""
        if rate >= 0.8:
            style = "QProgressBar::chunk { background-color: #4CAF50; }"  # 绿色
        elif rate >= 0.5:
            style = "QProgressBar::chunk { background-color: #FF9800; }"  # 橙色
        else:
            style = "QProgressBar::chunk { background-color: #F44336; }"  # 红色
        progress_bar.setStyleSheet(style)


class ConfigSaveLoadDialog(QDialog):
    """配置保存加载主对话框"""

    def __init__(self, exact_matcher: ExactMatcher, template_file: str = "", data_file: str = "", parent=None):
        super().__init__(parent)
        self.exact_matcher = exact_matcher
        # 使用全局配置管理器实例
        self.config_manager = ConfigManager("config/config.yaml")
        print(f"[DEBUG] 配置管理器创建，配置文件: config/config.yaml")  # 调试输出
        self.template_file = template_file
        self.data_file = data_file

        self.setWindowTitle("配置管理")
        self.setMinimumSize(500, 400)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        # 操作按钮组
        actions_group = QGroupBox("配置操作")
        actions_layout = QGridLayout()

        self.save_btn = QPushButton("保存当前配置")
        self.save_btn.clicked.connect(self.save_config)
        self.save_btn.setMinimumHeight(40)

        self.load_btn = QPushButton("加载配置文件")
        self.load_btn.clicked.connect(self.load_config)
        self.load_btn.setMinimumHeight(40)

        actions_layout.addWidget(self.save_btn, 0, 0)
        actions_layout.addWidget(self.load_btn, 0, 1)
        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)

        # 当前配置信息
        info_group = QGroupBox("当前配置信息")
        info_layout = QVBoxLayout()

        self.info_text = QTextEdit()
        self.info_text.setMaximumHeight(150)
        self.info_text.setReadOnly(True)
        self._update_current_info()
        info_layout.addWidget(self.info_text)
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # 最近使用的配置文件
        recent_group = QGroupBox("最近配置")
        recent_layout = QVBoxLayout()

        self.recent_list = QListWidget()
        self.recent_list.setMaximumHeight(120)
        self.recent_list.itemDoubleClicked.connect(self.load_recent_config)
        recent_layout.addWidget(self.recent_list)

        # 刷新最近文件列表
        self._refresh_recent_files()

        recent_group.setLayout(recent_layout)
        layout.addWidget(recent_group)

        # 按钮
        button_layout = QHBoxLayout()
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self._refresh_recent_files)
        self.close_btn = QPushButton("关闭")
        self.close_btn.clicked.connect(self.accept)

        button_layout.addWidget(self.refresh_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def _update_current_info(self):
        """更新当前配置信息显示"""
        try:
            # 获取当前匹配规则
            matching_rules = self.exact_matcher.get_matching_rules()
            text_rules = self.exact_matcher.get_all_text_addition_rules()

            info_text = f"=== 当前配置状态 ===\n"
            info_text += f"匹配规则数: {len(matching_rules)}\n"
            info_text += f"文本增加规则数: {len(text_rules)}\n\n"

            if matching_rules:
                info_text += "匹配规则列表:\n"
                for placeholder, column in matching_rules.items():
                    text_rule = text_rules.get(placeholder, {})
                    prefix = text_rule.get('prefix', '')
                    suffix = text_rule.get('suffix', '')
                    text_info = f" (前缀:'{prefix}' 后缀:'{suffix}')" if prefix or suffix else ""
                    info_text += f"  ✓ {placeholder} → {column}{text_info}\n"
            else:
                info_text += "⚠ 没有找到匹配规则\n"

            if text_rules:
                info_text += "\n文本增加规则:\n"
                for placeholder, rule in text_rules.items():
                    prefix = rule.get('prefix', '')
                    suffix = rule.get('suffix', '')
                    if prefix or suffix:
                        info_text += f"  ✓ {placeholder}: 前缀='{prefix}' 后缀='{suffix}'\n"

            # 获取模板和数据信息
            template_info = self.exact_matcher.get_template_info(self.template_file)
            data_info = self.exact_matcher.get_data_info(self.data_file)

            info_text += f"\n=== 文件信息 ===\n"
            info_text += f"模板文件: {template_info.get('file_name', '未选择')}\n"
            info_text += f"模板占位符数: {template_info.get('placeholder_count', 0)}\n"
            if template_info.get('placeholders'):
                info_text += f"占位符列表: {', '.join(template_info['placeholders'])}\n"

            info_text += f"\n数据文件: {data_info.get('file_name', '未选择')}\n"
            info_text += f"数据字段数: {data_info.get('column_count', 0)}\n"
            if data_info.get('columns'):
                info_text += f"数据列: {', '.join(data_info['columns'])}\n"

            self.info_text.setText(info_text)

        except Exception as e:
            self.info_text.setText(f"获取配置信息失败: {str(e)}")

    def save_config(self):
        """保存当前配置"""
        try:
            # 导出当前配置
            exported_rules = self.exact_matcher.export_matching_config()
            text_rules = self.exact_matcher.get_all_text_addition_rules()

            # 详细检查配置状态
            if not exported_rules and not text_rules:
                QMessageBox.warning(self, "保存配置", "当前没有有效的配置可以保存！\n请先进行自动匹配或添加文本增加规则。")
                return

            # 如果只有文本增加规则，询问用户
            if not exported_rules and text_rules:
                reply = QMessageBox.question(
                    self, "保存配置",
                    f"当前只有文本增加规则（{len(text_rules)}条），没有占位符匹配规则。\n是否只保存文本增加规则？",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return

            # 生成保存规则（包含匹配规则和文本增加规则的占位符）
            save_rules = exported_rules.copy()

            # 如果有文本增加规则但没有对应的匹配规则，创建占位符映射
            if text_rules:
                for placeholder, rule in text_rules.items():
                    # 检查这个占位符是否已经在匹配规则中
                    already_matched = any(r['placeholder'] == placeholder for r in save_rules)
                    if not already_matched:
                        # 创建一个空列名的匹配规则，只保存文本增加信息
                        save_rules.append({
                            'placeholder': placeholder,
                            'column': '',  # 空列名表示只有文本增加规则
                            'text_addition': rule
                        })

            # 选择保存路径 - 记住上次使用的目录
            last_save_dir = self.config_manager.get("paths.last_config_save_dir", "")
            if not last_save_dir:
                last_save_dir = os.path.expanduser('~')

            print(f"[DEBUG] 读取到的保存目录: {last_save_dir}")  # 调试输出

            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存配置文件",
                last_save_dir,
                "配置文件 (*.pptcfg);;所有文件 (*.*)"
            )

            if not file_path:
                return

            # 记住保存路径
            save_dir = os.path.dirname(file_path)
            print(f"[DEBUG] 保存新目录: {save_dir}")  # 调试输出
            self.config_manager.set("paths.last_config_save_dir", save_dir)

            # 保存配置到文件
            save_success = self.config_manager.save_config()
            print(f"[DEBUG] 配置保存结果: {save_success}")  # 调试输出

            if save_success:
                print("[DEBUG] 配置已成功保存到文件")  # 调试输出

            # 准备保存数据
            template_info = self.exact_matcher.get_template_info(self.template_file)
            data_info = self.exact_matcher.get_data_info(self.data_file)

            # 保存配置
            success = self.config_manager.save_placeholder_config(
                file_path, save_rules, template_info, data_info
            )

            if success:
                # 统计保存的配置类型
                matched_rules_count = len(exported_rules)
                text_only_rules_count = len(save_rules) - matched_rules_count

                success_msg = f"配置已成功保存到:\n{file_path}\n\n"
                success_msg += f"包含 {matched_rules_count} 条匹配规则"
                if text_only_rules_count > 0:
                    success_msg += f"\n包含 {text_only_rules_count} 条纯文本增加规则"

                QMessageBox.information(self, "保存成功", success_msg)
                self._update_current_info()
                self._refresh_recent_files()
            else:
                QMessageBox.critical(self, "保存失败", "配置保存失败，请检查文件权限和路径！")

        except Exception as e:
            QMessageBox.critical(self, "保存错误", f"保存配置时发生错误:\n{str(e)}")

    def load_config(self):
        """加载配置文件"""
        try:
            # 选择配置文件 - 记住上次使用的目录
            last_load_dir = self.config_manager.get("paths.last_config_load_dir", "")
            if not last_load_dir:
                last_load_dir = os.path.expanduser('~')

            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "加载配置文件",
                last_load_dir,
                "配置文件 (*.pptcfg);;所有文件 (*.*)"
            )

            if not file_path:
                return

            # 记住加载路径
            load_dir = os.path.dirname(file_path)
            self.config_manager.set("paths.last_config_load_dir", load_dir)
            self.config_manager.save_config()

            self._load_config_file(file_path)

        except Exception as e:
            QMessageBox.critical(self, "加载错误", f"加载配置时发生错误:\n{str(e)}")

    def load_recent_config(self, item: QListWidgetItem):
        """加载最近使用的配置文件"""
        file_path = item.data(Qt.UserRole)
        if file_path and os.path.exists(file_path):
            self._load_config_file(file_path)

    def _load_config_file(self, file_path: str):
        """加载指定的配置文件"""
        try:
            # 获取当前模板和数据信息用于兼容性检查
            current_placeholders = self.exact_matcher.template_placeholders
            current_columns = list(self.exact_matcher.data.columns) if self.exact_matcher.data is not None else []

            # 如果没有当前数据，跳过兼容性检查直接加载
            if not current_placeholders or not current_columns:
                self._direct_load_config(file_path)
                return

            # 加载配置文件
            load_result = self.config_manager.load_placeholder_config(file_path)
            if not load_result["success"]:
                QMessageBox.critical(self, "加载失败", load_result.get("error", "加载配置失败"))
                return

            config_data = load_result["config_data"]

            # 兼容性检查
            compatibility_result = self.config_manager.validate_placeholder_compatibility(
                config_data, current_placeholders, current_columns
            )

            # 显示兼容性检查结果
            compat_dialog = ConfigCompatibilityDialog(compatibility_result, self)
            if compat_dialog.exec() == QDialog.Accepted:
                # 用户选择加载兼容的规则
                compatible_rules = compatibility_result.get('compatible_rules', [])
                if compatible_rules:
                    self._apply_config_rules(compatible_rules)
                else:
                    QMessageBox.information(self, "加载结果", "没有兼容的规则可以加载")

        except Exception as e:
            QMessageBox.critical(self, "加载错误", f"处理配置文件时发生错误:\n{str(e)}")

    def _direct_load_config(self, file_path: str):
        """直接加载配置文件（跳过兼容性检查）"""
        load_result = self.config_manager.load_placeholder_config(file_path)
        if not load_result["success"]:
            QMessageBox.critical(self, "加载失败", load_result.get("error", "加载配置失败"))
            return

        config_data = load_result["config_data"]
        matching_rules = config_data.get("matching_rules", [])

        self._apply_config_rules(matching_rules)

    def _apply_config_rules(self, matching_rules: List[Dict[str, Any]]):
        """应用配置规则"""
        try:
            import_result = self.exact_matcher.import_matching_config(matching_rules)

            if import_result.get('success'):
                imported = import_result.get('imported_count', 0)
                skipped = import_result.get('skipped_count', 0)

                message = f"配置加载完成！\n\n成功导入: {imported} 条规则"
                if skipped > 0:
                    message += f"\n跳过: {skipped} 条规则"

                QMessageBox.information(self, "加载成功", message)
                self._update_current_info()
            else:
                QMessageBox.critical(self, "加载失败", import_result.get("error", "导入配置失败"))

        except Exception as e:
            QMessageBox.critical(self, "应用错误", f"应用配置规则时发生错误:\n{str(e)}")

    def _refresh_recent_files(self):
        """刷新最近使用的配置文件列表"""
        self.recent_list.clear()

        # 这里可以实现最近文件列表的逻辑
        # 暂时添加一个示例项
        if self.template_file and self.data_file:
            recent_dir = os.path.dirname(self.template_file) or os.path.dirname(self.data_file)
            if recent_dir and os.path.exists(recent_dir):
                for file in os.listdir(recent_dir):
                    if file.endswith('.pptcfg'):
                        file_path = os.path.join(recent_dir, file)
                        item = QListWidgetItem(file)
                        item.setData(Qt.UserRole, file_path)
                        item.setToolTip(f"路径: {file_path}")
                        self.recent_list.addItem(item)