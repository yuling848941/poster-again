"""
简化的配置管理对话框
直接基于GUI状态进行配置保存和加载
"""

import os
import logging
from typing import Optional
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTextEdit, QGroupBox, QListWidget, QListWidgetItem,
    QMessageBox, QFileDialog
)
from PySide6.QtCore import Qt

try:
    from .simple_config_manager import SimpleConfigManager
except ImportError:
    from simple_config_manager import SimpleConfigManager

logger = logging.getLogger(__name__)


class SimpleConfigDialog(QDialog):
    """简化的配置管理对话框"""

    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.config_manager = SimpleConfigManager()

        self.setWindowTitle("配置管理")
        self.setMinimumSize(600, 500)
        self.setup_ui()

    def setup_ui(self):
        """设置用户界面"""
        layout = QVBoxLayout()

        # 操作按钮组
        actions_group = QGroupBox("配置操作")
        actions_layout = QHBoxLayout()

        self.save_btn = QPushButton("保存当前配置")
        self.save_btn.clicked.connect(self.save_config)
        self.save_btn.setMinimumHeight(40)

        self.load_btn = QPushButton("加载配置文件")
        self.load_btn.clicked.connect(self.load_config)
        self.load_btn.setMinimumHeight(40)

        actions_layout.addWidget(self.save_btn)
        actions_layout.addWidget(self.load_btn)
        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)

        # 当前状态显示
        current_group = QGroupBox("当前设置状态")
        current_layout = QVBoxLayout()

        self.current_info = QTextEdit()
        self.current_info.setMaximumHeight(200)
        self.current_info.setReadOnly(True)
        self._update_current_info()
        current_layout.addWidget(self.current_info)
        current_group.setLayout(current_layout)
        layout.addWidget(current_group)

        # 最近配置文件
        recent_group = QGroupBox("最近使用的配置")
        recent_layout = QVBoxLayout()

        self.recent_list = QListWidget()
        self.recent_list.setMaximumHeight(150)
        self.recent_list.itemDoubleClicked.connect(self.load_recent_config)
        recent_layout.addWidget(self.recent_list)
        recent_group.setLayout(recent_layout)
        layout.addWidget(recent_group)

        # 按钮
        button_layout = QHBoxLayout()
        self.refresh_btn = QPushButton("刷新列表")
        self.refresh_btn.clicked.connect(self.refresh_recent_files)
        self.close_btn = QPushButton("关闭")
        self.close_btn.clicked.connect(self.accept)

        button_layout.addWidget(self.refresh_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.close_btn)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def _update_current_info(self):
        """更新当前状态信息"""
        try:
            info_text = "=== 当前设置状态 ===\n\n"

            # 统计文本增加规则
            if hasattr(self.main_window, 'text_addition_rules'):
                text_rules = self.main_window.text_addition_rules
                if text_rules:
                    info_text += f"文本增加规则 ({len(text_rules)} 条):\n"
                    for placeholder, rule in text_rules.items():
                        if isinstance(rule, dict):
                            prefix = rule.get('prefix', '')
                            suffix = rule.get('suffix', '')
                            if prefix or suffix:
                                info_text += f"  ✓ {placeholder}: 前缀='{prefix}' 后缀='{suffix}'\n"
                else:
                    info_text += "文本增加规则: 无\n"
            else:
                info_text += "文本增加规则: 无\n"

            info_text += "\n"

            # 统计递交趸期设置
            dundian_info = self._get_dundian_info()
            info_text += f"递交趸期设置: {dundian_info}\n\n"

            # 统计承保趸期映射
            info_text += "承保趸期设置:\n"
            if hasattr(self.main_window, 'chengbao_mappings'):
                chengbao_mappings = self.main_window.chengbao_mappings
                if chengbao_mappings:
                    info_text += f"  已配置 {len(chengbao_mappings)} 条映射关系:\n"
                    for placeholder, data_column in chengbao_mappings.items():
                        info_text += f"    ✓ {placeholder} -> {data_column}\n"
                else:
                    info_text += "  未配置映射关系\n"
            else:
                info_text += "  未配置映射关系\n"

            info_text += "\n提示：\n"
            info_text += "• 文本增加规则：为占位符添加前缀/后缀\n"
            info_text += "• 递交趸期设置：设置递交类型和默认选项\n"
            info_text += "• 承保趸期设置：记录占位符与承保趸期数据的映射关系\n"
            info_text += "• 配置文件会保存您在GUI界面所做的所有设置"

            self.current_info.setText(info_text)

        except Exception as e:
            self.current_info.setText(f"获取当前状态失败: {str(e)}")

    def _get_dundian_info(self):
        """获取递交趸期信息"""
        try:
            info = ""
            if hasattr(self.main_window, 'dundian_checkbox'):
                enabled = self.main_window.dundian_checkbox.isChecked()
                info += f"状态: {'启用' if enabled else '禁用'}"

                if hasattr(self.main_window, 'dundian_type_combo'):
                    default_type = self.main_window.dundian_type_combo.currentText()
                    info += f"\n默认类型: {default_type}"
            else:
                info = "未配置"

            return info
        except:
            return "获取失败"

    def save_config(self):
        """保存当前配置"""
        try:
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "保存配置文件",
                os.path.join(os.path.expanduser('~'), "gui_config.pptcfg"),
                "配置文件 (*.pptcfg);;所有文件 (*.*)"
            )

            if not file_path:
                return

            success = self.config_manager.save_gui_config(self.main_window, file_path)

            if success:
                QMessageBox.information(
                    self, "保存成功",
                    f"配置已成功保存到:\n{file_path}"
                )
                self._update_current_info()
                self.refresh_recent_files()
            else:
                QMessageBox.warning(self, "保存失败", "配置保存失败，请检查设置内容。")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存配置时发生错误:\n{str(e)}")

    def load_config(self):
        """加载配置文件"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "加载配置文件",
                os.path.expanduser('~'),
                "配置文件 (*.pptcfg);;所有文件 (*.*)"
            )

            if not file_path:
                return

            success = self.config_manager.load_gui_config(self.main_window, file_path)

            if success:
                QMessageBox.information(
                    self, "加载成功",
                    f"配置已成功加载:\n{file_path}"
                )
                self._update_current_info()
                self.refresh_recent_files()
            else:
                QMessageBox.warning(self, "加载失败", "配置加载失败，请检查文件格式。")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载配置时发生错误:\n{str(e)}")

    def load_recent_config(self, item: QListWidgetItem):
        """加载最近使用的配置文件"""
        file_path = item.data(Qt.UserRole)
        if file_path and os.path.exists(file_path):
            self.config_manager.load_gui_config(self.main_window, file_path)
            self._update_current_info()

    def refresh_recent_files(self):
        """刷新最近配置文件列表"""
        try:
            self.recent_list.clear()

            # 获取当前文件所在目录
            current_dir = None
            if hasattr(self.main_window, 'data_path') and self.main_window.data_path:
                current_dir = os.path.dirname(self.main_window.data_path)
            elif hasattr(self.main_window, 'template_path') and self.main_window.template_path:
                current_dir = os.path.dirname(self.main_window.template_path)

            if current_dir and os.path.exists(current_dir):
                # 搜索目录中的.pptcfg文件
                for file_name in os.listdir(current_dir):
                    if file_name.endswith('.pptcfg'):
                        file_path = os.path.join(current_dir, file_name)

                        # 获取配置预览
                        preview = self.config_manager.get_config_preview(file_path)

                        # 创建列表项
                        item = QListWidgetItem(f"{file_name} - {preview[:50]}...")
                        item.setData(Qt.UserRole, file_path)
                        item.setToolTip(f"{file_name}\n{preview}")
                        self.recent_list.addItem(item)

            if self.recent_list.count() == 0:
                item = QListWidgetItem("暂无最近使用的配置文件")
                # 使用setFlags来禁用项目，而不是setEnabled
                item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                self.recent_list.addItem(item)

        except Exception as e:
            logger.error(f"刷新最近配置文件列表失败: {str(e)}")
            item = QListWidgetItem("刷新列表失败")
            # 使用setFlags来禁用项目，而不是setEnabled
            item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
            self.recent_list.addItem(item)


def show_simple_config_dialog(main_window):
    """显示简化配置对话框"""
    dialog = SimpleConfigDialog(main_window)
    dialog.exec()