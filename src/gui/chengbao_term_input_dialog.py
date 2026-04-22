"""
承保趸期数据批量输入对话框
用于用户一次性输入多个承保趸期数据值
"""

import logging
from typing import List, Dict, Optional
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QWidget, QScrollArea, QFrame, QMessageBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIntValidator

logger = logging.getLogger(__name__)


class ChengbaoTermInputDialog(QDialog):
    """承保趸期数据批量输入对话框"""

    def __init__(self, pending_rows: List[Dict], parent=None):
        """
        初始化对话框

        Args:
            pending_rows: 需要用户输入的行信息列表
                         格式: [{'row_index': int, 'policy_number': str}, ...]
            parent: 父窗口
        """
        super().__init__(parent)
        self.pending_rows = pending_rows
        self.input_values = {}  # 存储用户输入值，格式: {row_index: value}
        self.validation_errors = []

        self.setWindowTitle("承保趸期数据批量输入")
        self.setMinimumSize(600, 400)
        self.setModal(True)

        # 窗口大小根据输入行数调整
        if len(pending_rows) > 10:
            self.resize(700, 600)
        else:
            self.resize(600, 400)

        self.setup_ui()

    def setup_ui(self):
        """设置用户界面"""
        main_layout = QVBoxLayout(self)

        # 添加说明文字
        info_label = QLabel(
            f"以下 {len(self.pending_rows)} 行需要输入承保趸期数据：\n"
            "请在每个输入框中填入缴费年期（≥2的正整数）"
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666; font-size: 12px;")
        main_layout.addWidget(info_label)

        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)

        # 创建内容Widget
        content_widget = QWidget()
        scroll_layout = QVBoxLayout(content_widget)
        scroll_layout.setSpacing(10)

        # 为每一行创建输入组件
        for row_info in self.pending_rows:
            row_index = row_info['row_index']
            policy_number = row_info.get('policy_number', 'N/A')

            # 创建行Widget
            row_widget = self.create_row_widget(row_index, policy_number)
            scroll_layout.addWidget(row_widget)

        scroll_layout.addStretch()
        scroll_area.setWidget(content_widget)
        main_layout.addWidget(scroll_area)

        # 添加按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setMinimumHeight(40)
        self.cancel_btn.clicked.connect(self.reject)

        self.ok_btn = QPushButton("确定")
        self.ok_btn.setMinimumHeight(40)
        self.ok_btn.clicked.connect(self.on_ok_clicked)
        self.ok_btn.setDefault(True)

        button_layout.addWidget(self.cancel_btn)
        button_layout.addWidget(self.ok_btn)

        main_layout.addLayout(button_layout)

    def create_row_widget(self, row_index: int, policy_number: str) -> QWidget:
        """
        创建单行输入组件

        Args:
            row_index: 行索引
            policy_number: 保单号

        Returns:
            QWidget: 行组件
        """
        row_frame = QFrame()
        row_frame.setFrameShape(QFrame.StyledPanel)
        row_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 8px;
            }
        """)
        row_layout = QHBoxLayout(row_frame)
        row_layout.setContentsMargins(10, 8, 10, 8)

        # 行号标签
        row_label = QLabel(f"第 {row_index + 1} 行")
        row_label.setMinimumWidth(80)
        row_label.setStyleSheet("font-weight: bold; color: #495057;")
        row_layout.addWidget(row_label)

        # ✅ 新增：保单号显示框（只读但可复制）
        policy_field = QLineEdit(str(policy_number))
        policy_field.setReadOnly(True)  # 设置为只读，不可更改
        policy_field.setMinimumWidth(150)
        policy_field.setMaximumWidth(200)

        # 设置样式：看起来像不可编辑但可以选中和复制
        policy_field.setStyleSheet("""
            QLineEdit {
                color: #6c757d;
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 3px;
                padding: 4px 8px;
                selection-background-color: #007bff;
            }
            QLineEdit:focus {
                border-color: #007bff;
            }
        """)

        # ✅ 新增：提示文本，告诉用户可以复制
        policy_field.setToolTip("可复制此保单号进行查询")

        row_layout.addWidget(policy_field)

        # 缴费年期标签
        term_label = QLabel("缴费年期:")
        term_label.setMinimumWidth(80)
        term_label.setStyleSheet("color: #495057;")
        row_layout.addWidget(term_label)

        # 输入框
        input_field = QLineEdit()
        input_field.setPlaceholderText("输入数值")
        input_field.setMinimumWidth(100)
        input_field.setMaximumWidth(150)

        # 设置验证器：只接受≥2的正整数
        validator = QIntValidator(2, 999, self)
        input_field.setValidator(validator)

        # 连接文本变化信号
        input_field.textChanged.connect(
            lambda text, idx=row_index: self.on_text_changed(idx, text)
        )

        # 存储输入框引用
        setattr(self, f'input_field_{row_index}', input_field)
        row_layout.addWidget(input_field)

        # 年交SFYP标签
        unit_label = QLabel("年交SFYP")
        unit_label.setMinimumWidth(80)
        unit_label.setStyleSheet("color: #6c757d; font-style: italic;")
        row_layout.addWidget(unit_label)

        row_layout.addStretch()

        return row_frame

    def on_text_changed(self, row_index: int, text: str):
        """
        文本变化处理

        Args:
            row_index: 行索引
            text: 输入文本
        """
        if text and text.isdigit():
            self.input_values[row_index] = int(text)
        else:
            if row_index in self.input_values:
                del self.input_values[row_index]

    def on_ok_clicked(self):
        """确定按钮点击处理"""
        # 验证所有输入
        errors = []
        valid_count = 0

        for row_info in self.pending_rows:
            row_index = row_info['row_index']

            # 检查是否有输入
            if row_index not in self.input_values:
                errors.append(f"第 {row_index + 1} 行 (保单号: {row_info.get('policy_number', 'N/A')}) 未输入")
            else:
                value = self.input_values[row_index]
                if value < 2:
                    errors.append(f"第 {row_index + 1} 行 (保单号: {row_info.get('policy_number', 'N/A')}) 输入值必须≥2")
                else:
                    valid_count += 1

        # 如果有验证错误，显示错误信息
        if errors:
            error_message = "请修正以下问题：\n\n" + "\n".join(errors)
            QMessageBox.warning(self, "输入验证失败", error_message)
            return

        # 所有验证通过，关闭对话框
        logger.info(f"用户输入承保趸期数据完成，共 {valid_count} 行")
        self.accept()

    def get_input_values(self) -> Dict[int, int]:
        """
        获取所有输入值

        Returns:
            Dict[int, int]: 行索引到输入值的映射
        """
        return self.input_values.copy()

    def keyPressEvent(self, event):
        """处理按键事件"""
        # 按ESC键取消
        if event.key() == Qt.Key_Escape:
            self.reject()
        else:
            super().keyPressEvent(event)
