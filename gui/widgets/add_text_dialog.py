from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QGroupBox, QApplication
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont


class AddTextDialog(QDialog):
    """增加文本对话框，用于为占位符添加前后文本"""
    
    def __init__(self, placeholder_name="", parent=None):
        super().__init__(parent)
        self.placeholder_name = placeholder_name
        self.prefix_text = ""
        self.suffix_text = ""
        
        self.init_ui()
        
    def init_ui(self):
        """初始化UI界面"""
        self.setWindowTitle("增加文本")
        self.setMinimumWidth(400)
        self.setWindowModality(Qt.ApplicationModal)  # 设置为模态对话框
        
        # 创建主布局
        main_layout = QVBoxLayout(self)
        
        # 创建标题标签
        title_label = QLabel(f"为占位符「{self.placeholder_name}」增加文本")
        title_label.setFont(QFont("Arial", 12, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 创建文本输入组
        input_group = QGroupBox("文本设置")
        input_layout = QVBoxLayout()
        
        # 前面文本输入
        prefix_layout = QHBoxLayout()
        prefix_label = QLabel("数据前面文本:")
        self.prefix_edit = QLineEdit()
        self.prefix_edit.setPlaceholderText("输入要添加到数据前面的文本（可选）")
        prefix_layout.addWidget(prefix_label)
        prefix_layout.addWidget(self.prefix_edit)
        input_layout.addLayout(prefix_layout)
        
        # 后面文本输入
        suffix_layout = QHBoxLayout()
        suffix_label = QLabel("数据后面文本:")
        self.suffix_edit = QLineEdit()
        self.suffix_edit.setPlaceholderText("输入要添加到数据后面的文本（可选）")
        suffix_layout.addWidget(suffix_label)
        suffix_layout.addWidget(self.suffix_edit)
        input_layout.addLayout(suffix_layout)
        
        # 示例预览
        preview_layout = QHBoxLayout()
        preview_label = QLabel("示例预览:")
        self.preview_label = QLabel("数据")
        preview_layout.addWidget(preview_label)
        preview_layout.addWidget(self.preview_label)
        input_layout.addLayout(preview_layout)
        
        input_group.setLayout(input_layout)
        main_layout.addWidget(input_group)
        
        # 创建按钮组
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        # 取消按钮
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        # 确认按钮
        confirm_btn = QPushButton("确认")
        confirm_btn.clicked.connect(self.accept_changes)
        confirm_btn.setDefault(True)
        button_layout.addWidget(confirm_btn)
        
        main_layout.addLayout(button_layout)
        
        # 连接文本变化信号，更新预览
        self.prefix_edit.textChanged.connect(self.update_preview)
        self.suffix_edit.textChanged.connect(self.update_preview)
        
    def update_preview(self):
        """更新预览标签"""
        prefix = self.prefix_edit.text()
        suffix = self.suffix_edit.text()
        self.preview_label.setText(f"{prefix}数据{suffix}")
        
    def accept_changes(self):
        """确认更改，保存文本设置"""
        self.prefix_text = self.prefix_edit.text()
        self.suffix_text = self.suffix_edit.text()
        self.accept()
        
    def get_text_values(self):
        """获取用户输入的前缀和后缀文本"""
        return self.prefix_edit.text().strip(), self.suffix_edit.text().strip()
        
    def set_text_settings(self, prefix="", suffix=""):
        """设置文本"""
        self.prefix_edit.setText(prefix)
        self.suffix_edit.setText(suffix)
        self.prefix_text = prefix
        self.suffix_text = suffix


# 测试代码
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    
    dialog = AddTextDialog("姓名")
    if dialog.exec() == QDialog.Accepted:
        settings = dialog.get_text_settings()
        print(f"前面文本: {settings['prefix']}")
        print(f"后面文本: {settings['suffix']}")
    
    sys.exit(app.exec())