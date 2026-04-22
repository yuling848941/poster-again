#!/usr/bin/env python3
"""
办公套件选择功能演示脚本
展示如何在GUI中使用办公套件选择器
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit
from PySide6.QtCore import Qt
from src.gui.office_suite_detector import OfficeSuiteManager, OfficeSuite
from src.config import ConfigManager

class OfficeSelectorDemo(QWidget):
    """办公套件选择器演示窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("办公套件选择器演示")
        self.setGeometry(100, 100, 500, 400)

        # 初始化办公套件管理器
        self.config_manager = ConfigManager()
        self.office_manager = OfficeSuiteManager(self.config_manager)
        self.office_manager.initialize()

        self.setup_ui()
        self.update_status()

    def setup_ui(self):
        """设置UI界面"""
        layout = QVBoxLayout(self)

        # 标题
        title = QLabel("办公套件选择器功能演示")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold; margin: 20px;")
        layout.addWidget(title)

        # 状态显示
        self.status_label = QLabel("状态: 已初始化")
        self.status_label.setStyleSheet("padding: 10px; background-color: #f0f0f0; border-radius: 5px;")
        layout.addWidget(self.status_label)

        # 当前选择显示
        self.current_label = QLabel("当前选择: ")
        self.current_label.setStyleSheet("padding: 10px; background-color: #e8f5e8; border-radius: 5px;")
        layout.addWidget(self.current_label)

        # 可用套件显示
        self.available_label = QLabel("可用套件: ")
        self.available_label.setStyleSheet("padding: 10px; background-color: #e8f4ff; border-radius: 5px;")
        layout.addWidget(self.available_label)

        # 操作按钮
        test_auto_btn = QPushButton("测试自动检测")
        test_auto_btn.clicked.connect(self.test_auto_detection)
        test_auto_btn.setStyleSheet("padding: 10px; margin: 5px;")
        layout.addWidget(test_auto_btn)

        test_ms_btn = QPushButton("强制选择Microsoft PowerPoint")
        test_ms_btn.clicked.connect(lambda: self.select_suite(OfficeSuite.MICROSOFT))
        test_ms_btn.setStyleSheet("padding: 10px; margin: 5px; background-color: #4CAF50; color: white;")
        layout.addWidget(test_ms_btn)

        test_wps_btn = QPushButton("强制选择WPS演示")
        test_wps_btn.clicked.connect(lambda: self.select_suite(OfficeSuite.WPS))
        test_wps_btn.setStyleSheet("padding: 10px; margin: 5px; background-color: #FF9800; color: white;")
        layout.addWidget(test_wps_btn)

        refresh_btn = QPushButton("刷新检测")
        refresh_btn.clicked.connect(self.refresh_detection)
        refresh_btn.setStyleSheet("padding: 10px; margin: 5px; background-color: #2196F3; color: white;")
        layout.addWidget(refresh_btn)

        # 日志显示
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(150)
        self.log_text.setPlaceholderText("操作日志...")
        layout.addWidget(self.log_text)

    def update_status(self):
        """更新状态显示"""
        current_suite = self.office_manager.get_current_suite()
        effective_suite = self.office_manager.get_effective_suite()
        available_suites = self.office_manager.get_available_suites()

        self.current_label.setText(
            f"当前选择: {current_suite.display_name}\n"
            f"实际使用: {effective_suite.display_name}"
        )

        if available_suites:
            suite_names = [suite.display_name for suite in available_suites]
            self.available_label.setText(f"可用套件: {', '.join(suite_names)}")
        else:
            self.available_label.setText("可用套件: 无")

        self.status_label.setText("状态: 已就绪")

    def test_auto_detection(self):
        """测试自动检测"""
        self.log("开始测试自动检测...")
        success = self.office_manager.set_current_suite(OfficeSuite.AUTO)
        if success:
            effective = self.office_manager.get_effective_suite()
            self.log(f"✓ 自动检测成功，选择了: {effective.display_name}")
        else:
            self.log("✗ 自动检测失败")
        self.update_status()

    def select_suite(self, suite: OfficeSuite):
        """选择特定套件"""
        self.log(f"尝试选择: {suite.display_name}")
        success = self.office_manager.set_current_suite(suite)
        if success:
            effective = self.office_manager.get_effective_suite()
            self.log(f"✓ 选择成功，实际使用: {effective.display_name}")
        else:
            self.log(f"✗ 选择失败: {suite.display_name} 不可用")
        self.update_status()

    def refresh_detection(self):
        """刷新检测"""
        self.log("开始刷新办公套件检测...")
        self.office_manager.initialize()

        available_suites = self.office_manager.get_available_suites()
        if available_suites:
            suite_names = [suite.display_name for suite in available_suites]
            self.log(f"✓ 检测完成，找到: {', '.join(suite_names)}")

            # 显示详细信息
            for suite in available_suites:
                info = self.office_manager.detector.get_suite_info(suite)
                self.log(f"  - {suite.display_name}: {info.get('name', '未知')} {info.get('version', '')}")
        else:
            self.log("✗ 未检测到可用的办公套件")

        self.update_status()

    def log(self, message: str):
        """添加日志消息"""
        self.log_text.append(message)

def main():
    """主函数"""
    print("=" * 60)
    print("办公套件选择器功能演示")
    print("=" * 60)

    app = QApplication(sys.argv)

    try:
        demo = OfficeSelectorDemo()
        demo.show()

        print("演示窗口已启动")
        print("请尝试以下功能:")
        print("1. 查看当前检测到的办公套件")
        print("2. 点击不同按钮测试套件切换")
        print("3. 查看日志输出了解详细信息")
        print("\n关闭窗口退出演示")

        return app.exec()

    except Exception as e:
        print(f"演示启动失败: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())