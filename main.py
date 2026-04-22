#!/usr/bin/env python3
"""
PPT批量生成工具 - GUI应用程序入口

注意：此版本使用纯 Python 的 python-pptx 库，无需 Microsoft Office 或 WPS
"""

import sys
import os
import logging

logger = logging.getLogger(__name__)

# 确保当前目录在Python路径中
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 导入新的内存管理器
from memory_management import MemoryDataManager
from config_manager import ConfigManager

# 程序启动时初始化内存管理器和配置管理器
memory_manager = MemoryDataManager()
config_manager = ConfigManager("config/config.yaml")
logger.info("Memory manager and config manager initialized successfully")
logger.info("使用纯 Python PPTX 处理器（无需 Office/WPS）")

from gui.main_window import MainWindow
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon


def main():
    """主函数"""
    app = QApplication(sys.argv)
    
    # 设置应用程序属性
    app.setApplicationName("PPT批量生成工具")
    app.setOrganizationName("Poster Generator")
    
    # 设置应用图标（任务栏显示）
    icon_path = os.path.join(current_dir, "resources", "icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    
    # 创建并显示主窗口
    window = MainWindow()
    window.show()
    
    # 运行应用程序
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
