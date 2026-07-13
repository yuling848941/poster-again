#!/usr/bin/env python3
"""
PPT 批量生成工具 - GUI 应用程序入口

使用纯 Python 的 python-pptx 库进行 PPT 生成，无需 Microsoft Office 或 WPS。
（图片导出为可选功能，需要 Office 套件，详见 AGENTS.md）
"""

import sys
import os
import logging

# 确保项目根目录在 Python 路径中，保证 src.xxx 导入可用
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 使用模块级 logger
logger = logging.getLogger(__name__)


def setup_logging():
    """
    初始化全局日志配置

    读取 config/config.yaml 的 advanced.log_level（默认 INFO），
    输出到控制台。这是整个应用唯一的 basicConfig 调用点；
    其他模块应继续使用 logging.getLogger(__name__)，避免重复配置。
    """
    log_level = logging.INFO
    try:
        import yaml
        config_path = os.path.join(current_dir, "config", "config.yaml")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            level_str = (cfg.get("advanced", {}) or {}).get("log_level", "INFO")
            log_level = getattr(logging, str(level_str).upper(), logging.INFO)
    except Exception as e:
        # 配置读取失败时退回默认 INFO，不阻断启动
        print(f"[WARN] 读取日志级别失败，使用默认 INFO: {e}")

    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


# 导入项目模块（保持 src. 前缀 + 扁平导入的双模式回退）
try:
    from src.memory_management import MemoryDataManager
    from src.config import ConfigManager
    from src.gui.main_window import MainWindow
except ImportError:
    from memory_management import MemoryDataManager
    from config_manager import ConfigManager
    from gui.main_window import MainWindow

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon

# 程序启动时初始化内存管理器和配置管理器（与历史版本保持一致）
# 注意：这些是模块级单例，导入 main 会有副作用（AGENTS.md 已说明）
config_manager = ConfigManager("config/config.yaml")
memory_manager = MemoryDataManager()


def main():
    """主函数"""
    setup_logging()
    logger.info("启动 PPT 批量生成工具")
    logger.info("使用纯 Python PPTX 处理器（PPT 生成无需 Office/WPS）")

    app = QApplication(sys.argv)

    # 设置应用程序属性
    app.setApplicationName("PPT 批量生成工具")
    app.setOrganizationName("Poster Generator")

    # 设置应用图标（任务栏显示）
    icon_path = os.path.join(current_dir, "resources", "icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    else:
        logger.debug(f"未找到图标文件: {icon_path}")

    # 创建并显示主窗口
    window = MainWindow()
    window.show()

    # 运行应用程序
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
