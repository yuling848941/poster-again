#!/usr/bin/env python3
"""
PPT 批量生成工具 - GUI 应用程序入口

使用纯 Python 的 python-pptx 库进行 PPT 生成，无需 Microsoft Office 或 WPS。
图片导出同样使用内置纯 Python 渲染引擎（Pillow），不依赖任何 Office 套件；
渲染前自动预检，不支持元素的告警见日志 [预检] 条目（详见 AGENTS.md）。
"""

import sys
import os
import logging
from logging.handlers import RotatingFileHandler

# 确保项目根目录在 Python 路径中，保证 src.xxx 导入可用
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# 使用模块级 logger
logger = logging.getLogger(__name__)


def _log_dir() -> str:
    """日志目录：放在用户本地数据目录，不污染项目/安装目录，打包后也可写。"""
    if sys.platform == 'win32':
        base = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
        return os.path.join(base, 'PosterAgain', 'logs')
    return os.path.join(os.path.expanduser('~'), '.local', 'share', 'PosterAgain', 'logs')


def _parse_size(text, default: int) -> int:
    """把 '10MB' / '512KB' / '1048576' 解析成字节数。"""
    try:
        s = str(text).strip().upper()
        for suffix, factor in (('GB', 1024 ** 3), ('MB', 1024 ** 2), ('KB', 1024), ('B', 1)):
            if s.endswith(suffix):
                return int(float(s[:-len(suffix)].strip()) * factor)
        return int(s)
    except (ValueError, TypeError):
        return default


def setup_logging():
    """
    初始化全局日志配置（整个应用唯一的 basicConfig 调用点）。

    - 级别取运行时配置 advanced.log_level（默认 INFO）；
    - 同时输出到控制台和滚动日志文件（advanced.max_log_size / log_file_count），
      打包成无控制台 exe 后日志仍可在日志目录查到；
    - 安装全局异常钩子，未捕获异常（含工作线程）写入日志而不是无声消失。

    其他模块应继续使用 logging.getLogger(__name__)，避免重复配置。
    """
    level_str = str(config_manager.get('advanced.log_level', 'INFO') or 'INFO').upper()
    log_level = getattr(logging, level_str, logging.INFO)
    max_bytes = _parse_size(config_manager.get('advanced.max_log_size', '10MB'), 10 * 1024 * 1024)
    backup_count = int(config_manager.get('advanced.log_file_count', 5) or 5)

    handlers = [logging.StreamHandler()]
    log_file = None
    try:
        log_dir = _log_dir()
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, 'posteragain.log')
        handlers.append(RotatingFileHandler(
            log_file, maxBytes=max_bytes, backupCount=backup_count, encoding='utf-8'))
    except Exception as e:
        print(f"[WARN] 日志文件初始化失败，仅输出到控制台: {e}")

    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
    )
    if log_file:
        logger.info(f"日志文件: {log_file}")
    _install_excepthooks()


def _install_excepthooks():
    """未捕获异常写入日志（主线程 + 工作线程），便于打包后排查问题。"""
    import threading

    def _hook(exc_type, exc_value, exc_tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        logging.getLogger('uncaught').critical(
            '未捕获异常', exc_info=(exc_type, exc_value, exc_tb))

    def _thread_hook(args):
        if issubclass(args.exc_type, SystemExit):
            return
        name = args.thread.name if args.thread else '?'
        logging.getLogger('uncaught').critical(
            f'线程 {name} 未捕获异常',
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback))

    sys.excepthook = _hook
    threading.excepthook = _thread_hook


# 导入项目模块（保持 src. 前缀 + 扁平导入的双模式回退）
try:
    from src.config import ConfigManager
    from src.gui.main_window import MainWindow
except ImportError:
    from config_manager import ConfigManager
    from gui.main_window import MainWindow

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon

# 程序启动时初始化配置管理器（模块级单例，导入 main 会有副作用，AGENTS.md 已说明）
config_manager = ConfigManager("config/config.yaml")


def main():
    """主函数"""
    setup_logging()
    logger.info("启动 PPT 批量生成工具")
    logger.info("使用纯 Python PPTX 处理器（PPT 生成无需 Office/WPS）")

    # 字体解析策略注入（image_generation.font_policy，见 font_registry.py）
    from render_core import font_registry
    font_registry.set_policy(str(config_manager.get(
        "image_generation.font_policy", "auto") or "auto"))

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
