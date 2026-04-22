"""
配置管理模块 - 兼容层
用于管理应用程序的配置信息

注意：此模块现在是向后兼容的包装器，实际实现在 src.config 包中
"""

import os
import sys

# 从新的配置包导入 ConfigManager，保持向后兼容
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

try:
    from src.config import ConfigManager
except ImportError:
    from config import ConfigManager

__all__ = ['ConfigManager']
