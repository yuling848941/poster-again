"""
配置管理模块 - 重构版本
用于管理应用程序的配置信息

模块结构：
- config_manager.py: 主配置管理器（外观模式）
- path_config.py: 路径配置管理
- image_config.py: 图片生成配置管理
- placeholder_config.py: 占位符配置管理
- gui_config.py: GUI 配置管理
"""

import os
import sys
import copy
import yaml
import json
import logging
from typing import Dict, Any, Optional, List
from pathlib import Path
from datetime import datetime

# 导入子模块
try:
    from .path_config import PathConfigMixin
    from .image_config import ImageConfigMixin
    from .placeholder_config import PlaceholderConfigMixin
    from .gui_config import GUIConfigMixin
except ImportError:
    from path_config import PathConfigMixin
    from image_config import ImageConfigMixin
    from placeholder_config import PlaceholderConfigMixin
    from gui_config import GUIConfigMixin

logger = logging.getLogger(__name__)


class ConfigManager(PathConfigMixin, ImageConfigMixin, PlaceholderConfigMixin,
                    GUIConfigMixin):
    """
    配置管理器，负责加载、保存和管理应用程序配置

    使用多重继承组合各个配置功能模块：
    - PathConfigMixin: 路径配置管理
    - ImageConfigMixin: 图片生成配置管理
    - PlaceholderConfigMixin: 占位符配置管理
    - GUIConfigMixin: GUI 配置管理
    """

    def __init__(self, config_file: str = "PosterAgain_config.yaml"):
        """
        初始化配置管理器

        Args:
            config_file: 配置文件路径
        """
        # 处理配置文件路径：区分开发环境和打包环境
        if getattr(sys, 'frozen', False):
            # 打包环境：使用 exe 所在目录，配置文件直接放在 exe 目录（隐藏文件）
            exe_dir = os.path.dirname(os.path.abspath(sys.executable))
            config_file_name = os.path.basename(config_file)
            # 配置文件直接放在 exe 目录（隐藏文件，以 . 开头）
            self.config_file = os.path.join(exe_dir, f".{config_file_name}")
        else:
            # 开发环境：保持原有相对路径，但在项目根目录
            if os.path.exists("PosterAgain_config.yaml"):
                self.config_file = "PosterAgain_config.yaml"
            elif os.path.exists("config/PosterAgain_config.yaml"):
                self.config_file = "config/PosterAgain_config.yaml"
            else:
                self.config_file = config_file

        self.config = {}
        self.default_config = {
            "app": {
                "name": "PPT 批量生成工具",
                "version": "1.0.0",
                "author": "AI Assistant",
                "description": "基于 Excel 数据批量生成 PPT 文件的工具"
            },
            "ui": {
                "theme": "light",
                "language": "zh_CN",
                "window_size": {
                    "width": 1000,
                    "height": 700
                },
                "remember_window_state": True,
                "show_tips": True
            },
            "processing": {
                "use_com_interface": False,
                "auto_save": True,
                "backup_files": True,
                "max_undo_steps": 10
            },
            "paths": {
                "last_template_dir": "",
                "last_data_dir": "",
                "last_output_dir": "",
                "default_output_dir": "./output",
                "last_config_save_dir": "",
                "last_config_load_dir": ""
            },
            "matching": {
                "auto_match": True,
                "case_sensitive": False,
                "fuzzy_match": True,
                "save_matching_rules": True
            },
            "output": {
                "file_name_pattern": "output_{index}.pptx",
                "overwrite_existing": False,
                "create_subfolders": False,
                "include_date_in_name": False
            },
            "advanced": {
                "log_level": "INFO",
                "max_log_size": "10MB",
                "log_file_count": 5,
                "enable_debug_mode": False
            },
            "image_generation": {
                "enabled": False,
                "format": "PNG",
                "quality": 1.0
            },
            "office_cache": {
                "cached": False,
                "type": None,
                "name": None,
                "path": None,
                "version": None,
                "progid": None,
                "detected_at": None
            }
        }

        # 确保配置文件目录存在
        self.config_dir = os.path.dirname(os.path.abspath(self.config_file))
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir, exist_ok=True)

        # 加载配置
        self.load_config()

        logger.info(f"配置管理器初始化完成，配置文件：{self.config_file}")

    def load_config(self) -> bool:
        """
        加载配置文件

        Returns:
            bool: 加载成功返回 True，失败返回 False
        """
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    if self.config_file.endswith('.yaml') or self.config_file.endswith('.yml'):
                        loaded_config = yaml.safe_load(f)
                    else:
                        loaded_config = json.load(f)

                # 合并默认配置和加载的配置
                self.config = self._merge_config(self.default_config, loaded_config)
                logger.info(f"成功加载配置文件：{self.config_file}")
                return True
            else:
                # 如果配置文件不存在，使用默认配置并保存
                self.config = copy.deepcopy(self.default_config)
                self.save_config()
                logger.info(f"配置文件不存在，使用默认配置并创建：{self.config_file}")
                return True
        except Exception as e:
            logger.error(f"加载配置文件失败：{str(e)}")
            self.config = copy.deepcopy(self.default_config)
            return False

    def save_config(self) -> bool:
        """
        保存配置到文件

        Returns:
            bool: 保存成功返回 True，失败返回 False
        """
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                if self.config_file.endswith('.yaml') or self.config_file.endswith('.yml'):
                    yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True, indent=2)
                else:
                    json.dump(self.config, f, ensure_ascii=False, indent=2)

            logger.info(f"成功保存配置文件：{self.config_file}")
            return True
        except Exception as e:
            logger.error(f"保存配置文件失败：{str(e)}")
            return False

    def _merge_config(self, default: Dict[str, Any], loaded: Dict[str, Any]) -> Dict[str, Any]:
        """
        合并默认配置和加载的配置

        Args:
            default: 默认配置
            loaded: 加载的配置

        Returns:
            Dict[str, Any]: 合并后的配置
        """
        result = copy.deepcopy(default)

        for key, value in loaded.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._merge_config(result[key], value)
            else:
                result[key] = value

        return result

    def get(self, key_path: str, default_value: Any = None) -> Any:
        """
        获取配置值

        Args:
            key_path: 配置键路径，如 "ui.theme" 或 "paths.last_template_dir"
            default_value: 默认值

        Returns:
            Any: 配置值
        """
        keys = key_path.split('.')
        value = self.config

        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default_value

    def set(self, key_path: str, value: Any) -> bool:
        """
        设置配置值

        Args:
            key_path: 配置键路径，如 "ui.theme" 或 "paths.last_template_dir"
            value: 配置值

        Returns:
            bool: 设置成功返回 True，失败返回 False
        """
        keys = key_path.split('.')
        config = self.config

        try:
            # 导航到最后一级的父级
            for key in keys[:-1]:
                if key not in config:
                    config[key] = {}
                config = config[key]

            # 设置值
            config[keys[-1]] = value
            return True
        except Exception as e:
            logger.error(f"设置配置值失败：{str(e)}")
            return False

    def get_app_info(self) -> Dict[str, str]:
        """
        获取应用程序信息

        Returns:
            Dict[str, str]: 应用程序信息
        """
        return {
            "name": self.get("app.name", "PPT 批量生成工具"),
            "version": self.get("app.version", "1.0.0"),
            "author": self.get("app.author", "AI Assistant"),
            "description": self.get("app.description", "基于 Excel 数据批量生成 PPT 文件的工具")
        }

    def get_ui_config(self) -> Dict[str, Any]:
        """
        获取 UI 配置

        Returns:
            Dict[str, Any]: UI 配置
        """
        return self.get("ui", {})

    def get_processing_config(self) -> Dict[str, Any]:
        """
        获取处理配置

        Returns:
            Dict[str, Any]: 处理配置
        """
        return self.get("processing", {})

    def get_output_config(self) -> Dict[str, Any]:
        """
        获取输出配置

        Returns:
            Dict[str, Any]: 输出配置
        """
        return self.get("output", {})

    def get_advanced_config(self) -> Dict[str, Any]:
        """
        获取高级配置

        Returns:
            Dict[str, Any]: 高级配置
        """
        return self.get("advanced", {})

    def export_config(self, file_path: str) -> bool:
        """
        导出配置到文件

        Args:
            file_path: 导出文件路径

        Returns:
            bool: 导出成功返回 True，失败返回 False
        """
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                if file_path.endswith('.yaml') or file_path.endswith('.yml'):
                    yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True, indent=2)
                else:
                    json.dump(self.config, f, ensure_ascii=False, indent=2)

            logger.info(f"成功导出配置到：{file_path}")
            return True
        except Exception as e:
            logger.error(f"导出配置失败：{str(e)}")
            return False

    def import_config(self, file_path: str) -> bool:
        """
        从文件导入配置

        Args:
            file_path: 导入文件路径

        Returns:
            bool: 导入成功返回 True，失败返回 False
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                if file_path.endswith('.yaml') or file_path.endswith('.yml'):
                    loaded_config = yaml.safe_load(f)
                else:
                    loaded_config = json.load(f)

            # 合并配置
            self.config = self._merge_config(self.default_config, loaded_config)

            # 保存配置
            self.save_config()

            logger.info(f"成功从文件导入配置：{file_path}")
            return True
        except Exception as e:
            logger.error(f"导入配置失败：{str(e)}")
            return False

    def reset_to_default(self) -> bool:
        """
        重置为默认配置

        Returns:
            bool: 重置成功返回 True，失败返回 False
        """
        try:
            self.config = copy.deepcopy(self.default_config)
            self.save_config()
            logger.info("配置已重置为默认值")
            return True
        except Exception as e:
            logger.error(f"重置配置失败：{str(e)}")
            return False

    def get_all_config(self) -> Dict[str, Any]:
        """
        获取所有配置

        Returns:
            Dict[str, Any]: 所有配置
        """
        return copy.deepcopy(self.config)

    # ========== Office 缓存管理方法 ==========

    def save_office_cache(self, office_info: Dict[str, Any]) -> bool:
        """
        保存 Office 检测到缓存

        Args:
            office_info: Office 检测信息，包含 type, name, path, version, progid 等

        Returns:
            bool: 保存成功返回 True，失败返回 False
        """
        try:
            from datetime import datetime

            self.set("office_cache.cached", True)
            self.set("office_cache.type", office_info.get('type'))
            self.set("office_cache.name", office_info.get('name'))
            self.set("office_cache.path", office_info.get('path'))
            self.set("office_cache.version", office_info.get('version'))
            self.set("office_cache.progid", office_info.get('progid'))
            self.set("office_cache.detected_at", datetime.now().isoformat())

            return self.save_config()
        except Exception as e:
            logger.error(f"保存 Office 缓存失败：{str(e)}")
            return False

    def load_office_cache(self) -> Optional[Dict[str, Any]]:
        """
        从缓存加载 Office 信息

        Returns:
            Optional[Dict]: 缓存的 Office 信息，如果缓存不存在或无效则返回 None
        """
        try:
            if not self.get("office_cache.cached", False):
                return None

            office_info = {
                'type': self.get("office_cache.type"),
                'name': self.get("office_cache.name"),
                'path': self.get("office_cache.path"),
                'version': self.get("office_cache.version"),
                'progid': self.get("office_cache.progid"),
                'detected_at': self.get("office_cache.detected_at")
            }

            # 验证缓存数据是否完整
            if not office_info.get('type') or not office_info.get('name'):
                logger.debug("Office 缓存数据不完整")
                return None

            # 验证路径是否仍然有效（如果是文件系统路径）
            if office_info.get('path') and office_info['path'] != 'COM 接口':
                import os
                if not os.path.exists(office_info['path']):
                    logger.info("Office 缓存路径已失效，需要重新检测")
                    return None

            logger.info(f"从缓存加载 Office 信息：{office_info['name']}")
            return office_info

        except Exception as e:
            logger.error(f"加载 Office 缓存失败：{str(e)}")
            return None

    def clear_office_cache(self) -> bool:
        """
        清除 Office 缓存

        Returns:
            bool: 清除成功返回 True，失败返回 False
        """
        try:
            self.set("office_cache.cached", False)
            self.set("office_cache.type", None)
            self.set("office_cache.name", None)
            self.set("office_cache.path", None)
            self.set("office_cache.version", None)
            self.set("office_cache.progid", None)
            self.set("office_cache.detected_at", None)

            return self.save_config()
        except Exception as e:
            logger.error(f"清除 Office 缓存失败：{str(e)}")
            return False

    def get_office_cache_status(self) -> Dict[str, Any]:
        """
        获取 Office 缓存状态

        Returns:
            Dict[str, Any]: 缓存状态信息
        """
        return {
            'cached': self.get("office_cache.cached", False),
            'type': self.get("office_cache.type"),
            'name': self.get("office_cache.name"),
            'version': self.get("office_cache.version"),
            'detected_at': self.get("office_cache.detected_at")
        }


# 向后兼容的导出
__all__ = ['ConfigManager']
