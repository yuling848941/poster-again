"""
路径配置模块
负责管理应用程序的路径相关配置
"""

import os
import logging
from typing import Dict

logger = logging.getLogger(__name__)


class PathConfigMixin:
    """路径配置混合类，提供路径相关的配置管理功能"""

    def update_last_paths(self, template_dir: str = None, data_dir: str = None, output_dir: str = None):
        """
        更新最近使用的路径

        Args:
            template_dir: 模板目录
            data_dir: 数据目录
            output_dir: 输出目录
        """
        try:
            if template_dir and self.validate_path(template_dir):
                self.set("paths.last_template_dir", template_dir)
                logger.info(f"已保存模板路径：{template_dir}")

            if data_dir and self.validate_path(data_dir):
                self.set("paths.last_data_dir", data_dir)
                logger.info(f"已保存数据路径：{data_dir}")

            if output_dir and self.validate_path(output_dir):
                self.set("paths.last_output_dir", output_dir)
                logger.info(f"已保存输出路径：{output_dir}")

            # 保存配置
            self.save_config()

        except Exception as e:
            logger.error(f"更新最近路径失败：{str(e)}")

    def load_last_paths(self) -> Dict[str, str]:
        """
        加载上次保存的路径

        Returns:
            Dict[str, str]: 包含所有路径类型的字典
        """
        try:
            paths = {
                'template_dir': self.get('paths.last_template_dir', ''),
                'data_dir': self.get('paths.last_data_dir', ''),
                'output_dir': self.get('paths.last_output_dir', '')
            }
            return paths
        except Exception as e:
            logger.warning(f"加载上次路径失败：{str(e)}")
            return {
                'template_dir': '',
                'data_dir': '',
                'output_dir': ''
            }

    def get_start_dir_for_file_type(self, file_type: str) -> str:
        """
        根据文件类型获取起始目录

        Args:
            file_type: 文件类型 ('template', 'data', 'output')

        Returns:
            str: 起始目录路径
        """
        try:
            path_map = {
                'template': self.get('paths.last_template_dir', ''),
                'data': self.get('paths.last_data_dir', ''),
                'output': self.get('paths.last_output_dir', '')
            }

            path = path_map.get(file_type, '')

            # 如果路径为空或无效，返回用户主目录
            if not path or not self.validate_path(path):
                return os.path.expanduser('~')

            # 如果是文件路径，返回其所在目录
            if os.path.isfile(path):
                return os.path.dirname(path)

            return path

        except Exception as e:
            logger.warning(f"获取文件类型起始目录失败：{str(e)}")
            return os.path.expanduser('~')

    def validate_path(self, path: str) -> bool:
        """
        验证路径有效性

        Args:
            path: 要验证的路径

        Returns:
            bool: 路径是否有效
        """
        if not path:
            return False

        try:
            return os.path.exists(path)
        except Exception as e:
            logger.warning(f"路径验证失败：{str(e)}")
            return False

    def get_paths_config(self) -> Dict:
        """
        获取路径配置

        Returns:
            Dict: 路径配置
        """
        return self.get("paths", {})
