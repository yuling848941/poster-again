"""
模板处理器抽象接口定义
提供统一的模板处理接口，隐藏不同办公套件的实现差异
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Tuple
import logging

logger = logging.getLogger(__name__)


class ITemplateProcessor(ABC):
    """模板处理器抽象接口"""

    def __init__(self):
        """初始化模板处理器"""
        self.logger = logging.getLogger(self.__class__.__name__)
        self.is_connected = False
        self.presentation = None

    @abstractmethod
    def connect_to_application(self) -> bool:
        """
        连接到办公套件应用程序

        Returns:
            bool: 连接是否成功
        """
        pass

    @abstractmethod
    def load_template(self, template_path: str) -> bool:
        """
        加载模板文件

        Args:
            template_path: 模板文件路径

        Returns:
            bool: 加载是否成功
        """
        pass

    @abstractmethod
    def find_placeholders(self) -> List[str]:
        """
        查找模板中的占位符

        Returns:
            List[str]: 占位符列表
        """
        pass

    @abstractmethod
    def replace_placeholders(self, data: Dict[str, str]) -> bool:
        """
        替换占位符内容

        Args:
            data: 占位符和对应值的字典

        Returns:
            bool: 替换是否成功
        """
        pass

    @abstractmethod
    def save_presentation(self, output_path: str) -> bool:
        """
        保存演示文稿

        Args:
            output_path: 输出文件路径

        Returns:
            bool: 保存是否成功
        """
        pass

    @abstractmethod
    def export_to_images(self, output_dir: str, format: str, quality: float) -> bool:
        """
        导出演示文稿为图片

        Args:
            output_dir: 输出目录
            format: 图片格式 (PNG, JPG)
            quality: 图片质量/分辨率 (1.0-4.0)

        Returns:
            bool: 导出是否成功
        """
        pass

    @abstractmethod
    def close_presentation(self):
        """关闭当前演示文稿"""
        pass

    @abstractmethod
    def close_application(self):
        """关闭办公套件应用程序"""
        pass

    def get_supported_formats(self) -> List[str]:
        """
        获取支持的文件格式列表

        Returns:
            List[str]: 支持的格式列表
        """
        return ["pptx"]

    def get_application_info(self) -> Dict[str, str]:
        """
        获取办公套件应用程序信息

        Returns:
            Dict[str, str]: 应用程序信息字典
        """
        return {
            "name": "Unknown",
            "version": "Unknown",
            "processor": self.__class__.__name__
        }

    def is_connected_to_app(self) -> bool:
        """
        检查是否已连接到办公套件应用程序

        Returns:
            bool: 是否已连接
        """
        return self.is_connected

    def validate_template_path(self, template_path: str) -> bool:
        """
        验证模板文件路径是否有效

        Args:
            template_path: 模板文件路径

        Returns:
            bool: 路径是否有效
        """
        import os
        return os.path.exists(template_path) and template_path.lower().endswith('.pptx')

    def validate_data_format(self, data: Dict[str, str]) -> bool:
        """
        验证替换数据格式是否有效

        Args:
            data: 替换数据字典

        Returns:
            bool: 数据格式是否有效
        """
        if not isinstance(data, dict):
            return False

        for key, value in data.items():
            if not isinstance(key, str) or not isinstance(value, str):
                return False

        return True

    def get_processor_capabilities(self) -> Dict[str, bool]:
        """
        获取处理器能力信息

        Returns:
            Dict[str, bool]: 能力信息字典
        """
        return {
            "can_connect": True,
            "can_load_template": True,
            "can_find_placeholders": True,
            "can_replace_placeholders": True,
            "can_save_presentation": True,
            "can_export_images": True,
            "supports_batch_processing": True,
            "supports_high_quality_export": True
        }


class TemplateProcessorError(Exception):
    """模板处理器异常基类"""
    pass


class ConnectionError(TemplateProcessorError):
    """连接异常"""
    pass


class TemplateLoadError(TemplateProcessorError):
    """模板加载异常"""
    pass


class PlaceholderError(TemplateProcessorError):
    """占位符处理异常"""
    pass


class SaveError(TemplateProcessorError):
    """保存异常"""
    pass


class ExportError(TemplateProcessorError):
    """导出异常"""
    pass