"""
PPTX处理器工厂
创建和管理PPTX处理器实例

注意：此模块现在仅用于创建 PPTXProcessor，所有 COM 相关代码已移除
"""

import logging
from typing import Optional, Dict, Any

try:
    from src.core.interfaces.template_processor import ITemplateProcessor
    from src.core.processors.pptx_processor import PPTXProcessor
except ImportError:
    from core.interfaces.template_processor import ITemplateProcessor
    from core.processors.pptx_processor import PPTXProcessor

logger = logging.getLogger(__name__)


class OfficeSuiteFactory:
    """
    PPTX处理器工厂类

    现在仅用于创建 PPTXProcessor（纯 Python，无需 Office/WPS）
    """

    # 默认使用 PPTXProcessor
    DEFAULT_PROCESSOR = PPTXProcessor

    @classmethod
    def create_pptx_processor(cls, config: Optional[Dict[str, Any]] = None) -> PPTXProcessor:
        """
        创建 PPTX 处理器

        纯 Python 实现，无需 Microsoft Office 或 WPS

        Args:
            config: 可选的配置参数

        Returns:
            PPTXProcessor: PPTX 处理器实例
        """
        processor = PPTXProcessor()

        if config:
            cls._apply_config(processor, config)

        logger.info("创建了 PPTX 处理器（纯 Python，无需 Office/WPS）")
        return processor

    @classmethod
    def _apply_config(cls, processor: ITemplateProcessor, config: Dict[str, Any]):
        """
        应用配置到处理器

        Args:
            processor: 处理器实例
            config: 配置参数
        """
        if hasattr(processor, 'apply_config'):
            processor.apply_config(config)


# 便捷函数
def create_processor(config: Optional[Dict[str, Any]] = None) -> PPTXProcessor:
    """
    创建 PPTX 处理器的便捷函数

    Args:
        config: 可选的配置参数

    Returns:
        PPTXProcessor: PPTX 处理器实例
    """
    return OfficeSuiteFactory.create_pptx_processor(config)
