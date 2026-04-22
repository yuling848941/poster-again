"""
图片生成配置模块
负责管理图片生成相关的配置
"""

import logging
from typing import Dict, Any

logger = logging.getLogger(__name__)


class ImageConfigMixin:
    """图片生成配置混合类，提供图片生成相关的配置管理功能"""

    # 贺语模板配置键名
    MESSAGE_TEMPLATE_KEY = "message_template"

    def get_message_template(self) -> str:
        """
        获取保存的贺语模板

        Returns:
            str: 贺语模板内容，如果没有则返回空字符串
        """
        try:
            template = self.get(self.MESSAGE_TEMPLATE_KEY, "")
            return template if template else ""
        except Exception as e:
            logger.error(f"获取贺语模板失败: {str(e)}")
            return ""

    def save_message_template(self, template: str) -> bool:
        """
        保存贺语模板到配置

        Args:
            template: 贺语模板内容

        Returns:
            bool: 保存成功返回True，失败返回False
        """
        try:
            self.set(self.MESSAGE_TEMPLATE_KEY, template)
            result = self.save_config()
            if result:
                logger.info(f"已保存贺语模板，长度: {len(template)} 字符")
            return result
        except Exception as e:
            logger.error(f"保存贺语模板失败: {str(e)}")
            return False

    def load_image_generation_settings(self) -> Dict[str, Any]:
        """
        加载图片生成设置

        Returns:
            Dict[str, Any]: 包含 enabled、format、quality 的字典
        """
        try:
            settings = {
                'enabled': self.get('image_generation.enabled', False),
                'format': self.get('image_generation.format', 'PNG'),
                'quality': self.get('image_generation.quality', 1.0)
            }

            # 验证格式参数
            if settings['format'] not in ['PNG', 'JPG']:
                logger.warning(f"无效的图片格式：{settings['format']}，使用默认值 PNG")
                settings['format'] = 'PNG'

            # 验证质量参数
            if not isinstance(settings['quality'], (int, float)) or settings['quality'] <= 0:
                logger.warning(f"无效的图片质量：{settings['quality']}，使用默认值 1.0")
                settings['quality'] = 1.0

            return settings

        except Exception as e:
            logger.warning(f"加载图片生成设置失败：{str(e)}，使用默认设置")
            return {
                'enabled': False,
                'format': 'PNG',
                'quality': 1.0
            }

    def update_image_generation_settings(self, enabled: bool = None, format: str = None, quality: float = None):
        """
        更新图片生成设置

        Args:
            enabled: 是否启用图片生成
            format: 图片格式 (PNG/JPG)
            quality: 图片质量 (1.0, 1.5, 2.0, 3.0, 4.0)
        """
        try:
            if enabled is not None:
                if not isinstance(enabled, bool):
                    logger.warning(f"enabled 参数必须是布尔值，收到：{type(enabled)}")
                else:
                    self.set("image_generation.enabled", enabled)
                    logger.info(f"已保存图片生成启用状态：{enabled}")

            if format is not None:
                if format not in ['PNG', 'JPG']:
                    logger.warning(f"无效的图片格式：{format}，必须是 PNG 或 JPG")
                else:
                    self.set("image_generation.format", format)
                    logger.info(f"已保存图片格式：{format}")

            if quality is not None:
                if not isinstance(quality, (int, float)) or quality <= 0:
                    logger.warning(f"无效的图片质量：{quality}，必须是正数")
                else:
                    self.set("image_generation.quality", float(quality))
                    logger.info(f"已保存图片质量：{quality}")

            # 保存配置
            self.save_config()

        except Exception as e:
            logger.error(f"更新图片生成设置失败：{str(e)}")
