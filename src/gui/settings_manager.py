"""
SettingsManager — 管理图片生成设置、贺语模板和防抖保存

从 main_window.py 中拆分出来的设置管理模块，负责:
- 图片生成设置 (启用/格式/质量) 的加载、保存、防抖持久化
- 贺语模板的加载和保存
- 质量值与显示文本的双向转换
"""

import logging

logger = logging.getLogger(__name__)

# 图片质量映射表 (唯一数据源)
QUALITY_VALUE_TO_TEXT = {
    1.0: "原始大小",
    1.5: "1.5倍",
    2.0: "2倍",
    3.0: "3倍",
    4.0: "4倍"
}

QUALITY_TEXT_TO_VALUE = {v: k for k, v in QUALITY_VALUE_TO_TEXT.items()}


class SettingsManager:
    """管理图片生成设置和贺语模板的加载/保存"""

    def __init__(self, config_manager):
        self.config_manager = config_manager

    # ── 图片生成设置 ──────────────────────────────────────────

    def load_image_settings(self):
        """
        加载图片生成设置

        Returns:
            dict: {'enabled': bool, 'format': str, 'quality': float}
        """
        defaults = {'enabled': False, 'format': 'PNG', 'quality': 1.0}
        try:
            settings = self.config_manager.load_image_generation_settings()
            if not isinstance(settings, dict):
                logger.warning("图片生成设置格式无效，使用默认值")
                return defaults
            return {
                'enabled': settings.get('enabled', defaults['enabled']),
                'format': settings.get('format', defaults['format']),
                'quality': settings.get('quality', defaults['quality']),
            }
        except Exception as e:
            logger.error(f"加载图片生成设置失败: {e}")
            return defaults

    def save_image_settings(self, enabled, fmt, quality):
        """保存图片生成设置"""
        try:
            self.config_manager.update_image_generation_settings(
                enabled=enabled,
                format=fmt,
                quality=quality
            )
            logger.info("图片生成设置保存成功")
        except Exception as e:
            logger.error(f"保存图片生成设置失败: {e}")

    # ── 贺语模板 ──────────────────────────────────────────────

    def load_message_template(self):
        """加载保存的贺语模板"""
        try:
            return self.config_manager.get_message_template() or ""
        except Exception as e:
            logger.error(f"加载贺语模板失败: {e}")
            return ""

    def save_message_template(self, template):
        """保存贺语模板"""
        try:
            self.config_manager.save_message_template(template)
        except Exception as e:
            logger.error(f"保存贺语模板失败: {e}")

    # ── 质量值转换 ────────────────────────────────────────────

    @staticmethod
    def quality_value_to_text(value):
        return QUALITY_VALUE_TO_TEXT.get(value, "原始大小")

    @staticmethod
    def quality_text_to_value(text):
        return QUALITY_TEXT_TO_VALUE.get(text, 1.0)
