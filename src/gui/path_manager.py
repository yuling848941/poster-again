"""
PathManager — 管理文件路径的加载、保存和恢复

从 main_window.py 中拆分出来的路径管理模块，负责:
- 加载上次使用的路径 (模板/数据/输出)
- 恢复路径到输入框并设置到工作线程
- 路径验证
"""

import logging

logger = logging.getLogger(__name__)


class PathManager:
    """管理文件路径的加载、保存和恢复"""

    def __init__(self, config_manager):
        self.config_manager = config_manager

    def load_last_paths(self):
        """
        加载上次使用的路径

        Returns:
            dict: {'template_dir': str, 'data_dir': str, 'output_dir': str}
        """
        try:
            return self.config_manager.load_last_paths() or {}
        except Exception as e:
            logger.warning(f"加载上次路径失败: {e}")
            return {}

    def restore_path(self, path, edit_widget, setter_func):
        """
        将路径恢复到输入框并设置到工作线程

        Args:
            path: 要恢复的路径
            edit_widget: 输入框控件 (QLineEdit)
            setter_func: 设置路径的方法
        """
        try:
            if path and edit_widget and setter_func:
                edit_widget.setText(path)
                setter_func(path)
        except Exception as e:
            logger.warning(f"恢复路径失败: {e}")

    def restore_all_paths(self, template_edit, data_edit, output_edit,
                          set_template, set_data, set_output):
        """
        恢复所有上次使用的路径到输入框

        Args:
            template_edit: 模板路径输入框
            data_edit: 数据路径输入框
            output_edit: 输出路径输入框
            set_template: 设置模板路径的方法
            set_data: 设置数据路径的方法
            set_output: 设置输出路径的方法
        """
        paths = self.load_last_paths()

        if paths.get('template_dir') and self.config_manager.validate_path(paths['template_dir']):
            self.restore_path(paths['template_dir'], template_edit, set_template)

        if paths.get('data_dir') and self.config_manager.validate_path(paths['data_dir']):
            self.restore_path(paths['data_dir'], data_edit, set_data)

        if paths.get('output_dir') and self.config_manager.validate_path(paths['output_dir']):
            self.restore_path(paths['output_dir'], output_edit, set_output)

    def save_path_on_select(self, file_type, file_path):
        """
        选择文件后自动保存路径

        Args:
            file_type: 'template', 'data', 或 'output'
            file_path: 选中的文件/目录路径
        """
        try:
            kwargs = {f'{file_type}_dir': file_path}
            self.config_manager.update_last_paths(**kwargs)
        except Exception as e:
            logger.warning(f"保存路径失败: {e}")
