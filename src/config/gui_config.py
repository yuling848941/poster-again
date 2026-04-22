"""
GUI 配置模块
负责管理 GUI 相关的配置，包括承保趸期用户输入和 GUI 状态保存
"""

import os
import json
import logging
from typing import Dict, Any
from datetime import datetime

logger = logging.getLogger(__name__)


class GUIConfigMixin:
    """GUI 配置混合类，提供 GUI 相关的配置管理功能"""

    # ==================== 承保趸期数据管理 ====================

    def save_chengbao_user_inputs(self, user_inputs: Dict[int, str]) -> bool:
        """
        保存用户输入的承保趸期数据

        Args:
            user_inputs: 用户输入数据字典，键为行号，值为输入值

        Returns:
            bool: 保存是否成功
        """
        try:
            # 保存到配置中
            self.config['chengbao_user_inputs'] = user_inputs

            # 保存配置
            self.save_config()

            logger.info(f"[承保趸期] 已保存用户输入数据，共 {len(user_inputs)} 条")
            return True

        except Exception as e:
            logger.error(f"保存承保趸期用户输入失败：{str(e)}")
            return False

    def load_chengbao_user_inputs(self) -> Dict[int, str]:
        """
        加载用户输入的承保趸期数据

        Returns:
            Dict[int, str]: 用户输入数据字典，键为行号，值为输入值
        """
        try:
            # 从配置中获取数据
            user_inputs = self.config.get('chengbao_user_inputs', {})

            if user_inputs:
                logger.info(f"[承保趸期] 已加载用户输入数据，共 {len(user_inputs)} 条")
            else:
                logger.info("[承保趸期] 未找到用户输入数据")

            return user_inputs

        except Exception as e:
            logger.error(f"加载承保趸期用户输入失败：{str(e)}")
            return {}

    def clear_chengbao_user_inputs(self) -> bool:
        """
        清除承保趸期用户输入数据

        Returns:
            bool: 清除是否成功
        """
        try:
            if 'chengbao_user_inputs' in self.config:
                del self.config['chengbao_user_inputs']
                self.save_config()
                logger.info("[承保趸期] 已清除用户输入数据")
            return True

        except Exception as e:
            logger.error(f"清除承保趸期用户输入失败：{str(e)}")
            return False

    # ==================== GUI 配置保存/加载 ====================

    def save_gui_config(self, main_window, file_path: str) -> bool:
        """
        保存 GUI 配置，包括占位符匹配规则

        Args:
            main_window: 主窗口对象
            file_path: 配置文件保存路径

        Returns:
            bool: 保存是否成功
        """
        try:
            # 构建配置数据结构
            config_data = {
                "version": "1.0.0",
                "created_at": datetime.now().isoformat(),
                "metadata": {
                    "type": "gui_config",
                    "description": "GUI 配置和占位符匹配规则"
                },
                "matching_rules": {},
                "template_info": {},
                "data_info": {}
            }

            # 保存匹配规则
            if hasattr(main_window, 'worker_thread') and main_window.worker_thread:
                config_data["matching_rules"] = dict(main_window.worker_thread.matching_rules)

            # 保存模板和数据信息
            if hasattr(main_window, 'template_path') and main_window.template_path:
                config_data["template_info"]["path"] = main_window.template_path

            if hasattr(main_window, 'data_path') and main_window.data_path:
                config_data["data_info"]["path"] = main_window.data_path

            # 保存 Excel 列信息
            if hasattr(main_window, 'data_frame') and main_window.data_frame is not None:
                config_data["data_info"]["columns"] = list(main_window.data_frame.columns)

            # 保存承保趸期相关配置
            if hasattr(main_window, 'chengbao_user_inputs'):
                config_data["chengbao_user_inputs"] = main_window.chengbao_user_inputs

            # 确保文件扩展名为.pptcfg
            if not file_path.endswith('.pptcfg'):
                file_path += '.pptcfg'

            # 保存配置文件
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)

            logger.info(f"成功保存 GUI 配置到：{file_path}")
            logger.info(f"保存了 {len(config_data['matching_rules'])} 条匹配规则")
            return True

        except Exception as e:
            logger.error(f"保存 GUI 配置失败：{str(e)}")
            return False

    def load_gui_config(self, main_window, file_path: str) -> bool:
        """
        加载 GUI 配置，包括占位符匹配规则

        Args:
            main_window: 主窗口对象
            file_path: 配置文件路径

        Returns:
            bool: 加载是否成功
        """
        try:
            if not os.path.exists(file_path):
                logger.error(f"配置文件不存在：{file_path}")
                return False

            # 加载配置文件
            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)

            # 验证配置文件格式
            if "matching_rules" not in config_data:
                logger.error("配置文件格式不正确，缺少 matching_rules")
                return False

            # 恢复匹配规则
            matching_rules = config_data["matching_rules"]
            if matching_rules and hasattr(main_window, 'worker_thread') and main_window.worker_thread:
                main_window.worker_thread.set_matching_rules(matching_rules)
                logger.info(f"已恢复 {len(matching_rules)} 条匹配规则")

                # 同步到 PPT 生成器
                if hasattr(main_window.worker_thread, 'get_ppt_generator'):
                    ppt_generator = main_window.worker_thread.get_ppt_generator()
                    if ppt_generator:
                        for placeholder, column in matching_rules.items():
                            ppt_generator.add_matching_rule(placeholder, column)

            # 恢复承保趸期用户输入数据
            if "chengbao_user_inputs" in config_data and config_data["chengbao_user_inputs"]:
                main_window.chengbao_user_inputs = config_data["chengbao_user_inputs"]
                logger.info(f"已恢复承保趸期用户输入数据：{len(config_data['chengbao_user_inputs'])} 条")

            # 更新界面显示
            if hasattr(main_window, 'populate_match_table'):
                main_window.populate_match_table()

            logger.info(f"成功加载 GUI 配置：{file_path}")
            return True

        except Exception as e:
            logger.error(f"加载 GUI 配置失败：{str(e)}")
            return False
