"""
PPT生成器模块
整合模板处理、数据读取和精确匹配功能，生成最终的PPT文件

注意：此版本使用纯 Python 的 python-pptx 库，无需 Microsoft Office 或 WPS
"""

import os
import re
import logging
import shutil
import tempfile
import time
from typing import Dict, List, Optional, Any, Callable
import pandas as pd
from pathlib import Path

# 导入自定义模块
try:
    from src.core.processors.pptx_processor import PPTXProcessor
    from src.core.interfaces.template_processor import ITemplateProcessor
    from src.core.factory.office_suite_factory import OfficeSuiteFactory
    from src.data_reader import DataReader
    from src.exact_matcher import ExactMatcher
except ImportError:
    from core.processors.pptx_processor import PPTXProcessor
    from core.interfaces.template_processor import ITemplateProcessor
    from core.factory.office_suite_factory import OfficeSuiteFactory
    from data_reader import DataReader
    from exact_matcher import ExactMatcher

# 设置日志 - 使用模块级 logger，避免 basicConfig 全局配置冲突
logger = logging.getLogger(__name__)


class PPTGenerator:
    """PPT生成器，整合所有功能模块生成PPT文件"""

    def __init__(self):
        """初始化PPT生成器"""
        self.data_reader = DataReader()
        self.exact_matcher = ExactMatcher()

        # 创建模板处理器（纯 Python）
        self.factory = OfficeSuiteFactory()
        self.template_processor: Optional[ITemplateProcessor] = self.factory.create_pptx_processor()

        # 生成状态
        self.template_loaded = False
        self.data_loaded = False
        self.matching_completed = False
        self.template_path = None

        # 进度回调函数
        self.progress_callback: Optional[Callable[[int, str], None]] = None

        # 贺语模板
        self.message_template = ""

        logger.info("PPT生成器初始化完成，使用纯 Python PPTX 处理器（无需 Office/WPS）")

    def set_progress_callback(self, callback: Optional[Callable[[int, str], None]]):
        """
        设置进度回调函数
        
        Args:
            callback: 回调函数，接收进度百分比和状态信息
        """
        self.progress_callback = callback
        
    def _update_progress(self, progress: int, message: str):
        """
        更新进度
        
        Args:
            progress: 进度百分比 (0-100)
            message: 状态信息
        """
        if self.progress_callback:
            self.progress_callback(progress, message)
        
    def load_template(self, template_path: str) -> bool:
        """
        加载PPT模板文件
        
        Args:
            template_path: PPT模板文件路径
            
        Returns:
            bool: 加载成功返回True，失败返回False
        """
        self._update_progress(5, "正在加载PPT模板...")
        
        try:
            success = self.template_processor.load_template(template_path)
            if success:
                self.template_loaded = True
                self.template_path = template_path
                self._update_progress(10, "PPT模板加载成功")
                logger.info(f"成功加载PPT模板: {template_path}")
            else:
                self._update_progress(0, "PPT模板加载失败")
                logger.error(f"加载PPT模板失败: {template_path}")
            return success
        except Exception as e:
            self._update_progress(0, f"PPT模板加载出错: {str(e)}")
            logger.error(f"加载模板出错: {str(e)}")
            return False
    
    def load_data(self, data_path: str, sheet_name: Optional[str] = None) -> bool:
        """
        加载数据文件
        
        Args:
            data_path: 数据文件路径
            sheet_name: 工作表名称
            
        Returns:
            bool: 加载成功返回True，失败返回False
        """
        self._update_progress(15, "正在加载数据文件...")
        
        try:
            success = self.data_reader.load_excel(data_path, sheet_name)
            if success:
                self.data_loaded = True
                self.exact_matcher.set_data(self.data_reader.data)
                self._update_progress(20, "数据文件加载成功")
                logger.info(f"成功加载数据: {data_path}")
            else:
                self._update_progress(0, "数据文件加载失败")
                logger.error(f"加载数据失败: {data_path}")
            return success
        except Exception as e:
            self._update_progress(0, f"数据文件加载出错: {str(e)}")
            logger.error(f"加载数据出错: {str(e)}")
            return False
    
    def extract_placeholders(self) -> List[str]:
        """
        从模板中提取占位符
        
        Returns:
            List[str]: 占位符列表
        """
        if not self.template_loaded:
            logger.error("模板未加载")
            return []
            
        self._update_progress(25, "正在提取模板占位符...")
        
        try:
            # 记录使用的处理器类型
            processor_type = type(self.template_processor).__name__
            logger.info(f"使用处理器: {processor_type}")

            # 使用PPTXProcessor的find_placeholders方法
            placeholders = self.template_processor.find_placeholders()
            self.exact_matcher.set_template_placeholders(placeholders)
            self._update_progress(30, f"提取到 {len(placeholders)} 个占位符")
            logger.info(f"提取到占位符: {placeholders}")
            return placeholders
        except Exception as e:
            self._update_progress(0, f"提取占位符出错: {str(e)}")
            logger.error(f"提取占位符出错: {str(e)}")
            return []
    
    def auto_match_placeholders(self) -> Dict[str, str]:
        """
        自动匹配占位符与数据列

        Returns:
            Dict[str, str]: 匹配规则
        """
        if not self.data_loaded:
            logger.error("数据未加载")
            return {}

        if not self.template_loaded:
            logger.error("模板未加载")
            return {}

        # 首先提取模板占位符
        placeholders = self.extract_placeholders()
        if not placeholders:
            logger.error("模板中未找到占位符")
            return {}

        self._update_progress(35, "正在自动匹配占位符...")

        try:
            matching_rules = self.exact_matcher.auto_match_placeholders()
            if matching_rules:
                self.matching_completed = True
                self._update_progress(40, f"自动匹配完成，匹配了 {len(matching_rules)} 个占位符")
                logger.info(f"匹配规则: {matching_rules}")
            else:
                self._update_progress(0, "自动匹配失败，未找到任何匹配")
                logger.warning("自动匹配失败，未找到任何匹配")
            return matching_rules
        except Exception as e:
            self._update_progress(0, f"自动匹配出错: {str(e)}")
            logger.error(f"自动匹配出错: {str(e)}")
            return {}

    def set_matching_rules(self, matching_rules: Dict[str, str]):
        """
        设置占位符与数据列的匹配规则
        
        Args:
            matching_rules: 匹配规则字典，键为占位符，值为数据列名
        """
        self.exact_matcher.set_matching_rules(matching_rules)
        self.matching_completed = True
        logger.info(f"手动设置匹配规则: {matching_rules}")
    
    def add_matching_rule(self, placeholder: str, column: str):
        """
        添加单个匹配规则
        
        Args:
            placeholder: 占位符
            column: 数据列名
        """
        self.exact_matcher.set_matching_rule(placeholder, column)
        self.matching_completed = True
        logger.info(f"添加匹配规则: '{placeholder}' -> '{column}'")
    
    def get_matching_rules(self) -> Dict[str, str]:
        """
        获取当前的匹配规则
        
        Returns:
            Dict[str, str]: 匹配规则字典
        """
        return self.exact_matcher.get_matching_rules()
    
    def get_data_columns(self) -> List[str]:
        """
        获取数据文件的列名
        
        Returns:
            List[str]: 列名列表
        """
        if not self.data_loaded:
            return []
        return self.data_reader.get_column_names()
    
    def clear_matching_rules(self):
        """
        清除匹配规则
        """
        self.exact_matcher.clear_matching_rules()
        self.matching_completed = False
        logger.info("匹配规则已清除")

    def set_message_template(self, template: str):
        """
        设置贺语模板

        Args:
            template: 贺语模板文本，支持 {{表头}} 格式的占位符
        """
        self.message_template = template
        logger.info(f"已设置贺语模板，长度: {len(template)} 字符")

    def generate_message(self, row_data: pd.Series) -> str:
        """
        根据模板和数据生成贺语

        Args:
            row_data: 一行数据（pandas Series）

        Returns:
            str: 生成的贺语文本
        """
        if not self.message_template:
            return ""

        message = self.message_template

        # 使用正则表达式查找所有 {{表头}} 格式的占位符
        placeholders = re.findall(r'\{\{(.*?)\}\}', message)

        # 替换每个占位符
        for placeholder in placeholders:
            column_name = placeholder.strip()
            if column_name in row_data:
                value = str(row_data[column_name])
                message = message.replace(f'{{{{{column_name}}}}}', value)
            else:
                logger.warning(f"贺语模板中的占位符 '{{{{{column_name}}}}}' 未在数据中找到对应列")
                # 替换为空字符串
                message = message.replace(f'{{{{{column_name}}}}}', '')

        return message

    def save_message_file(self, message: str, output_path: str) -> bool:
        """
        保存贺语到txt文件

        Args:
            message: 贺语文本
            output_path: 输出文件路径（与贺报同名，扩展名为.txt）

        Returns:
            bool: 保存成功返回True
        """
        try:
            # 确保输出目录存在
            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)

            # 写入文件（UTF-8编码）
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(message)

            logger.info(f"贺语文件已保存: {output_path}")
            return True

        except Exception as e:
            logger.error(f"保存贺语文件失败 {output_path}: {str(e)}")
            return False

    def batch_generate(self, output_dir: str, progress_callback=None, 
                      generate_images: bool = False, image_format: str = "PNG",
                      image_quality: str = "原始大小") -> int:
        """
        批量生成PPT或图片（GUI调用接口）
        
        Args:
            output_dir: 输出目录
            progress_callback: 进度回调函数，接收(current, total, message)参数
            generate_images: 是否生成图片
            image_format: 图片格式
            image_quality: 图片质量
            
        Returns:
            int: 成功生成的文件数量
        """
        if not self.template_loaded:
            raise ValueError("模板未加载，请先调用load_template()")
        
        if not self.data_loaded:
            raise ValueError("数据未加载，请先调用load_data()")
            
        if not self.matching_completed:
            logger.warning("匹配未完成，将使用自动匹配")
            self.auto_match_placeholders()
        
        try:
            # 确保输出目录存在
            os.makedirs(output_dir, exist_ok=True)
            
            # 获取匹配规则
            matching_rules = self.exact_matcher.get_matching_rules()
            if not matching_rules:
                raise ValueError("没有有效的匹配规则")
            
            generated_count = 0
            data_count = len(self.data_reader.data)
            
            for index, row in self.data_reader.data.iterrows():
                # 调用进度回调
                if progress_callback:
                    progress_callback(index + 1, data_count, f"正在生成第 {index + 1}/{data_count} 个文件...")
                
                # 准备数据（使用 ExactMatcher 以应用文本增加规则）
                data = self.exact_matcher.get_all_data_for_row(index)
                if not data:
                    data = {}
                    for placeholder, column in matching_rules.items():
                        if column in row:
                            data[placeholder] = str(row[column])
                        else:
                            data[placeholder] = ""
                
                # 生成文件名
                try:
                    filename = f"{index+1}_{row.get('姓名', 'output')}"
                except (KeyError, AttributeError, TypeError) as e:
                    logger.debug(f"获取行数据失败：{str(e)}，使用默认文件名")
                    filename = f"{index+1}_output"
                
                # 确保文件名有效
                filename = self._sanitize_filename(filename)

                # 生成贺语文件（如果设置了贺语模板）
                if self.message_template:
                    try:
                        message = self.generate_message(row)
                        if message:
                            # 构建贺语文件路径（去掉原有扩展名，添加.txt）
                            base_filename = os.path.splitext(filename)[0]
                            message_filename = f"{base_filename}.txt"
                            message_output_path = os.path.join(output_dir, message_filename)
                            self.save_message_file(message, message_output_path)
                    except Exception as e:
                        logger.error(f"生成贺语文件时出错: {str(e)}")

                if generate_images:
                    # 生成图片
                    temp_pptx = None
                    try:
                        # 使用PPTXProcessor的导出图片功能
                        # 先替换占位符
                        self.template_processor.replace_placeholders(data)
                        # 保存临时PPT（使用临时文件，不在输出目录）
                        temp_pptx = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False).name
                        self.template_processor.save_presentation(temp_pptx)
                        # 导出为图片到输出目录（不使用子文件夹）
                        os.makedirs(output_dir, exist_ok=True)

                        # 将质量文本转换为数值
                        quality_map = {
                            "原始大小": 1.0,
                            "1.5倍": 1.5,
                            "2倍": 2.0,
                            "3倍": 3.0,
                            "4倍": 4.0
                        }
                        quality_float = quality_map.get(image_quality, 1.0)

                        # 获取基础文件名（不含扩展名）
                        base_filename = os.path.splitext(filename)[0]
                        generated_files = self.template_processor.export_to_images(
                            output_dir=output_dir,
                            format=image_format,
                            quality=quality_float,
                            base_filename=base_filename
                        )

                        if generated_files:
                            generated_count += 1
                            logger.info(f"生成图片成功: {filename} -> {len(generated_files)} 张图片")
                        else:
                            logger.error(f"导出图片失败")
                    except RuntimeError as e:
                        # Office 未安装
                        logger.error(f"导出图片失败: {e}")
                        raise
                    except Exception as e:
                        logger.error(f"导出图片时出错: {str(e)}")
                        # 如果导出图片失败，尝试生成PPT作为备选
                        output_path = os.path.join(output_dir, f"{filename}.pptx")
                        try:
                            self.template_processor.replace_placeholders(data)
                            self.template_processor.save_presentation(output_path)
                            generated_count += 1
                            logger.info(f"生成PPT: {output_path}")
                        except Exception as e2:
                            logger.error(f"生成第 {index + 1} 个文件时出错: {str(e2)}")
                    finally:
                        # 删除临时PPT文件
                        if temp_pptx and os.path.exists(temp_pptx):
                            try:
                                os.remove(temp_pptx)
                                logger.debug(f"删除临时PPT文件: {temp_pptx}")
                            except Exception as e:
                                logger.warning(f"删除临时PPT文件失败: {e}")
                else:
                    # 生成PPT
                    output_path = os.path.join(output_dir, f"{filename}.pptx")
                    try:
                        self.template_processor.replace_placeholders(data)
                        self.template_processor.save_presentation(output_path)
                        generated_count += 1
                        logger.info(f"生成PPT: {output_path}")
                    except Exception as e:
                        logger.error(f"生成第 {index + 1} 个文件时出错: {str(e)}")
                        continue
                
                # 重新加载模板以准备下一次生成
                if index < data_count - 1:
                    self.template_processor.load_template(self.template_path)
            
            # 重新加载模板，准备下一次批量生成
            self.template_processor.load_template(self.template_path)
            
            logger.info(f"批量生成完成，共生成 {generated_count} 个文件")
            return generated_count
            
        except Exception as e:
            logger.error(f"批量生成时出错: {str(e)}")
            raise
    
    def _sanitize_filename(self, filename: str) -> str:
        """
        清理文件名，移除非法字符
        
        Args:
            filename: 原始文件名
            
        Returns:
            str: 清理后的文件名
        """
        # Windows文件名非法字符
        illegal_chars = '<>:"/\\|?*'
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        
        # 移除首尾空格和点
        filename = filename.strip('. ')
        
        # 如果文件名为空，使用默认值
        if not filename:
            filename = "output"
            
        return filename
    
    def get_office_info(self) -> str:
        """
        获取 Office 套件信息

        Returns:
            str: Office 套件信息字符串
        """
        try:
            try:
                from src.core.detectors.office_suite_detector import OfficeSuiteDetector
            except ImportError:
                from core.detectors.office_suite_detector import OfficeSuiteDetector
            return OfficeSuiteDetector.get_office_info()
        except Exception as e:
            logger.debug(f"获取 Office 信息时出错: {e}")
            return "无法检测 Office 套件"

    def is_office_available(self) -> bool:
        """
        检查是否有可用的 Office 套件

        Returns:
            bool: 是否有可用的 Office 套件
        """
        try:
            try:
                from src.core.detectors.office_suite_detector import OfficeSuiteDetector
            except ImportError:
                from core.detectors.office_suite_detector import OfficeSuiteDetector
            return OfficeSuiteDetector.is_office_available()
        except Exception as e:
            logger.debug(f"检查 Office 可用性时出错: {e}")
            return False

    def close(self):
        """关闭PPT生成器，释放资源"""
        try:
            if self.template_processor:
                self.template_processor.close()
                logger.info("PPT生成器资源已释放")
        except Exception as e:
            logger.error(f"关闭PPT生成器时出错: {str(e)}")

    def __enter__(self):
        """上下文管理器入口"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.close()
        return False
