import sys
import os
from PySide6.QtCore import QThread, Signal
try:
    from src.ppt_generator import PPTGenerator
    from src.config import ConfigManager
    from src.data_reader import DataReader
except ImportError:
    from ppt_generator import PPTGenerator
    from config_manager import ConfigManager
    from data_reader import DataReader


class PPTWorkerThread(QThread):
    """PPT生成工作线程，避免阻塞UI"""

    # 定义信号
    progress_updated = Signal(int, str)  # 进度值，状态文本
    log_message = Signal(str)  # 日志消息
    match_completed = Signal(list, list, dict)  # 占位符列表，数据列列表，匹配规则
    generation_completed = Signal(int)  # 生成的文件数量
    error_occurred = Signal(str)  # 错误消息

    def __init__(self, office_manager=None, main_window=None):
        super().__init__()
        self.template_path = ""
        self.data_path = ""
        self.output_path = ""
        self.matching_rules = {}
        self.text_addition_rules = {}  # 新增：文本增加规则
        self.config_manager = ConfigManager()
        # office_manager 参数已废弃，保留用于向后兼容
        self.office_manager = None
        self.main_window = main_window  # 保存MainWindow引用以获取用户输入数据
        # PPTGenerator现在通过get_ppt_generator()延迟初始化
        self.task_type = ""  # 当前任务类型：auto_match 或 batch_generate
        self.generate_images = False  # 新增：是否生成图片
        self.image_format = "PNG"  # 新增：图片格式
        self.image_quality = "原始大小"  # 新增：图片质量
        self.message_template = ""  # 新增：贺语模板
        
    def set_template_path(self, path):
        """设置模板文件路径"""
        self.template_path = path
        
    def set_data_path(self, path):
        """设置数据文件路径"""
        self.data_path = path
        
    def set_output_path(self, path):
        """设置输出路径"""
        self.output_path = path
        
    def set_matching_rules(self, rules):
        """设置匹配规则"""
        self.matching_rules = rules

    def clear_matching_rules(self):
        """清除匹配规则"""
        self.matching_rules = {}
        # 使用延迟获取PPTGenerator
        if hasattr(self, '_ppt_generator') and self._ppt_generator:
            self._ppt_generator.clear_matching_rules()

    def set_text_addition_rules(self, rules):
        """设置文本增加规则"""
        self.text_addition_rules = rules

    def set_image_generation_params(self, generate_images: bool, image_format: str = "PNG", image_quality: str = "原始大小"):
        """设置图片生成参数"""
        self.generate_images = generate_images
        self.image_format = image_format
        self.image_quality = image_quality

    def set_message_template(self, template: str):
        """设置贺语模板"""
        self.message_template = template
        
    def start_auto_match(self):
        """启动自动匹配任务"""
        self.task_type = "auto_match"
        self.start()
        
    def start_batch_generate(self):
        """启动批量生成任务"""
        self.task_type = "batch_generate"
        self.start()

    def get_ppt_generator(self):
        """延迟获取PPT生成器 - 性能优化"""
        if not hasattr(self, '_ppt_generator') or self._ppt_generator is None:
            # 使用纯 Python PPTX 处理器，无需 Office/WPS
            self._ppt_generator = PPTGenerator()
            self.log_message.emit("使用纯 Python PPTX 处理器（无需 Microsoft Office 或 WPS）")
        return self._ppt_generator

    def run(self):
        """线程运行方法"""
        if self.task_type == "auto_match":
            self.auto_match_placeholders()
        elif self.task_type == "batch_generate":
            self.batch_generate()
        
    def auto_match_placeholders(self):
        """自动匹配占位符"""
        try:
            # 获取PPT生成器（延迟初始化）
            ppt_generator = self.get_ppt_generator()

            self.log_message.emit("开始加载模板文件...")
            self.progress_updated.emit(10, "加载模板文件...")

            # 加载模板
            if not ppt_generator.load_template(self.template_path):
                self.error_occurred.emit("无法加载模板文件")
                return

            self.log_message.emit("模板文件加载成功")
            self.progress_updated.emit(30, "加载数据文件...")

            # 加载数据
            if not ppt_generator.load_data(self.data_path):
                self.error_occurred.emit("无法加载数据文件")
                return

            self.log_message.emit("数据文件加载成功")
            self.progress_updated.emit(50, "提取占位符...")

            # 提取占位符
            placeholders = ppt_generator.extract_placeholders()
            if not placeholders:
                self.error_occurred.emit("模板中未找到占位符")
                return

            self.log_message.emit(f"找到 {len(placeholders)} 个占位符")
            self.progress_updated.emit(70, "匹配数据列...")

            # 自动匹配
            ppt_generator.auto_match_placeholders()
            matching_rules = ppt_generator.get_matching_rules()
            
            self.progress_updated.emit(90, "准备匹配结果...")
            
            # 准备匹配结果
            data_columns = ppt_generator.get_data_columns()
            
            self.progress_updated.emit(100, "匹配完成")
            self.match_completed.emit(placeholders, data_columns, matching_rules)
            self.log_message.emit("自动匹配完成")
            
        except Exception as e:
            self.error_occurred.emit(f"自动匹配过程中发生错误: {str(e)}")
            
    def batch_generate(self):
        """批量生成PPT或图片"""
        try:
            # 获取PPT生成器（延迟初始化）
            ppt_generator = self.get_ppt_generator()

            file_type = "图片" if self.generate_images else "PPT"
            self.log_message.emit(f"开始批量生成{file_type}...")
            self.progress_updated.emit(5, "初始化...")

            # 加载模板
            if not ppt_generator.load_template(self.template_path):
                self.error_occurred.emit("无法加载模板文件")
                return

            self.progress_updated.emit(10, "加载数据...")

            # ⚠️ 重要：在加载数据前，先检查并保存承保趸期数据（如果有的话）
            # 这样可以避免配置管理器计算的数据被覆盖
            saved_chengbao_data = None
            if hasattr(ppt_generator, 'data_reader') and ppt_generator.data_reader:
                if "承保趸期数据" in ppt_generator.data_reader.data.columns:
                    saved_chengbao_data = ppt_generator.data_reader.data["承保趸期数据"].copy()
                    # 移除承保趸期数据相关日志，简化输出

            # 加载数据
            if not ppt_generator.load_data(self.data_path):
                self.error_occurred.emit("无法加载数据文件")
                return

            # ⚠️ 重要：加载数据后，恢复承保趸期数据（如果有保存的话）
            if saved_chengbao_data is not None:
                # 移除恢复日志，简化输出
                if "承保趸期数据" not in ppt_generator.data_reader.data.columns:
                    ppt_generator.data_reader.data["承保趸期数据"] = saved_chengbao_data

            # 计算承保趸期数据（如果需要）
            data_reader = DataReader()
            try:
                # 直接从ppt_generator的data_reader获取data
                # 注意：ppt_generator.load_data会将数据存储在data_reader.data中，而不是ppt_generator.data
                if hasattr(ppt_generator, 'data_reader') and ppt_generator.data_reader:
                    # 使用ppt_generator的data_reader实例
                    data_reader = ppt_generator.data_reader

                    # 检查是否有承保趸期数据列，如果没有则计算
                    if "承保趸期数据" not in data_reader.data.columns:
                        self.progress_updated.emit(12, "计算承保趸期数据...")

                        # ⚠️ 重要：使用DataReader的方法计算承保趸期数据（不弹出对话框）
                        # 注意：这里传递parent_widget=None，避免在batch_generate过程中弹出对话框
                        data_reader._process_chengbao_term_data(parent_widget=None)
                    else:
                        # 移除承保趸期数据存在日志，简化输出
                        # 检查是否有保存的用户输入数据，如果有则应用
                        try:
                            if self.main_window and hasattr(self.main_window, 'chengbao_user_inputs'):
                                user_inputs = getattr(self.main_window, 'chengbao_user_inputs', {})
                                if user_inputs:
                                    for row_index, value in user_inputs.items():
                                        if row_index < len(data_reader.data):
                                            data_reader.data.loc[row_index, "承保趸期数据"] = f"{value}年交SFYP"
                        except Exception as e:
                            # 移除错误日志，简化输出
                            pass

                    # 重要：确保所有匹配规则都可以正确设置
                    # 检查是否有关联到承保趸期数据的占位符
                    if self.matching_rules:
                        chengbao_related_placeholders = [
                            ph for ph, col in self.matching_rules.items()
                            if col == "承保趸期数据"
                        ]
                        if chengbao_related_placeholders:
                            # 移除检测日志，简化输出
                            # 确保承保趸期数据列存在
                            if "承保趸期数据" in data_reader.data.columns:
                                # 移除验证日志，简化输出
                                non_empty_count = (data_reader.data["承保趸期数据"] != "").sum()
                                # 移除计数日志，简化输出
                            else:
                                # 移除错误日志，简化输出
                                # 移除不匹配的占位符
                                self.matching_rules = {
                                    ph: col for ph, col in self.matching_rules.items()
                                    if ph not in chengbao_related_placeholders
                                }
            except Exception as e:
                # 移除承保趸期数据计算错误日志，简化输出
                pass

            self.progress_updated.emit(15, "应用匹配规则...")

            # 应用匹配规则
            if self.matching_rules:
                ppt_generator.set_matching_rules(self.matching_rules)
            else:
                ppt_generator.auto_match_placeholders()

            # 应用文本增加规则
            if self.text_addition_rules:
                for placeholder, rule in self.text_addition_rules.items():
                    ppt_generator.exact_matcher.set_text_addition_rule(
                        placeholder, rule.get("prefix", ""), rule.get("suffix", "")
                    )

            self.progress_updated.emit(20, f"开始生成{file_type}...")

            # 设置贺语模板
            if self.message_template:
                ppt_generator.set_message_template(self.message_template)
                self.log_message.emit("已启用贺语生成功能")

            # 定义进度回调
            def progress_callback(current, total, message):
                progress = int(20 + (current / total) * 75)  # 20-95%
                self.progress_updated.emit(progress, f"生成中: {message}")
                self.log_message.emit(f"进度: {current}/{total} - {message}")

            # 批量生成
            success_count = ppt_generator.batch_generate(
                self.output_path,
                progress_callback=progress_callback,
                generate_images=self.generate_images,
                image_format=self.image_format,
                image_quality=self.image_quality
            )

            self.progress_updated.emit(100, "生成完成")
            self.generation_completed.emit(success_count)
            self.log_message.emit(f"批量生成完成，成功生成 {success_count} 个{file_type}文件")

        except Exception as e:
            self.error_occurred.emit(f"批量生成过程中发生错误: {str(e)}")