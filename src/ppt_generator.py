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
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any, Callable
import pandas as pd
from pathlib import Path

# 渲染引擎与文字特效是独立的库（libs/render_core、libs/fontfx），直接按包名导入
from render_core import RenderPrecheck, font_registry

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


@dataclass
class FailedRow:
    """单行生成失败记录（批量容错：跳过并汇报，可单独重试）"""
    index: int      # DataFrame 行标签（0 起）
    row_number: int # 用户可见行号（Excel 含表头行，标签+2）
    name: str       # 姓名标识（无姓名列时为 "output"）
    stage: str      # 失败阶段：export(导出图片) / fallback(降级PPT) / ppt(生成PPT)
    error: str      # 错误摘要


@dataclass
class BatchResult:
    """批量生成结果摘要（成功数 + 失败明细 + 降级明细 + 模板预检提示）"""
    success_count: int = 0
    failed_rows: List[FailedRow] = field(default_factory=list)
    total: int = 0
    precheck_issues: List[Any] = field(default_factory=list)  # List[PrecheckIssue]
    # 图片导出失败、降级输出为 PPT 的行（仍计入成功数，但单独提示用户）
    degraded_rows: List[FailedRow] = field(default_factory=list)
    cancelled: bool = False     # 是否被用户中途取消（结果为已完成部分的摘要）

    def __int__(self) -> int:
        # 兼容旧调用方把返回值当成功计数使用
        return self.success_count


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

    def _precheck_template(self) -> List[Any]:
        """
        对当前模板跑一次渲染预检（模板级，与行数据无关）

        Returns:
            List[PrecheckIssue]: 预检问题列表（空列表 = 模板完全支持；
            预检本身失败时返回空列表并不影响生成）
        """
        if not self.template_path:
            return []
        try:
            installed = font_registry.available_typefaces()
            try:
                try:
                    from src.core.utils.font_checker import FontChecker
                except ImportError:
                    from core.utils.font_checker import FontChecker
                installed |= FontChecker().get_installed_fonts()
            except Exception as e:
                logger.debug(f"系统字体检查跳过（仅用注册字体集）: {e}")
            return list(RenderPrecheck(installed_fonts=installed,
                                       font_resolver=font_registry.is_resolvable)
                        .check_pptx(self.template_path))
        except Exception as e:
            logger.debug(f"模板预检失败（不影响生成）: {e}")
            return []

    def batch_generate(self, output_dir: str, progress_callback=None,
                      generate_images: bool = False, image_format: str = "PNG",
                      image_quality: float = 1.0,
                      row_indices: Optional[List[int]] = None,
                      cancel_check: Optional[Callable[[], bool]] = None,
                      filename_prefix: str = "", filename_suffix: str = "") -> BatchResult:
        """
        批量生成PPT或图片（GUI调用接口）

        单行失败不再静默跳过：记录到 result.failed_rows 并继续后续行，
        可通过 row_indices 只重跑指定行。仅环境级致命错误（如 Office 缺失
        的 RuntimeError）仍然整批中止。

        Args:
            output_dir: 输出目录
            progress_callback: 进度回调函数，接收(current, total, message)参数
            generate_images: 是否生成图片
            image_format: 图片格式
            image_quality: 图片缩放倍数 (1.0=原始, 2.0=增强, 3.0=高质量)
            row_indices: 只生成指定的 DataFrame 行标签列表（失败行重试用）；
                         None = 全部行
            cancel_check: 可选回调，返回 True 时在下一行开始前中止批量，
                          返回已完成部分的摘要（result.cancelled=True）
            filename_prefix: 输出文件名前缀（可选，默认空＝与原来完全一致）
            filename_suffix: 输出文件名后缀（可选，默认空；加在扩展名之前）

        Returns:
            BatchResult: success_count / failed_rows / degraded_rows / total /
                         precheck_issues / cancelled
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
            failed_rows: List[FailedRow] = []
            degraded_rows: List[FailedRow] = []
            data_count = len(self.data_reader.data)

            # 模板级渲染预检（只跑一次，结果随 BatchResult 上报界面）
            precheck_issues = self._precheck_template()

            # 计算本次要处理的行集合与总数（row_indices 用于失败行重试）
            selected = None if row_indices is None else {int(i) for i in row_indices}
            if selected is None:
                to_process = data_count
            else:
                to_process = sum(1 for i in self.data_reader.data.index if i in selected)
            processed_rows = 0
            cancelled = False

            for index, row in self.data_reader.data.iterrows():
                if selected is not None and index not in selected:
                    continue
                # 每行开始前检查取消请求（协作式取消，最多再跑完当前一行）
                if cancel_check is not None and cancel_check():
                    cancelled = True
                    logger.info(f"批量生成已取消（已完成 {processed_rows}/{to_process}）")
                    break
                processed_rows += 1

                # 调用进度回调
                if progress_callback:
                    progress_callback(processed_rows, to_process,
                                      f"正在生成第 {processed_rows}/{to_process} 个文件...")

                # 用户可见行号（Excel 首行为表头，DataFrame 标签 0 对应第 2 行）
                try:
                    row_number = int(index) + 2
                except (TypeError, ValueError):
                    row_number = processed_rows + 1

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
                # 优先使用含"姓名"的列内容作为文件名标识，找不到则用 output 兜底
                row_name = 'output'
                try:
                    name_value = 'output'
                    # 优先精确匹配"姓名"列；否则取第一个包含"姓名"二字的列
                    if '姓名' in row:
                        name_value = row['姓名']
                    else:
                        for col in row.index:
                            if isinstance(col, str) and '姓名' in col:
                                name_value = row[col]
                                break
                    # 处理 NaN/空值
                    if name_value is None or (isinstance(name_value, float) and pd.isna(name_value)) or str(name_value).strip() == '':
                        name_value = 'output'
                    else:
                        name_value = str(name_value).strip()
                    row_name = name_value
                    filename = f"{index+1}_{name_value}"
                except (KeyError, AttributeError, TypeError) as e:
                    logger.debug(f"获取行数据失败：{str(e)}，使用默认文件名")
                    filename = f"{index+1}_output"

                # 确保文件名有效（CLI 可选文件名前后缀，默认空＝与原来完全一致）
                if filename_prefix or filename_suffix:
                    filename = f"{filename_prefix}{filename}{filename_suffix}"
                filename = self._sanitize_filename(filename)

                def _record_failure(stage: str, error: str):
                    """记录失败行并通过进度回调告知（不中断批次）"""
                    failed_rows.append(FailedRow(
                        index=index, row_number=row_number,
                        name=row_name, stage=stage, error=str(error)[:300]))
                    logger.error(f"第 {row_number} 行（{row_name}）{stage} 阶段失败: {error}")
                    if progress_callback:
                        progress_callback(processed_rows, to_process,
                                          f"第 {row_number} 行（{row_name}）失败，已跳过: {str(error)[:120]}")

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
                    # 生成图片：export_to_images 内部按需写临时文件（COM 引擎）
                    # 或直接内存渲染（内置引擎），这里不必再另存一份临时 PPT
                    try:
                        self.template_processor.replace_placeholders(data)
                        os.makedirs(output_dir, exist_ok=True)
                        base_filename = os.path.splitext(filename)[0]
                        generated_files = self.template_processor.export_to_images(
                            output_dir=output_dir,
                            format=image_format,
                            quality=image_quality,
                            base_filename=base_filename
                        )

                        if generated_files:
                            generated_count += 1
                            logger.info(f"生成图片成功: {filename} -> {len(generated_files)} 张图片")
                        else:
                            reason = getattr(self.template_processor, 'last_export_error', None)
                            _record_failure("export", reason or "渲染引擎未产出图片（详见日志）")
                    except RuntimeError as e:
                        # Office 未安装（office 策略）等环境级致命错误，整批中止
                        logger.error(f"导出图片失败: {e}")
                        raise
                    except Exception as e:
                        logger.error(f"导出图片时出错: {str(e)}")
                        # 导出失败时降级输出 PPT，并单独记录为「降级行」提示用户
                        output_path = os.path.join(output_dir, f"{filename}.pptx")
                        try:
                            self.template_processor.replace_placeholders(data)
                            self.template_processor.save_presentation(output_path)
                            generated_count += 1
                            degraded_rows.append(FailedRow(
                                index=index, row_number=row_number, name=row_name,
                                stage='degraded', error=str(e)[:300]))
                            logger.warning(f"第 {row_number} 行（{row_name}）图片导出失败，"
                                           f"已降级输出 PPT: {output_path}（原因: {e}）")
                        except Exception as e2:
                            _record_failure("fallback", e2)
                else:
                    # 生成PPT
                    output_path = os.path.join(output_dir, f"{filename}.pptx")
                    try:
                        self.template_processor.replace_placeholders(data)
                        self.template_processor.save_presentation(output_path)
                        generated_count += 1
                        logger.info(f"生成PPT: {output_path}")
                    except Exception as e:
                        _record_failure("ppt", e)
                        # 注意：失败后不能 continue，必须继续执行下方的 restore_template，
                        # 否则下一行会在被污染的内存模板上生成，导致数据串行（H2 修复）

                # 恢复模板原状以准备下一次生成（仅当还有后续待处理行）。
                # 用「恢复幻灯片 XML」代替重新加载：同一 Presentation 逐行复用，
                # 省去每行重新解压解析整个包，渲染器的字体/图片缓存也跨行保留。
                if processed_rows < to_process:
                    self.template_processor.restore_template()

            # 恢复模板原状，准备下一次批量生成
            self.template_processor.restore_template()

            result = BatchResult(
                success_count=generated_count,
                failed_rows=failed_rows,
                total=to_process,
                precheck_issues=precheck_issues,
                degraded_rows=degraded_rows,
                cancelled=cancelled,
            )
            if failed_rows:
                logger.warning(f"批量生成完成：成功 {generated_count}/{to_process}，"
                               f"失败 {len(failed_rows)} 行（行号: "
                               f"{', '.join(str(f.row_number) for f in failed_rows[:10])}"
                               f"{'...' if len(failed_rows) > 10 else ''}）")
            else:
                logger.info(f"批量生成完成，共生成 {generated_count} 个文件")
            return result

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
        获取图片导出引擎信息

        纯 Python 分支：图片导出不再依赖 Office，COM 检测已禁用

        Returns:
            str: 引擎信息字符串
        """
        return "纯 Python 渲染引擎（Pillow，无需 Office/WPS）"

    def is_office_available(self) -> bool:
        """
        检查图片导出是否可用

        纯 Python 分支：渲染引擎内置，始终可用，COM 检测已禁用

        Returns:
            bool: 始终 True
        """
        return True

    def close_template_processor(self):
        """关闭当前模板处理器，释放资源"""
        try:
            if self.template_processor:
                self.template_processor.close()
                logger.info("模板处理器已关闭")
        except Exception as e:
            logger.error(f"关闭模板处理器时出错: {str(e)}")

    def _initialize_template_processor(self):
        """
        重新初始化模板处理器

        在批量生成结束后调用，确保下次生成使用全新的模板处理器实例，
        避免旧状态（已替换的占位符等）残留。
        """
        try:
            self.close_template_processor()
            self.template_processor = self.factory.create_pptx_processor()
            self.template_loaded = False
            self.template_path = None
            logger.info("模板处理器已重新初始化")
        except Exception as e:
            logger.error(f"重新初始化模板处理器时出错: {str(e)}")

    def close(self):
        """关闭PPT生成器，释放资源"""
        self.close_template_processor()

    def __enter__(self):
        """上下文管理器入口"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器出口"""
        self.close()
        return False
