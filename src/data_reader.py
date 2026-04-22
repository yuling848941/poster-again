"""
数据读取模块
用于读取和处理Excel数据文件
"""

import pandas as pd
import os
import gc
from typing import Dict, List, Optional, Any
import logging
try:
    from PySide6.QtWidgets import QWidget
except ImportError:
    QWidget = None
try:
    from .memory_management.memory_data_manager import MemoryDataManager
except ImportError:
    from memory_management.memory_data_manager import MemoryDataManager

# 设置日志 - 使用模块级 logger，避免 basicConfig 全局配置冲突
logger = logging.getLogger(__name__)


class DataReader:
    """数据读取器类，负责读取和处理Excel数据"""

    def __init__(self):
        """初始化数据读取器"""
        self.data = None
        self.file_path = None
        self.sheet_names = []
        self.current_sheet = None
        self.excel_file = None  # 添加Excel文件对象
        self.memory_manager = MemoryDataManager()  # 内存数据管理器
    
    @staticmethod
    def is_numeric_value(value):
        """检查值是否为数字"""
        if value is None or value == "":
            return False
        
        try:
            # 尝试转换为浮点数
            float(value)
            return True
        except (ValueError, TypeError):
            return False
    
    @staticmethod
    def format_number_with_separator(value):
        """为数字添加千位分隔符"""
        if value is None:
            return None
            
        try:
            # 转换为浮点数
            num = float(value)
            
            # 检查是否为整数
            if num.is_integer():
                return f"{int(num):,}"
            else:
                return f"{num:,}"
        except (ValueError, TypeError):
            return value
    
    def get_processed_data(self, file_path: str, use_thousand_separator: bool = True,
                          add_term_data: bool = True) -> pd.DataFrame:
        """
        获取处理后的数据（替代临时文件功能）

        Args:
            file_path: Excel文件路径
            use_thousand_separator: 是否使用千位分隔符
            add_term_data: 是否添加期趸数据

        Returns:
            处理后的DataFrame
        """
        try:
            return self.memory_manager.get_formatted_data_with_term_and_separator(
                file_path,
                use_thousand_separator=use_thousand_separator,
                add_term_data=add_term_data
            )
        except Exception as e:
            logger.error(f"获取处理数据失败: {str(e)}")
            # 降级为直接读取原始数据
            return self.memory_manager.load_excel_data(file_path)
        
    def load_excel(self, file_path: str, sheet_name: Optional[str] = None,
                   use_thousand_separator: bool = True, parent_widget=None) -> bool:
        """
        加载Excel文件（使用内存处理）

        Args:
            file_path: Excel文件路径
            sheet_name: 工作表名称，如果为None则加载第一个工作表
            use_thousand_separator: 是否使用千位分隔符格式化数字
            parent_widget: 父窗口组件，用于显示对话框（可选）

        Returns:
            bool: 加载成功返回True，失败返回False
        """
        try:
            # 关闭之前打开的文件
            self.close()

            if not os.path.exists(file_path):
                logger.error(f"文件不存在: {file_path}")
                return False

            # 使用内存处理获取数据
            try:
                processed_data = self.get_processed_data(
                    file_path,
                    use_thousand_separator=use_thousand_separator,
                    add_term_data=True  # 默认添加期趸数据
                )
            except Exception as e:
                logger.warning(f"内存处理失败，使用原始数据: {str(e)}")
                processed_data = self.memory_manager.load_excel_data(file_path)

            # 如果指定了工作表，需要重新加载特定工作表
            if sheet_name is not None:
                self.data = self.memory_manager.load_excel_data(file_path, sheet_name=sheet_name)
                # 如果需要千位分隔符，再次处理
                if use_thousand_separator:
                    self.data = self.memory_manager.formatter.add_thousands_separator(self.data)
                # 如果需要期趸数据，再次处理
                if "缴费年期（主险） ID" in self.data.columns:
                    self.data = self.memory_manager.formatter.calculate_term_data(
                        self.data,
                        source_column="缴费年期（主险） ID",
                        target_column="期趸数据"
                    )
                # 检查并添加承保趸期数据
                self._process_chengbao_term_data(parent_widget)
            else:
                self.data = processed_data
                # 检查并添加承保趸期数据
                self._process_chengbao_term_data(parent_widget)

            # 设置文件信息
            self.file_path = file_path
            self.current_sheet = sheet_name or "默认工作表"

            # 获取工作表信息
            try:
                temp_excel = pd.ExcelFile(file_path)
                self.sheet_names = temp_excel.sheet_names
                temp_excel.close()
            except Exception as e:
                logger.debug(f"获取工作表信息失败：{str(e)}")
                self.sheet_names = [self.current_sheet]

            logger.info(f"成功加载Excel文件: {file_path}")
            if use_thousand_separator:
                logger.info("使用千位分隔符格式化")
            if "期趸数据" in self.data.columns:
                logger.info("已添加期趸数据列")
            logger.info(f"数据行数: {len(self.data)}, 列数: {len(self.data.columns)}")

            return True

        except Exception as e:
            logger.error(f"加载Excel文件失败: {str(e)}")
            self.close()
            return False
    
    def close(self):
        """关闭Excel文件，释放资源"""
        if self.excel_file is not None:
            try:
                self.excel_file.close()
                logger.debug("Excel文件已关闭")
            except Exception as e:
                logger.warning(f"关闭Excel文件时出错: {str(e)}")
            finally:
                self.excel_file = None
        
        # 确保所有临时文件都被关闭
        try:
            # 强制垃圾回收，释放可能的文件句柄
            gc.collect()
        except Exception as e:
            logger.debug(f"垃圾回收时出错：{str(e)}")
    
    def get_sheet_names(self) -> List[str]:
        """
        获取所有工作表名称
        
        Returns:
            List[str]: 工作表名称列表
        """
        return self.sheet_names
    
    def switch_sheet(self, sheet_name: str) -> bool:
        """
        切换到指定工作表
        
        Args:
            sheet_name: 工作表名称
            
        Returns:
            bool: 切换成功返回True，失败返回False
        """
        if sheet_name not in self.sheet_names:
            logger.error(f"工作表 '{sheet_name}' 不存在")
            return False
            
        try:
            self.data = pd.read_excel(self.file_path, sheet_name=sheet_name)
            self.current_sheet = sheet_name
            logger.info(f"切换到工作表: {sheet_name}")
            return True
        except Exception as e:
            logger.error(f"切换工作表失败: {str(e)}")
            return False
    
    def get_column_names(self) -> List[str]:
        """
        获取当前工作表的列名
        
        Returns:
            List[str]: 列名列表
        """
        if self.data is None:
            return []
        return list(self.data.columns)
    
    def get_data_preview(self, num_rows: int = 5) -> pd.DataFrame:
        """
        获取数据预览
        
        Args:
            num_rows: 预览行数
            
        Returns:
            pd.DataFrame: 预览数据
        """
        if self.data is None:
            return pd.DataFrame()
        return self.data.head(num_rows)
    
    def get_column(self, column_name: str) -> List[Any]:
        """
        获取指定列的所有值（别名方法）
        
        Args:
            column_name: 列名
            
        Returns:
            List[Any]: 列值列表
        """
        return self.get_column_values(column_name)
    
    def get_column_values(self, column_name: str) -> List[Any]:
        """
        获取指定列的所有值
        
        Args:
            column_name: 列名
            
        Returns:
            List[Any]: 列值列表
        """
        if self.data is None or column_name not in self.data.columns:
            return []
        return self.data[column_name].tolist()
    
    def get_unique_values(self, column_name: str) -> List[Any]:
        """
        获取指定列的唯一值
        
        Args:
            column_name: 列名
            
        Returns:
            List[Any]: 唯一值列表
        """
        if self.data is None or column_name not in self.data.columns:
            return []
        return self.data[column_name].unique().tolist()
    
    def filter_data(self, column_name: str, value: Any) -> pd.DataFrame:
        """
        根据列值筛选数据
        
        Args:
            column_name: 列名
            value: 筛选值
            
        Returns:
            pd.DataFrame: 筛选后的数据
        """
        if self.data is None or column_name not in self.data.columns:
            return pd.DataFrame()
            
        try:
            return self.data[self.data[column_name] == value]
        except Exception as e:
            logger.error(f"筛选数据失败: {str(e)}")
            return pd.DataFrame()
    
    def get_row_count(self) -> int:
        """
        获取数据行数
        
        Returns:
            int: 行数
        """
        if self.data is None:
            return 0
        return len(self.data)
    
    def get_column_count(self) -> int:
        """
        获取数据列数
        
        Returns:
            int: 列数
        """
        if self.data is None:
            return 0
        return len(self.data.columns)
    
    def get_data_info(self) -> Dict[str, Any]:
        """
        获取数据基本信息
        
        Returns:
            Dict[str, Any]: 数据信息字典
        """
        if self.data is None:
            return {
                "file_path": None,
                "current_sheet": None,
                "sheet_names": [],
                "row_count": 0,
                "column_count": 0,
                "column_names": [],
                "data_types": {}
            }
            
        return {
            "file_path": self.file_path,
            "current_sheet": self.current_sheet,
            "sheet_names": self.sheet_names,
            "row_count": len(self.data),
            "column_count": len(self.data.columns),
            "column_names": list(self.data.columns),
            "data_types": self.data.dtypes.to_dict()
        }
    
    def validate_data(self, required_columns: List[str]) -> Dict[str, Any]:
        """
        验证数据是否包含所需列
        
        Args:
            required_columns: 必需的列名列表
            
        Returns:
            Dict[str, Any]: 验证结果
        """
        if self.data is None:
            return {
                "valid": False,
                "missing_columns": required_columns,
                "message": "未加载数据"
            }
            
        column_names = list(self.data.columns)
        missing_columns = [col for col in required_columns if col not in column_names]
        
        return {
            "valid": len(missing_columns) == 0,
            "missing_columns": missing_columns,
            "message": f"缺少 {len(missing_columns)} 个必需列" if missing_columns else "数据验证通过"
        }

    def _process_chengbao_term_data(self, parent_widget: Optional['QWidget'] = None) -> None:
        """
        检查并处理承保趸期数据

        检查是否存在SFYP2(不含短险续保)和首年保费列，
        如果存在则计算承保趸期数据

        Args:
            parent_widget: 父窗口组件，用于显示对话框
        """
        if self.data is None or self.data.empty:
            logger.debug("数据为空，跳过承保趸期数据处理")
            return

        # 检查必要列是否存在
        sfyp2_column = "SFYP2(不含短险续保)"
        premium_column = "首年保费"

        has_sfyp2 = sfyp2_column in self.data.columns
        has_premium = premium_column in self.data.columns

        logger.info(f"检查承保趸期数据列: SFYP2存在={has_sfyp2}, 首年保费存在={has_premium}")

        if not has_sfyp2 or not has_premium:
            if not has_sfyp2:
                logger.info(f"未找到列 '{sfyp2_column}'，跳过承保趸期数据计算")
            if not has_premium:
                logger.info(f"未找到列 '{premium_column}'，跳过承保趸期数据计算")
            return

        logger.info("检测到承保趸期数据必要列，开始计算...")
        try:
            # 调用DataFormatter计算承保趸期数据
            processed_data, pending_rows = self.memory_manager.formatter.calculate_chengbao_term_data(
                self.data,
                sfyp2_column=sfyp2_column,
                premium_column=premium_column,
                target_column="承保趸期数据",
                policy_number_column="保单号"
            )

            # 如果有需要用户输入的行，且提供了父窗口，则弹出对话框
            if pending_rows and parent_widget:
                from .gui.chengbao_term_input_dialog import ChengbaoTermInputDialog
                from PySide6.QtWidgets import QApplication

                logger.info("弹出批量输入对话框...")
                dialog = ChengbaoTermInputDialog(pending_rows, parent_widget)

                # 显示对话框
                if dialog.exec() == ChengbaoTermInputDialog.Accepted:
                    # 用户点击确定，获取输入值
                    input_values = dialog.get_input_values()
                    logger.info(f"用户提交了 {len(input_values)} 个输入值")

                    # 更新结果
                    for row_index, value in input_values.items():
                        processed_data.loc[row_index, "承保趸期数据"] = f"{value}年交SFYP"
                        logger.info(f"  第 {row_index + 1} 行更新为: {value}年交SFYP")
                else:
                    logger.info("用户取消了输入")
            elif pending_rows:
                # 没有提供父窗口，记录待输入信息
                logger.info(f"发现 {len(pending_rows)} 行需要用户输入（未提供GUI环境）")
                logger.info("需要在GUI环境中调用并提供parent_widget参数")

            # 更新数据
            self.data = processed_data

            # 记录处理结果
            logger.info(f"承保趸期数据计算完成，新增 '承保趸期数据' 列")
            logger.info(f"总行数: {len(self.data)}")

            if pending_rows:
                if parent_widget:
                    logger.info(f"已处理 {len(pending_rows)} 行用户输入")
                else:
                    logger.info(f"需要用户输入的行数：{len(pending_rows)}")
                    for row_info in pending_rows[:5]:  # 只记录前5行
                        row_idx = row_info['row_index']
                        policy_num = row_info.get('policy_number', 'N/A')
                        logger.info(f"  第 {row_idx + 1} 行 (保单号: {policy_num}) 需要用户输入")
                    if len(pending_rows) > 5:
                        logger.info(f"  ... 还有 {len(pending_rows) - 5} 行")
            else:
                logger.info("所有行的承保趸期数据已自动计算完成，无需用户输入")

        except Exception as e:
            logger.error(f"承保趸期数据计算失败: {str(e)}", exc_info=True)
            # 不抛出异常，避免影响主要功能

        # 更新日志信息
        if "承保趸期数据" in self.data.columns:
            logger.info("已添加承保趸期数据列")

