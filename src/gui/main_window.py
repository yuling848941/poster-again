import sys
import os
import logging
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QLineEdit, QPushButton, QFileDialog,
    QTextEdit, QTableWidget, QTableWidgetItem, QProgressBar,
    QMessageBox, QSplitter, QFrame, QHeaderView, QComboBox, QMenu, QCheckBox,
    QDialog
)
from PySide6.QtCore import Qt, QPoint, QTimer
from PySide6.QtGui import QFont, QAction, QCursor

# 导入UI样式常量
try:
    from src.gui.ui_constants import (
        Colors, Fonts, Dimensions, ButtonStyles,
        InputStyles, PanelStyles, TableStyles,
        ComboBoxStyles, ProgressStyles, LogStyles,
        TitleStyles, CheckBoxStyles
    )
except ImportError:
    from gui.ui_constants import (
        Colors, Fonts, Dimensions, ButtonStyles,
        InputStyles, PanelStyles, TableStyles,
        ComboBoxStyles, ProgressStyles, LogStyles,
        TitleStyles, CheckBoxStyles
    )

# 设置日志
logger = logging.getLogger(__name__)

# 导入我们的功能模块
try:
    from src.gui.ppt_worker_thread import PPTWorkerThread
    from src.config import ConfigManager
    from src.gui.widgets.add_text_dialog import AddTextDialog
    from src.gui.simple_config_dialog import SimpleConfigDialog, show_simple_config_dialog
    from src.gui.config_dialog import ConfigSaveLoadDialog
    from src.gui.settings_manager import SettingsManager
    from src.gui.path_manager import PathManager
    from src.gui.match_table_manager import MatchTableManager
    from src.exact_matcher import ExactMatcher
    from src.ppt_generator import PPTGenerator
    from src.core.utils.font_checker import FontChecker
except ImportError:
    from gui.ppt_worker_thread import PPTWorkerThread
    from config import ConfigManager
    from gui.widgets.add_text_dialog import AddTextDialog
    from gui.simple_config_dialog import SimpleConfigDialog, show_simple_config_dialog
    from gui.config_dialog import ConfigSaveLoadDialog
    from gui.settings_manager import SettingsManager
    from gui.path_manager import PathManager
    from gui.match_table_manager import MatchTableManager
    from exact_matcher import ExactMatcher
    from ppt_generator import PPTGenerator
    from core.utils.font_checker import FontChecker


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT批量生成工具")
        self.setGeometry(100, 100, 1000, 700)
        
        # 初始化配置管理器
        self.config_manager = ConfigManager()

        # 初始化设置管理器 (从 main_window 拆分)
        self.settings_manager = SettingsManager(self.config_manager)

        # 初始化路径管理器 (从 main_window 拆分)
        self.path_manager = PathManager(self.config_manager)

        # 初始化精确匹配器
        self.exact_matcher = ExactMatcher()

        # 初始化字体检查器
        self.font_checker = FontChecker()

        # 注意：不再使用办公套件管理器，现在使用纯 Python PPTX 处理器

        # 初始化文本增加规则字典
        self.text_addition_rules = {}
        self.dundian_mappings = {}  # 初始化递交趸期映射关系
        self.chengbao_mappings = {}  # 初始化承保趸期映射关系

        # 初始化自动匹配标志
        self.auto_match_executed = False

        # 初始化防抖定时器（用于图片生成设置保存）
        self.settings_save_timer = QTimer()
        self.settings_save_timer.setSingleShot(True)
        self.settings_save_timer.timeout.connect(self._save_settings_debounced)

        # 初始化工作线程 - 延迟初始化office_manager
        self.worker_thread = PPTWorkerThread(office_manager=None, main_window=self)
        self.setup_worker_connections()

        # 初始化匹配表格管理器 (从 main_window 拆分)
        # Note: match_table is created in init_ui, so we'll create this manager later
        self.match_table_manager = None
        
        # 初始化UI
        self.init_ui()
        
    def setup_worker_connections(self):
        """设置工作线程信号连接"""
        self.worker_thread.progress_updated.connect(self.update_progress)
        self.worker_thread.log_message.connect(self.append_log)
        self.worker_thread.match_completed.connect(self.populate_match_table)
        self.worker_thread.generation_completed.connect(self.on_generation_completed)
        self.worker_thread.error_occurred.connect(self.show_error)

    def get_office_manager(self):
        """
        获取办公套件管理器（已废弃）

        注意：此方法保留用于向后兼容，现在始终返回 None
        因为项目已改用纯 Python PPTX 处理器，无需 Office/WPS
        """
        import warnings
        warnings.warn(
            "get_office_manager() 已废弃，项目现在使用纯 Python PPTX 处理器",
            DeprecationWarning,
            stacklevel=2
        )
        return None

    def _load_saved_configurations(self):
        """加载保存的配置数据"""
        try:
            # 暂时不需要加载承保趸期数据，每次Excel加载时会重新计算
            self.append_log("[配置管理] 承保趸期数据将在每次Excel加载时重新计算")
        except Exception as e:
            self.append_log(f"[配置管理] 加载保存的配置失败: {str(e)}")

    def init_ui(self):
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 创建主布局
        main_layout = QVBoxLayout(central_widget)
        
        # 创建标题
        title_label = QLabel("PPT批量生成工具")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(Fonts.title_font())
        title_label.setStyleSheet(TitleStyles.main_title_style())
        main_layout.addWidget(title_label)
        
        # 创建分割器
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)
        
        # 左侧面板 - 文件选择和配置
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)
        
        # 右侧面板 - 匹配结果和日志
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)
        
        # 设置分割器比例
        splitter.setSizes([400, 600])
        
        # 底部按钮面板
        button_panel = self.create_button_panel()
        main_layout.addWidget(button_panel)

        # 初始化匹配表格管理器 (match_table 在 create_right_panel 中创建)
        self.match_table_manager = MatchTableManager(self.match_table, self.exact_matcher)

        # 加载上次使用的路径
        self.load_last_paths()

        # 加载图片生成设置
        self.load_image_generation_settings()

        # 加载保存的贺语模板
        self.load_message_template()

        # 加载保存的配置数据（在所有UI组件创建完成后）
        self._load_saved_configurations()

    def create_left_panel(self):
        """创建左侧面板 - 文件选择和配置"""
        panel = QFrame()
        panel.setStyleSheet(PanelStyles.frame_style())
        layout = QVBoxLayout(panel)
        layout.setSpacing(Dimensions.SPACE_MD)

        # 模板文件选择组
        template_group = QGroupBox("模板文件")
        template_group.setStyleSheet(PanelStyles.group_box_style())
        template_layout = QHBoxLayout()
        
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setPlaceholderText("请选择PPT模板文件...")
        self.template_path_edit.setStyleSheet(InputStyles.line_edit_style())
        self.template_path_edit.setMinimumHeight(Dimensions.HEIGHT_INPUT)
        template_layout.addWidget(self.template_path_edit)

        template_browse_btn = QPushButton("浏览...")
        template_browse_btn.setStyleSheet(ButtonStyles.secondary_style())
        template_browse_btn.clicked.connect(self.browse_template_file)
        template_layout.addWidget(template_browse_btn)
        
        template_group.setLayout(template_layout)
        layout.addWidget(template_group)
        
        # 数据源文件选择组
        data_group = QGroupBox("数据源文件")
        data_group.setStyleSheet(PanelStyles.group_box_style())
        data_layout = QHBoxLayout()

        self.data_path_edit = QLineEdit()
        self.data_path_edit.setPlaceholderText("请选择Excel/CSV/JSON数据文件...")
        self.data_path_edit.setStyleSheet(InputStyles.line_edit_style())
        self.data_path_edit.setMinimumHeight(Dimensions.HEIGHT_INPUT)
        data_layout.addWidget(self.data_path_edit)

        data_browse_btn = QPushButton("浏览...")
        data_browse_btn.setStyleSheet(ButtonStyles.secondary_style())
        data_browse_btn.clicked.connect(self.browse_data_file)
        data_layout.addWidget(data_browse_btn)
        
        data_group.setLayout(data_layout)
        layout.addWidget(data_group)
        
        # 输出目录选择组
        output_group = QGroupBox("输出目录")
        output_group.setStyleSheet(PanelStyles.group_box_style())
        output_layout = QHBoxLayout()

        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("请选择输出目录...")
        self.output_path_edit.setStyleSheet(InputStyles.line_edit_style())
        self.output_path_edit.setMinimumHeight(Dimensions.HEIGHT_INPUT)
        output_layout.addWidget(self.output_path_edit)

        output_browse_btn = QPushButton("浏览...")
        output_browse_btn.setStyleSheet(ButtonStyles.secondary_style())
        output_browse_btn.clicked.connect(self.browse_output_dir)
        output_layout.addWidget(output_browse_btn)
        
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)
        
        # 进度显示
        progress_group = QGroupBox("处理进度")
        progress_group.setStyleSheet(PanelStyles.group_box_style())
        progress_layout = QVBoxLayout()

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setStyleSheet(ProgressStyles.progress_bar_style())
        progress_layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("就绪")
        progress_layout.addWidget(self.progress_label)

        progress_group.setLayout(progress_layout)
        layout.addWidget(progress_group)

        # ========== 直接生成图片模块 ==========
        image_group = QGroupBox("直接生成图片")
        image_group.setStyleSheet(PanelStyles.group_box_style())
        image_layout = QVBoxLayout(image_group)
        image_layout.setSpacing(8)

        # 复选框
        self.direct_image_checkbox = QCheckBox("启用图片生成功能")
        self.direct_image_checkbox.setStyleSheet(CheckBoxStyles.check_box_style())
        self.direct_image_checkbox.stateChanged.connect(self.on_direct_image_checkbox_changed)
        image_layout.addWidget(self.direct_image_checkbox)

        # Office 检测按钮
        self.office_refresh_btn = QPushButton("🔄 检测 Office (Microsoft/WPS)")
        self.office_refresh_btn.setToolTip("点击重新检测 Microsoft Office 或 WPS Office")
        self.office_refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 16px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #999999;
            }
        """)
        self.office_refresh_btn.setMinimumHeight(35)
        self.office_refresh_btn.clicked.connect(self.on_office_refresh_clicked)
        self.office_refresh_btn.hide()
        image_layout.addWidget(self.office_refresh_btn)

        # Office 状态标签
        self.office_status_label = QLabel("")
        self.office_status_label.setStyleSheet("color: #666; font-size: 12px;")
        self.office_status_label.hide()
        image_layout.addWidget(self.office_status_label)

        # 图片格式
        image_format_layout = QHBoxLayout()
        image_format_label = QLabel("图片格式:")
        image_format_label.setStyleSheet("font-weight: bold;")
        image_format_label.setFont(Fonts.label_font())
        self.image_format_combo = QComboBox()
        self.image_format_combo.addItems(["PNG", "JPG"])
        self.image_format_combo.setCurrentText("PNG")
        self.image_format_combo.setStyleSheet(ComboBoxStyles.combo_box_style())
        self.image_format_combo.setFont(Fonts.body_font())
        self.image_format_combo.setMinimumHeight(Dimensions.HEIGHT_INPUT)
        self.image_format_combo.currentTextChanged.connect(self.on_image_format_changed)
        image_format_layout.addWidget(image_format_label)
        image_format_layout.addWidget(self.image_format_combo)
        image_format_layout.addStretch()

        self.image_format_widget = QWidget()
        self.image_format_widget.setLayout(image_format_layout)
        self.image_format_widget.hide()
        image_layout.addWidget(self.image_format_widget)

        # 图片质量
        image_quality_layout = QHBoxLayout()
        image_quality_label = QLabel("图片质量:")
        image_quality_label.setStyleSheet("font-weight: bold;")
        image_quality_label.setFont(Fonts.label_font())
        self.image_quality_combo = QComboBox()
        self.image_quality_combo.addItems(["原始大小", "1.5 倍", "2 倍", "3 倍", "4 倍"])
        self.image_quality_combo.setCurrentText("原始大小")
        self.image_quality_combo.setStyleSheet(ComboBoxStyles.combo_box_style())
        self.image_quality_combo.setFont(Fonts.body_font())
        self.image_quality_combo.setMinimumHeight(Dimensions.HEIGHT_INPUT)
        self.image_quality_combo.currentTextChanged.connect(self.on_image_quality_changed)
        image_quality_layout.addWidget(image_quality_label)
        image_quality_layout.addWidget(self.image_quality_combo)
        image_quality_layout.addStretch()

        self.image_quality_widget = QWidget()
        self.image_quality_widget.setLayout(image_quality_layout)
        self.image_quality_widget.hide()
        image_layout.addWidget(self.image_quality_widget)

        image_group.setLayout(image_layout)
        layout.addWidget(image_group)
        # ========== 直接生成图片模块结束 ==========

        # 添加弹性空间
        layout.addStretch()
        
        return panel
        
    def on_direct_image_checkbox_changed(self, state):
        """处理直接生成图片复选框状态变化"""
        try:
            # 在 PySide6 中，state 是整数：0=未选中，2=选中
            logger.info(f"复选框状态变化：state={state}")
            if state == 2:  # Qt.Checked 的值是 2
                # 显示图片格式和质量选择下拉菜单
                self.image_format_widget.show()
                self.image_quality_widget.show()
                self.office_refresh_btn.show()  # 显示刷新按钮
                self.office_status_label.show()  # 显示状态标签
                logger.info("直接生成图片已启用，显示刷新按钮")
            else:  # 0=Qt.Unchecked 或其他值
                # 隐藏图片格式和质量选择下拉菜单
                self.image_format_widget.hide()
                self.image_quality_widget.hide()
                self.office_refresh_btn.hide()  # 隐藏刷新按钮
                self.office_status_label.hide()  # 隐藏状态标签
                logger.info("已禁用直接生成图片选项")

            # 使用防抖机制保存图片生成设置
            self._trigger_settings_save_debounced()
        except Exception as e:
            logger.error(f"处理复选框状态变化失败：{str(e)}")

    def on_image_format_changed(self, format_text):
        """处理图片格式变化"""
        try:
            logger.info(f"图片格式已更改为: {format_text}")

            # 使用防抖机制保存图片生成设置
            self._trigger_settings_save_debounced()

        except Exception as e:
            logger.error(f"处理图片格式变化失败: {str(e)}")
    def on_image_quality_changed(self, quality_text):
        """处理图片质量变化"""
        try:
            logger.info(f"图片质量已更改为: {quality_text}")

            # 使用防抖机制保存图片生成设置
            self._trigger_settings_save_debounced()

        except Exception as e:
            logger.error(f"处理图片质量变化失败: {str(e)}")
    def on_message_template_changed(self):
        """贺语模板内容变化时自动保存"""
        # 使用防抖机制保存
        self._trigger_settings_save_debounced()

    def on_office_refresh_clicked(self):
        """处理 Office 检测刷新按钮点击"""
        try:
            logger.info("用户点击刷新 Office 检测")
            # 调用检测逻辑
            result = self.check_office_availability(show_result=False)

            # 更新状态标签
            self._update_office_status_label()

            if result:
                QMessageBox.information(
                    self,
                    "Office 检测成功",
                    f"Office 检测成功！\n\n现在可以批量生成图片了。"
                )
            else:
                QMessageBox.warning(
                    self,
                    "Office 检测失败",
                    "未检测到 Microsoft Office 或 WPS Office。\n\n请确保已安装 Office 套件后重试。"
                )
        except Exception as e:
            logger.error(f"Office 检测失败：{str(e)}")
            QMessageBox.critical(
                self,
                "Office 检测错误",
                f"Office 检测过程中发生错误：\n{str(e)}"
            )

    def _update_office_status_label(self):
        """更新 Office 状态标签"""
        try:
            from src.config import ConfigManager
            config_manager = ConfigManager()
            cache_status = config_manager.get_office_cache_status()

            if cache_status.get('cached') and cache_status.get('name'):
                status_text = f"✓ 已检测到：{cache_status['name']} {cache_status.get('version', '')}"
                self.office_status_label.setText(status_text)
                self.office_status_label.setStyleSheet("color: #2e7d32; font-size: 12px; font-weight: bold;")
            else:
                self.office_status_label.setText("⚠ 未检测到 Office，请点击检测")
                self.office_status_label.setStyleSheet("color: #f57c00; font-size: 12px;")
        except Exception as e:
            logger.error(f"更新 Office 状态失败：{str(e)}")

    def check_office_availability(self, show_result: bool = False) -> bool:
        """
        检查 Office 可用性（带缓存机制）

        Args:
            show_result: 是否显示检测结果对话框

        Returns:
            bool: Office 是否可用
        """
        try:
            from src.config import ConfigManager
            from src.core.detectors.office_suite_detector import OfficeSuiteDetector

            config_manager = ConfigManager()

            # 1. 首先尝试从缓存加载
            cached_office = config_manager.load_office_cache()
            if cached_office:
                if show_result:
                    QMessageBox.information(
                        self,
                        "Office 检测成功（缓存）",
                        f"使用缓存的 Office 信息：\n"
                        f"{cached_office['name']}\n"
                        f"版本：{cached_office.get('version', '未知')}\n"
                        f"检测时间：{cached_office.get('detected_at', '未知')}"
                    )
                return True

            # 2. 缓存失效，重新检测
            if show_result:
                QMessageBox.information(self, "Office 检测", "开始重新检测 Office...")

            office_info = OfficeSuiteDetector.detect_office_suite()
            if office_info:
                # 保存缓存
                config_manager.save_office_cache(office_info)
                if show_result:
                    QMessageBox.information(
                        self,
                        "Office 检测成功",
                        f"检测到 Office 套件：\n"
                        f"{office_info['name']}\n"
                        f"版本：{office_info.get('version', '未知')}"
                    )
                return True
            else:
                if show_result:
                    QMessageBox.warning(
                        self,
                        "Office 检测失败",
                        "未检测到 Microsoft Office 或 WPS Office。\n"
                        "请确保已安装 Office 套件，或尝试重新安装。"
                    )
                return False

        except Exception as e:
            logger.error(f"Office 可用性检查失败：{str(e)}")
            if show_result:
                QMessageBox.critical(
                    self,
                    "Office 检测错误",
                    f"检测过程中发生错误：\n{str(e)}"
                )
            return False

    def load_message_template(self):
        """加载保存的贺语模板 (委托给 SettingsManager)"""
        template = self.settings_manager.load_message_template()
        if template and hasattr(self, 'message_template_edit'):
            self.message_template_edit.setPlainText(template)
            logger.info("已加载保存的贺语模板")

    def create_right_panel(self):
        """创建右侧面板 - 匹配结果和日志"""
        panel = QFrame()
        panel.setStyleSheet(PanelStyles.frame_style())
        layout = QVBoxLayout(panel)
        layout.setSpacing(Dimensions.SPACE_MD)

        # 占位符匹配结果组
        match_group = QGroupBox("占位符匹配结果")
        match_group.setStyleSheet(PanelStyles.group_box_style())
        match_layout = QVBoxLayout()

        # 创建匹配结果表格
        self.match_table = QTableWidget()
        self.match_table.setColumnCount(3)  # 增加一列用于数据列选择
        self.match_table.setHorizontalHeaderLabels(["PPT占位符", "数据字段", "操作"])
        self.match_table.setAlternatingRowColors(True)
        self.match_table.setStyleSheet(TableStyles.table_style())
        self.match_table.setFont(Fonts.body_font())

        # 初始化匹配表格管理器 (从 main_window 拆分)
        self.match_table_manager = MatchTableManager(self.match_table, self.exact_matcher)

        # 设置表格列宽
        header = self.match_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Fixed)
        header.resizeSection(2, 100)

        match_layout.addWidget(self.match_table)

        match_group.setLayout(match_layout)
        layout.addWidget(match_group)

        # 贺语模板组
        message_group = QGroupBox("贺语模板")
        message_group.setStyleSheet(PanelStyles.group_box_style())
        message_layout = QVBoxLayout()

        # 提示标签
        hint_label = QLabel("支持 {{表头}} 格式的占位符，例如：{{姓名}}、{{保费}}")
        hint_label.setFont(Fonts.body_font())
        hint_label.setStyleSheet("color: #666; margin-bottom: 5px;")
        message_layout.addWidget(hint_label)

        # 贺语模板输入框
        self.message_template_edit = QTextEdit()
        self.message_template_edit.setPlaceholderText("在此输入贺语模板...")
        self.message_template_edit.setMaximumHeight(200)
        self.message_template_edit.setStyleSheet(LogStyles.text_edit_style())
        self.message_template_edit.textChanged.connect(self.on_message_template_changed)
        message_layout.addWidget(self.message_template_edit)

        message_group.setLayout(message_layout)
        layout.addWidget(message_group)

        return panel
        
    def create_button_panel(self):
        """创建底部按钮面板"""
        panel = QFrame()
        layout = QHBoxLayout(panel)

        # 添加弹性空间
        layout.addStretch()

        # 注意：已移除办公套件选择器，现在使用纯 Python PPTX 处理器
        # 无需 Microsoft Office 或 WPS

        # 配置管理按钮
        self.config_btn = QPushButton("配置管理")
        self.config_btn.setStyleSheet(ButtonStyles.success_style())
        self.config_btn.setMinimumWidth(Dimensions.MIN_BUTTON_WIDTH)
        self.config_btn.setToolTip("请先点击\"自动匹配\"按钮")
        self.config_btn.setEnabled(False)  # 【新增】初始化时禁用
        self.config_btn.clicked.connect(self.open_config_dialog)
        layout.addWidget(self.config_btn)

        # 自动匹配按钮
        self.auto_match_btn = QPushButton("自动匹配")
        self.auto_match_btn.setStyleSheet(ButtonStyles.secondary_style())
        self.auto_match_btn.setMinimumWidth(Dimensions.MIN_BUTTON_WIDTH)
        self.auto_match_btn.clicked.connect(self.auto_match_placeholders)
        layout.addWidget(self.auto_match_btn)

        # 批量生成按钮
        self.batch_generate_btn = QPushButton("批量生成")
        self.batch_generate_btn.setStyleSheet(ButtonStyles.primary_action_style())
        self.batch_generate_btn.setMinimumWidth(Dimensions.MIN_BUTTON_WIDTH)
        self.batch_generate_btn.clicked.connect(self.batch_generate)
        layout.addWidget(self.batch_generate_btn)

        return panel





    def browse_template_file(self):
        """浏览并选择模板文件"""
        try:
            # 获取上次使用的模板目录作为起始目录
            start_dir = self.config_manager.get_start_dir_for_file_type('template')

            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择PPT模板文件", start_dir,
                "PowerPoint文件 (*.pptx);;所有文件 (*.*)"
            )
            if file_path:
                self.template_path_edit.setText(file_path)
                logger.info(f"已选择模板文件: {file_path}")
                self.worker_thread.set_template_path(file_path)

                # 自动保存路径到配置
                self.path_manager.save_path_on_select('template', file_path)

                # 检查模板使用的字体是否已安装
                QTimer.singleShot(100, lambda: self.check_template_fonts(file_path))

        except Exception as e:
            logger.warning(f"选择模板文件失败：{str(e)}")

    def check_template_fonts(self, template_path: str):
        """检查模板使用的字体是否已安装到系统"""
        try:
            result = self.font_checker.check_template_fonts(template_path)

            if not result['all_installed']:
                # 有字体缺失，显示警告
                missing_fonts = result['missing_fonts']
                msg = "⚠️ 模板使用了以下系统未安装的字体：\n\n"
                for font in missing_fonts:
                    msg += f"  • {font}\n"
                msg += "\n"
                msg += "如果启用\"图片生成功能\"，缺失字体的文本将使用微软雅黑显示。\n\n"
                msg += "建议：\n"
                msg += "  1. 安装缺失的字体文件后重新运行程序\n"
                msg += "  2. 或在 PowerPoint 中将模板字体替换为系统已安装字体"

                QMessageBox.warning(self, "字体缺失警告", msg)
            else:
                logger.info("模板使用的所有字体已安装到系统")
        except Exception as e:
            logger.warning(f"字体检查失败：{e}")

    def browse_data_file(self):
        """浏览并选择数据源文件"""
        try:
            # 获取上次使用的数据目录作为起始目录
            start_dir = self.config_manager.get_start_dir_for_file_type('data')

            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择数据源文件", start_dir,
                "Excel文件 (*.xlsx);;CSV文件 (*.csv);;JSON文件 (*.json);;所有文件 (*.*)"
            )
            if file_path:
                self.data_path_edit.setText(file_path)
                logger.info(f"已选择数据文件: {file_path}")
                self.worker_thread.set_data_path(file_path)

                # 自动保存路径到配置
                self.path_manager.save_path_on_select('data', file_path)

        except Exception as e:
            logger.warning(f"选择数据文件失败: {str(e)}")
            
    def browse_output_dir(self):
        """浏览并选择输出目录"""
        try:
            # 获取上次使用的输出目录作为起始目录
            start_dir = self.config_manager.get_start_dir_for_file_type('output')

            dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录", start_dir)
            if dir_path:
                self.output_path_edit.setText(dir_path)
                logger.info(f"已选择输出目录: {dir_path}")
                self.worker_thread.set_output_path(dir_path)

                # 自动保存路径到配置
                self.path_manager.save_path_on_select('output', dir_path)

        except Exception as e:
            logger.warning(f"选择输出目录失败: {str(e)}")
            
    def auto_match_placeholders(self):
        """自动匹配占位符"""
        template_path = self.template_path_edit.text()
        data_path = self.data_path_edit.text()
        
        if not template_path or not data_path:
            QMessageBox.warning(self, "警告", "请先选择模板文件和数据源文件")
            return
            
        # 禁用按钮，防止重复点击
        self.auto_match_btn.setEnabled(False)
        self.batch_generate_btn.setEnabled(False)
        
        # 清除旧的匹配规则，确保使用新的自动匹配算法
        self.worker_thread.clear_matching_rules()

        # 启动工作线程进行自动匹配
        self.worker_thread.start_auto_match()

        # 设置自动匹配标志为True
        self.auto_match_executed = True
    def batch_generate(self):
        """批量生成PPT或图片"""
        template_path = self.template_path_edit.text()
        data_path = self.data_path_edit.text()
        base_output_path = self.output_path_edit.text()

        if not template_path or not data_path or not base_output_path:
            QMessageBox.warning(self, "警告", "请先选择模板文件、数据源文件和输出目录")
            return

        # 创建带时间戳的子文件夹
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        folder_name = f"贺报{timestamp}"
        output_path = os.path.join(base_output_path, folder_name)

        # 创建文件夹
        try:
            os.makedirs(output_path, exist_ok=True)
            logger.info(f"创建输出文件夹: {folder_name}")
        except Exception as e:
            QMessageBox.warning(self, "警告", f"创建输出文件夹失败: {str(e)}")
            return

        # 检查是否已执行自动匹配
        if not self.auto_match_executed:
            QMessageBox.warning(self, "警告", "请先执行自动匹配")
            return

        # 检查是否有匹配结果
        if self.match_table.rowCount() == 0:
            QMessageBox.warning(self, "警告", "请先执行自动匹配")
            return

        # 检查是否生成图片
        generate_images = self.direct_image_checkbox.isChecked()
        image_format = "PNG"  # 默认格式
        image_quality = "原始大小"  # 默认质量

        if generate_images:
            image_format = self.image_format_combo.currentText()
            image_quality = self.image_quality_combo.currentText()
            logger.info(f"将生成{image_format}格式的图片文件，质量：{image_quality}")

            # 预检 Office 可用性
            if not self.check_office_availability(show_result=False):
                QMessageBox.warning(
                    self,
                    "Office 不可用",
                    "未检测到 Microsoft Office 或 WPS Office，无法生成图片。\n\n"
                    "可能原因：\n"
                    "1. 未安装 Office 套件\n"
                    "2. Office 安装在非标准路径\n"
                    "3. COM 接口异常\n\n"
                    "请检查 Office 安装后重试，或点击'🔄 检测 Office'按钮手动检测。"
                )
                return
        else:
            logger.info("将生成PPT文件")

        # 收集匹配规则和文本增加规则
        matching_rules = {}
        text_addition_rules = {}  # 新增：文本增加规则

        for row in range(self.match_table.rowCount()):
            placeholder_item = self.match_table.item(row, 0)
            data_column_item = self.match_table.item(row, 1)
            if placeholder_item and data_column_item:
                placeholder = placeholder_item.text()
                data_column = data_column_item.text()

                # 检查是否是自定义文本（新的格式：[原始列名] 前缀:xxx 后缀:xxx）
                if data_column.startswith("[") and "]" in data_column:
                    # 提取方括号中的原始列名和自定义文本信息
                    bracket_end = data_column.find("]")
                    original_column = data_column[1:bracket_end]  # 提取方括号中的内容
                    custom_text = data_column[bracket_end + 2:]  # 移除"[列名] "前缀

                    # 解析前缀和后缀
                    prefix = ""
                    suffix = ""

                    if "前缀:" in custom_text:
                        prefix_start = custom_text.find("前缀:") + 3
                        prefix_end = custom_text.find(" ", prefix_start) if " " in custom_text[prefix_start:] else len(custom_text)
                        prefix = custom_text[prefix_start:prefix_end]

                    if "后缀:" in custom_text:
                        suffix_start = custom_text.find("后缀:") + 3
                        suffix_end = custom_text.find(" ", suffix_start) if " " in custom_text[suffix_start:] else len(custom_text)
                        suffix = custom_text[suffix_start:suffix_end]

                    # 设置匹配规则为原始列名
                    if original_column and original_column != "自定义":
                        matching_rules[placeholder] = original_column
                        # 添加文本增加规则
                        text_addition_rules[placeholder] = {"prefix": prefix, "suffix": suffix}
                elif data_column:  # 普通匹配
                    matching_rules[placeholder] = data_column

        self.worker_thread.set_matching_rules(matching_rules)
        self.worker_thread.set_text_addition_rules(text_addition_rules)  # 新增：设置文本增加规则

        # 设置图片生成参数
        self.worker_thread.set_image_generation_params(generate_images, image_format, image_quality)

        # 设置贺语模板
        if hasattr(self, 'message_template_edit'):
            message_template = self.message_template_edit.toPlainText().strip()
            self.worker_thread.set_message_template(message_template)
            if message_template:
                logger.info("已启用贺语生成功能")

        # 禁用按钮，防止重复点击
        self.auto_match_btn.setEnabled(False)
        self.batch_generate_btn.setEnabled(False)

        # 启动工作线程进行批量生成
        # 直接使用原始数据路径（所有处理现在都在内存中完成）
        self.worker_thread.set_data_path(data_path)
        # 更新输出路径为新建的子文件夹
        self.worker_thread.set_output_path(output_path)
        self.worker_thread.start_batch_generate()
        
    def update_progress(self, value, message):
        """更新进度条和状态文本"""
        self.progress_bar.setValue(value)
        self.progress_label.setText(message)
        
    def append_log(self, message):
        """添加日志消息（输出到控制台）"""
        logger.info(message)
        
    def populate_match_table(self, placeholders, data_columns, matching_rules=None):
        """填充匹配结果表格 (委托给 MatchTableManager)"""
        self.match_table_manager.populate_table(placeholders, data_columns, matching_rules)

        # 重新启用按钮
        self.auto_match_btn.setEnabled(True)
        self.batch_generate_btn.setEnabled(True)
        self.config_btn.setEnabled(True)
        self.config_btn.setToolTip("")

    def edit_match_rule(self, row):
        """编辑匹配规则（已停用）"""
        # 此方法已停用，但保留以防将来需要
        pass
        
    def show_custom_menu(self, row, button):
        """显示自定义菜单"""
        menu = QMenu(self)

        # 添加"增加文本"选项
        add_text_action = QAction("增加文本", self)
        add_text_action.triggered.connect(lambda: self.add_text_to_match(row))
        menu.addAction(add_text_action)

        # 添加分隔线
        menu.addSeparator()

        # 添加"递交趸期数据"选项
        submit_data_action = QAction("递交趸期数据", self)
        submit_data_action.triggered.connect(lambda: self.submit_term_data(row))
        menu.addAction(submit_data_action)

        # 添加"承保趸期数据"选项
        submit_chengbao_action = QAction("承保趸期数据", self)
        submit_chengbao_action.triggered.connect(lambda: self.submit_chengbao_term_data(row))
        menu.addAction(submit_chengbao_action)

        # 获取鼠标全局位置并显示菜单
        cursor_pos = QCursor.pos()
        menu.popup(cursor_pos)
        
    def add_text_to_match(self, row):
        """为匹配规则添加自定义文本"""
        placeholder_item = self.match_table.item(row, 0)
        if not placeholder_item:
            return
            
        placeholder = placeholder_item.text()
        
        # 获取原始的Excel列名
        original_column = ""
        if hasattr(self.worker_thread, 'get_ppt_generator'):
            # 从PPTGenerator的ExactMatcher中获取原始匹配规则
            ppt_generator = self.worker_thread.get_ppt_generator()
            if ppt_generator:
                matching_rules = ppt_generator.exact_matcher.get_matching_rules()
            original_column = matching_rules.get(placeholder, "")
        
        # 创建并显示增加文本对话框
        dialog = AddTextDialog(placeholder, self)
        if dialog.exec() == AddTextDialog.Accepted:
            prefix, suffix = dialog.get_text_values()
            if prefix or suffix:
                # 更新表格中的数据列显示为自定义文本
                display_text = ""
                if prefix:
                    display_text += f"前缀:{prefix} "
                if suffix:
                    display_text += f"后缀:{suffix}"

                # 显示格式为: [原始列名] 前缀:xxx 后缀:xxx
                column_display = original_column if original_column else "自定义"
                self.match_table.setItem(row, 1, QTableWidgetItem(f"[{column_display}] {display_text}"))
                # 重要：保存文本增加规则到主窗口的状态中
                if not hasattr(self, 'text_addition_rules'):
                    self.text_addition_rules = {}

                self.text_addition_rules[placeholder] = {
                    'prefix': prefix,
                    'suffix': suffix
                }
    def submit_term_data(self, row):
        """递交趸期数据功能"""
        try:
            # 检查是否有占位符表格
            if not hasattr(self, 'match_table') or self.match_table.rowCount() == 0:
                QMessageBox.warning(self, "警告", "没有可处理的占位符数据")
                return
            
            # 获取当前Excel文件路径
            excel_file = self.data_path_edit.text()
            if not excel_file:
                QMessageBox.warning(self, "警告", "请先加载Excel文件")
                return
            
            # 使用内存处理获取包含期趸数据的数据
            try:
                from src.data_reader import DataReader
            except ImportError:
                from data_reader import DataReader
            data_reader = DataReader()

            # 加载Excel文件并自动处理期趸数据和千位分隔符
            if not data_reader.load_excel(excel_file, use_thousand_separator=True):
                QMessageBox.critical(self, "错误", "加载数据失败")
                return
            
            # 获取期趸数据列
            term_data_column = data_reader.get_column("期趸数据")
            
            if not term_data_column:
                QMessageBox.warning(self, "警告", "未找到期趸数据列")
                return
            
            # 获取当前选中的占位符行
            placeholder_item = self.match_table.item(row, 0)
            
            if not placeholder_item:
                return
            
            placeholder_name = placeholder_item.text()
            
            # 将期趸数据列与占位符关联
            # 更新占位符表格中的匹配值列
            match_item = QTableWidgetItem("期趸数据")
            self.match_table.setItem(row, 1, match_item)
            
            # 更新PPT生成器中的匹配规则
            if hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    ppt_generator.add_matching_rule(placeholder_name, "期趸数据")
            
            # 更新工作线程中的匹配规则
            if hasattr(self, 'worker_thread') and self.worker_thread:
                current_rules = self.worker_thread.matching_rules.copy()
                current_rules[placeholder_name] = "期趸数据"
                self.worker_thread.set_matching_rules(current_rules)

                # 使用原始Excel文件路径（数据已在内存中处理）
                self.worker_thread.set_data_path(excel_file)

            # 重要：记录递交趸期映射关系到主窗口状态
            if not hasattr(self, 'dundian_mappings'):
                self.dundian_mappings = {}

            # 记录占位符与期趸数据的映射关系
            self.dundian_mappings[placeholder_name] = "期趸数据"

            QMessageBox.information(self, "成功", f"已将占位符 '{placeholder_name}' 与期趸数据列关联")

            # 记录日志
        except Exception as e:
            QMessageBox.critical(self, "错误", f"递交趸期数据失败: {str(e)}")

    def submit_chengbao_term_data(self, row):
        """承保趸期数据功能"""
        try:
            # 检查是否有占位符表格
            if not hasattr(self, 'match_table') or self.match_table.rowCount() == 0:
                QMessageBox.warning(self, "警告", "没有可处理的占位符数据")
                return

            # 获取当前Excel文件路径
            excel_file = self.data_path_edit.text()
            if not excel_file:
                QMessageBox.warning(self, "警告", "请先加载Excel文件")
                return

            # 使用内存处理获取包含承保趸期数据的数据
            try:
                from src.data_reader import DataReader
            except ImportError:
                from data_reader import DataReader
            data_reader = DataReader()

            # 加载Excel文件并自动处理承保趸期数据（传递self作为parent_widget以支持对话框）
            if not data_reader.load_excel(excel_file, use_thousand_separator=True, parent_widget=self):
                QMessageBox.critical(self, "错误", "加载数据失败")
                return

            # 获取承保趸期数据列
            chengbao_data_column = data_reader.get_column("承保趸期数据")

            if not chengbao_data_column:
                QMessageBox.warning(self, "警告", "未找到承保趸期数据列\n请确保Excel文件包含'SFYP2(不含短险续保)'和'首年保费'列")
                return

            # 重要：确保PPT生成器已加载数据且包含承保趸期数据列
            if hasattr(self, 'worker_thread') and self.worker_thread and hasattr(self.worker_thread, 'get_ppt_generator'):
                try:
                    ppt_generator = self.worker_thread.get_ppt_generator()
                    if ppt_generator:
                        # 如果ppt_generator还没有数据，加载数据
                        if not hasattr(ppt_generator, 'data') or ppt_generator.data is None:
                            logger.info("为PPT生成器加载数据...")
                            if not ppt_generator.load_data(excel_file):
                                logger.info("PPT生成器数据加载失败，将重新加载...")
                                # 如果load_data失败，尝试手动设置
                                ppt_generator.data = data_reader.data
                                ppt_generator.data_loaded = True

                        # 如果承保趸期数据列不存在，添加到ppt_generator的数据中
                        if "承保趸期数据" not in ppt_generator.data.columns:
                            logger.info("为PPT生成器添加承保趸期数据列...")
                            ppt_generator.data = data_reader.data.copy()
                except Exception as e:
                    logger.debug(f"PPT 生成器数据同步警告：{str(e)}")
            # 检查是否有需要用户输入的行（R=1的情况）
            has_empty_values = any(value == "" for value in chengbao_data_column)

            if has_empty_values:
                reply = QMessageBox.question(
                    self,
                    "需要用户输入",
                    "检测到部分承保趸期数据需要手动输入（比例R=1的情况）。\n"
                    "将显示批量输入对话框，请填写所有需要输入的行。\n\n"
                    "是否继续？",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )

                if reply == QMessageBox.Yes:
                    # 重新加载，这次会弹出对话框
                    data_reader.load_excel(excel_file, use_thousand_separator=True, parent_widget=self)
                    chengbao_data_column = data_reader.get_column("承保趸期数据")

                    # 重新检查是否还有空值
                    # 注意：这里只检查之前检测到的需要用户输入的行，而不是整个列
                    # 如果对话框被用户取消，会返回Cancelled，这时候chengbao_data_column可能还是包含空值
                    remaining_empty = any(value == "" for value in chengbao_data_column)

                    if remaining_empty:
                        # 检查是否是因为对话框被取消
                        # 如果对话框被取消，DataReader会记录日志
                        # 这里我们简单提示用户
                        result = QMessageBox.question(
                            self,
                            "输入被取消",
                            "批量输入对话框被取消或未完成输入。\n\n"
                            "是否仍然关联占位符（未填写的行将显示为空）？",
                            QMessageBox.Yes | QMessageBox.No,
                            QMessageBox.No
                        )

                        if result == QMessageBox.No:
                            logger.info("用户取消了承保趸期数据输入，关联操作已取消")
                            return
                        # 如果用户选择Yes，继续执行关联操作，即使有空值

            # 获取当前选中的占位符行
            placeholder_item = self.match_table.item(row, 0)

            if not placeholder_item:
                return

            placeholder_name = placeholder_item.text()

            # 将承保趸期数据列与占位符关联
            # 更新占位符表格中的匹配值列
            match_item = QTableWidgetItem("承保趸期数据")
            self.match_table.setItem(row, 1, match_item)

            # 更新PPT生成器中的匹配规则
            if hasattr(self, 'worker_thread') and self.worker_thread and hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    ppt_generator.add_matching_rule(placeholder_name, "承保趸期数据")

            # 更新工作线程中的匹配规则
            if hasattr(self, 'worker_thread') and self.worker_thread:
                current_rules = self.worker_thread.matching_rules.copy()
                current_rules[placeholder_name] = "承保趸期数据"
                self.worker_thread.set_matching_rules(current_rules)

                # 使用原始Excel文件路径（数据已在内存中处理）
                self.worker_thread.set_data_path(excel_file)

            # 重要：记录承保趸期映射关系到主窗口状态
            if not hasattr(self, 'chengbao_mappings'):
                self.chengbao_mappings = {}

            # 记录占位符与承保趸期数据的映射关系
            self.chengbao_mappings[placeholder_name] = "承保趸期数据"

            # 重要：保存用户输入的承保趸期数据到主窗口状态
            if not hasattr(self, 'chengbao_user_inputs'):
                self.chengbao_user_inputs = {}

            # 如果DataReader中有承保趸期数据列，获取用户输入的行
            if "承保趸期数据" in data_reader.data.columns:
                chengbao_column = data_reader.get_column("承保趸期数据")
                user_inputs = {}

                # 遍历承保趸期数据列，找出用户输入的行（R=1的行，格式为"X年交SFYP"）
                for i, value in enumerate(chengbao_column):
                    if value and "年交SFYP" in value:
                        # 这是用户输入的数据，提取数值
                        try:
                            years = int(value.replace("年交SFYP", ""))
                            user_inputs[i] = years
                        except (ValueError, TypeError):
                            pass

                if user_inputs:
                    self.chengbao_user_inputs = user_inputs
            QMessageBox.information(self, "成功", f"已将占位符 '{placeholder_name}' 与承保趸期数据列关联")

            # 记录日志
        except Exception as e:
            QMessageBox.critical(self, "错误", f"承保趸期数据失败: {str(e)}")

    def show_chengbao_input_dialog(self, data_reader, user_input_rows):
        """
        显示承保趸期数据输入对话框
        用于配置加载后自动触发用户输入

        Args:
            data_reader: DataReader实例，包含已加载的Excel数据
            user_input_rows: 需要用户输入的行索引列表
        """
        try:
            # 准备pending_rows数据
            pending_rows = []
            policy_number_column = data_reader.get_column("保单号")

            for row_index in user_input_rows:
                # 获取保单号
                policy_number = "N/A"
                if policy_number_column and row_index < len(policy_number_column):
                    policy_number = policy_number_column[row_index]

                pending_rows.append({
                    'row_index': row_index,
                    'policy_number': policy_number
                })

            # 导入对话框类
            try:
                from src.gui.chengbao_term_input_dialog import ChengbaoTermInputDialog
            except ImportError:
                from gui.chengbao_term_input_dialog import ChengbaoTermInputDialog

            # 创建并显示对话框
            dialog = ChengbaoTermInputDialog(pending_rows, self)
            result = dialog.exec()

            if result == QDialog.DialogCode.Accepted:
                # 用户点击确定，获取输入值
                input_values = dialog.get_input_values()

                # 保存用户输入到主窗口状态
                if not hasattr(self, 'chengbao_user_inputs'):
                    self.chengbao_user_inputs = {}

                self.chengbao_user_inputs.update(input_values)
                logger.info(f"[配置管理] 承保趸期数据用户输入完成，共 {len(input_values)} 行")

                # 更新数据到 PPT 生成器
                if hasattr(self, 'worker_thread') and self.worker_thread:
                    if hasattr(self.worker_thread, 'get_ppt_generator'):
                        ppt_generator = self.worker_thread.get_ppt_generator()
                        if ppt_generator and hasattr(ppt_generator, 'data'):
                            # 将用户输入更新到承保趸期数据列
                            chengbao_column = data_reader.get_column("承保趸期数据")
                            if chengbao_column:
                                for row_index, years in input_values.items():
                                    chengbao_column[row_index] = f"{years}年交 SFYP"
                                logger.info("已更新承保趸期数据到 PPT 生成器")
            else:
                # 用户取消
                logger.info("用户取消了承保趸期数据输入")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"承保趸期数据输入失败：{str(e)}")

    def on_generation_completed(self, count):
        """生成完成后的处理"""
        self.progress_label.setText("完成")

        # 检查是否是图片生成模式
        generate_images = self.direct_image_checkbox.isChecked()
        file_type = "图片" if generate_images else "PPT"

        QMessageBox.information(self, "完成", f"{file_type}批量生成完成！成功生成 {count} 个文件")

        # ✅ 新增：完全释放所有COM资源，包括PowerPoint应用程序
        try:
            if hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    # 关闭演示文稿
                    if hasattr(ppt_generator.template_processor, 'close_presentation'):
                        ppt_generator.template_processor.close_presentation()

                    # 关闭PowerPoint应用程序
                    ppt_generator.close_template_processor()

                    # ✅ 新增：重新初始化模板处理器，确保下次使用全新的实例
                    ppt_generator._initialize_template_processor()

        except Exception as e:
            logger.warning(f"释放PPT生成器资源时出错: {e}")

        # 重新启用按钮
        self.auto_match_btn.setEnabled(True)
        self.batch_generate_btn.setEnabled(True)
        
    def show_error(self, message):
        """显示错误消息"""
        self.progress_label.setText("错误")
        QMessageBox.critical(self, "错误", message)

        # ✅ 新增：即使出错也要完全释放所有资源，避免文件锁定
        try:
            if hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    # 关闭演示文稿
                    if hasattr(ppt_generator.template_processor, 'close_presentation'):
                        ppt_generator.template_processor.close_presentation()

                    # 关闭PowerPoint应用程序
                    ppt_generator.close_template_processor()

                    # ✅ 新增：重新初始化模板处理器，确保下次使用时正常
                    ppt_generator._initialize_template_processor()

        except Exception as e:
            # 移除错误处理日志，简化输出
            pass

        # 重新启用按钮
        self.auto_match_btn.setEnabled(True)
        self.batch_generate_btn.setEnabled(True)

    def load_last_paths(self):
        """加载上次使用的路径 (委托给 PathManager)"""
        self.path_manager.restore_all_paths(
            self.template_path_edit,
            self.data_path_edit,
            self.output_path_edit,
            self.worker_thread.set_template_path,
            self.worker_thread.set_data_path,
            self.worker_thread.set_output_path,
        )

    def restore_path_to_edit(self, path: str, edit_widget, setter_func):
        """将路径恢复到输入框并设置到工作线程 (委托给 PathManager)"""
        self.path_manager.restore_path(path, edit_widget, setter_func)

    def load_image_generation_settings(self):
        """加载图片生成设置 (委托给 SettingsManager)"""
        settings = self.settings_manager.load_image_settings()
        logger.info(f"加载图片生成设置：enabled={settings.get('enabled', False)}")

        if hasattr(self, 'direct_image_checkbox') and self.direct_image_checkbox is not None:
            self.direct_image_checkbox.setChecked(settings['enabled'])
            logger.info(f"设置复选框状态：{settings['enabled']}")
            # 手动调用状态变化处理函数以同步按钮可见性
            state = 2 if settings['enabled'] else 0
            logger.info(f"手动调用 on_direct_image_checkbox_changed，state={state}")
            self.on_direct_image_checkbox_changed(state)
        else:
            logger.warning("direct_image_checkbox 不存在或未初始化")

        if hasattr(self, 'image_format_combo') and self.image_format_combo is not None:
            if settings['format'] in ['PNG', 'JPG']:
                self.image_format_combo.setCurrentText(settings['format'])

        if hasattr(self, 'image_quality_combo') and self.image_quality_combo is not None:
            quality_text = self.settings_manager.quality_value_to_text(settings['quality'])
            self.image_quality_combo.setCurrentText(quality_text)

        # 更新 Office 状态标签
        self._update_office_status_label()

    def save_image_generation_settings(self):
        """保存图片生成设置 (委托给 SettingsManager)"""
        enabled = False
        fmt = 'PNG'
        quality = 1.0

        if hasattr(self, 'direct_image_checkbox') and self.direct_image_checkbox is not None:
            enabled = self.direct_image_checkbox.isChecked()

        if hasattr(self, 'image_format_combo') and self.image_format_combo is not None:
            fmt = self.image_format_combo.currentText()
            if fmt not in ['PNG', 'JPG']:
                fmt = 'PNG'

        if hasattr(self, 'image_quality_combo') and self.image_quality_combo is not None:
            quality = self.settings_manager.quality_text_to_value(
                self.image_quality_combo.currentText()
            )

        self.settings_manager.save_image_settings(enabled, fmt, quality)

    def _save_settings_debounced(self):
        """防抖保存设置的实际执行方法 (委托给 SettingsManager)"""
        try:
            self.save_image_generation_settings()
            if hasattr(self, 'message_template_edit'):
                self.settings_manager.save_message_template(
                    self.message_template_edit.toPlainText()
                )
        except Exception as e:
            logger.error(f"防抖保存设置失败: {str(e)}")

    def _trigger_settings_save_debounced(self):
        """触发防抖保存设置"""
        try:
            self.settings_save_timer.stop()
            self.settings_save_timer.start(500)
        except Exception as e:
            logger.error(f"触发防抖保存失败: {str(e)}")

    def open_config_dialog(self):
        """打开配置管理对话框 - 使用简化的配置管理"""
        try:
            logger.info("用户点击配置管理按钮")

            # 使用简化的配置管理，直接基于GUI状态
            show_simple_config_dialog(self)

        except Exception as e:
            error_msg = f"打开配置管理对话框失败: {str(e)}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "错误", error_msg)

    def _sync_exact_matcher(self):
        """将主界面的匹配状态同步到ExactMatcher"""
        try:
            logger.info("开始同步主界面状态到ExactMatcher...")

            # 获取当前数据
            if hasattr(self, 'data_frame') and self.data_frame is not None:
                self.exact_matcher.set_data(self.data_frame)
                logger.info(f"已同步数据，列数: {len(self.data_frame.columns)}")

            # 获取当前模板占位符
            if hasattr(self, 'template_placeholders') and self.template_placeholders:
                self.exact_matcher.set_template_placeholders(self.template_placeholders)
                logger.info(f"已同步模板占位符: {len(self.template_placeholders)} 个")

            # 使用 MatchTableManager 同步匹配规则
            if self.match_table_manager:
                self.match_table_manager.sync_to_exact_matcher(self.text_addition_rules)
            else:
                logger.warning("match_table_manager 未初始化，跳过同步")

            logger.info("匹配状态已成功同步到ExactMatcher")

        except Exception as e:
            logger.error(f"同步到ExactMatcher失败: {str(e)}", exc_info=True)

    def _sync_from_exact_matcher(self):
        """从ExactMatcher同步匹配状态回主界面"""
        try:
            logger.info("开始从ExactMatcher同步状态回主界面...")

            # 使用 MatchTableManager 同步匹配规则
            if self.match_table_manager:
                text_rules = self.match_table_manager.sync_from_exact_matcher()
                self.text_addition_rules = text_rules
                logger.info("匹配状态已从ExactMatcher同步回主界面")
            else:
                logger.warning("match_table_manager 未初始化")
                # 回退到原有逻辑
                matching_rules = self.exact_matcher.get_matching_rules()
                if matching_rules:
                    if hasattr(self, 'match_table') and self.match_table is not None:
                        self.match_table.setRowCount(0)
                        if self.match_table_manager:
                            for placeholder, column in matching_rules.items():
                                self.match_table_manager._add_match_row(placeholder, column)
                        else:
                            for placeholder, column in matching_rules.items():
                                self._add_match_row(placeholder, column)
                    self.text_addition_rules = self.exact_matcher.get_all_text_addition_rules()
                    logger.info(f"已从ExactMatcher同步 {len(matching_rules)} 条匹配规则")
                else:
                    logger.info("ExactMatcher中没有匹配规则，保持主界面状态不变")

        except Exception as e:
            logger.error(f"从ExactMatcher同步失败: {str(e)}", exc_info=True)

    def _add_match_row(self, placeholder: str, column: str):
        """添加匹配行到表格"""
        try:
            if hasattr(self, 'match_table') and self.match_table is not None:
                row = self.match_table.rowCount()
                self.match_table.insertRow(row)

                # 添加占位符
                placeholder_item = QTableWidgetItem(placeholder)
                self.match_table.setItem(row, 0, placeholder_item)

                # 添加数据字段
                column_item = QTableWidgetItem(column)
                self.match_table.setItem(row, 1, column_item)

                # 添加文本增加预览
                text_rule = self.exact_matcher.get_text_addition_rule(placeholder)
                prefix = text_rule.get('prefix', '')
                suffix = text_rule.get('suffix', '')
                text_preview = f"'{prefix}'内容'{suffix}'" if prefix or suffix else ""
                text_item = QTableWidgetItem(text_preview)
                self.match_table.setItem(row, 2, text_item)

        except Exception as e:
            logger.error(f"添加匹配行失败: {str(e)}")

    def closeEvent(self, event):
        """窗口关闭事件处理"""
        try:
            # 确保PPT生成器资源被释放
            if hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    ppt_generator.close()
                    logger.info("窗口关闭：已释放PPTX处理器资源")

            # 等待工作线程完成
            if hasattr(self, 'worker_thread') and self.worker_thread.isRunning():
                self.worker_thread.quit()
                self.worker_thread.wait(3000)  # 最多等待3秒

        except Exception as e:
            logger.error(f"窗口关闭时资源释放失败: {str(e)}")

        event.accept()

    def __del__(self):
        """析构函数，确保资源释放"""
        try:
            if hasattr(self, 'worker_thread') and hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    ppt_generator.close()
        except Exception:
            # 析构函数中忽略所有异常，避免干扰 Python 垃圾回收
            pass

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())