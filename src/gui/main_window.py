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

        template_recent_btn = QPushButton("最近")
        template_recent_btn.setStyleSheet(ButtonStyles.secondary_style())
        template_recent_btn.setToolTip("从最近使用的模板中快速选择")
        template_recent_btn.clicked.connect(
            lambda: self.show_recent_menu('template', self.template_path_edit,
                                          self.worker_thread.set_template_path))
        template_layout.addWidget(template_recent_btn)

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

        data_recent_btn = QPushButton("最近")
        data_recent_btn.setStyleSheet(ButtonStyles.secondary_style())
        data_recent_btn.setToolTip("从最近使用的数据文件中快速选择")
        data_recent_btn.clicked.connect(
            lambda: self.show_recent_menu('data', self.data_path_edit,
                                          self.worker_thread.set_data_path))
        data_layout.addWidget(data_recent_btn)

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

        output_recent_btn = QPushButton("最近")
        output_recent_btn.setStyleSheet(ButtonStyles.secondary_style())
        output_recent_btn.setToolTip("从最近使用的输出目录中快速选择")
        output_recent_btn.clicked.connect(
            lambda: self.show_recent_menu('output', self.output_path_edit,
                                          self.worker_thread.set_output_path))
        output_layout.addWidget(output_recent_btn)

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

        # 导出引擎选择：默认内置渲染（快、无需 Office）；特殊模板可切 Office 高保真
        engine_layout = QHBoxLayout()
        engine_label = QLabel("导出引擎:")
        engine_label.setStyleSheet("font-weight: bold;")
        engine_label.setFont(Fonts.label_font())
        self.engine_combo = QComboBox()
        self.engine_combo.addItems(["内置渲染（快，无需 Office）", "Office 高保真（需 Office/WPS）", "自动（有 Office 用 Office）"])
        self._engine_values = ["python", "office", "auto"]   # 与下拉项一一对应
        self.engine_combo.setStyleSheet(ComboBoxStyles.combo_box_style())
        self.engine_combo.setFont(Fonts.body_font())
        self.engine_combo.setMinimumHeight(Dimensions.HEIGHT_INPUT)
        self.engine_combo.currentIndexChanged.connect(self.on_engine_changed)
        engine_layout.addWidget(engine_label)
        engine_layout.addWidget(self.engine_combo)
        engine_layout.addStretch()
        self.engine_widget = QWidget()
        self.engine_widget.setLayout(engine_layout)
        self.engine_widget.hide()
        image_layout.addWidget(self.engine_widget)

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
        self.image_quality_combo.addItems(["原始", "增强", "高质量"])
        self.image_quality_combo.setCurrentText("原始")
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
                self.engine_widget.show()
                logger.info("直接生成图片已启用")
            else:  # 0=Qt.Unchecked 或其他值
                # 隐藏图片格式和质量选择下拉菜单
                self.image_format_widget.hide()
                self.image_quality_widget.hide()
                self.engine_widget.hide()
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

    def on_engine_changed(self, index):
        """导出引擎切换：写入运行时配置，并使生成器的引擎缓存失效（下一批生效）"""
        try:
            if index < 0 or index >= len(self._engine_values):
                return
            value = self._engine_values[index]
            self.config_manager.set('processing.image_export_engine', value)
            self.config_manager.save_config()
            logger.info(f"导出引擎已切换为: {value}")
            # 使已存在的生成器引擎缓存失效（不触发懒初始化），下一批按新配置解析
            gen = getattr(self.worker_thread, '_ppt_generator', None)
            proc = getattr(gen, 'template_processor', None) if gen else None
            if proc is not None:
                proc.invalidate_engine_cache()
        except Exception as e:
            logger.error(f"切换导出引擎失败: {e}")

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

        # 启用右键自定义菜单（"增加文本"等功能的入口）
        # 必须设置 ContextMenu 策略并连接信号，否则 show_custom_menu 永远不会被触发
        self.match_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.match_table.customContextMenuRequested.connect(self._on_match_table_context_menu)

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

        # 取消生成按钮（仅批量进行中可用）
        self.cancel_btn = QPushButton("取消生成")
        self.cancel_btn.setStyleSheet(ButtonStyles.secondary_style())
        self.cancel_btn.setMinimumWidth(Dimensions.MIN_BUTTON_WIDTH)
        self.cancel_btn.setToolTip("停止批量生成（已完成当前行后停止，已生成的文件保留）")
        self.cancel_btn.setEnabled(False)
        self.cancel_btn.clicked.connect(self.cancel_generation)
        layout.addWidget(self.cancel_btn)

        return panel

    def cancel_generation(self):
        """请求取消正在进行的批量生成"""
        if hasattr(self, 'worker_thread') and self.worker_thread.isRunning():
            self.worker_thread.request_cancel()
            self.cancel_btn.setEnabled(False)
            self.progress_label.setText("正在取消...")





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

    def _collect_matching_rules_from_table(self):
        """从匹配表格收集匹配规则与文本增加规则（批量生成/失败行重试共用）"""
        matching_rules = {}
        text_addition_rules = {}

        for row in range(self.match_table.rowCount()):
            placeholder_item = self.match_table.item(row, 0)
            data_column_item = self.match_table.item(row, 1)
            if placeholder_item and data_column_item:
                placeholder = placeholder_item.text()
                data_column = data_column_item.text()

                # 检查是否是自定义文本（新的格式：[原始列名] 前缀:xxx 后缀:xxx）
                if data_column.startswith("[") and "]" in data_column:
                    # 优先读结构化数据（UserRole），避免解析显示文本出错
                    structured = data_column_item.data(Qt.UserRole)
                    if isinstance(structured, dict) and "column" in structured:
                        original_column = structured.get("column", "")
                        prefix = structured.get("prefix", "")
                        suffix = structured.get("suffix", "")
                    else:
                        # 回退：解析显示文本。提取方括号中的原始列名
                        bracket_end = data_column.find("]")
                        original_column = data_column[1:bracket_end]
                        custom_text = data_column[bracket_end + 2:]  # 移除"[列名] "前缀

                        # 按「前缀:」「后缀:」标记本身切分（而不是按空格），
                        # 这样含空格的前缀/后缀（如 "Dear Mr."）不会被截断
                        prefix = ""
                        suffix = ""
                        p_idx = custom_text.find("前缀:")
                        s_idx = custom_text.find("后缀:")
                        if p_idx != -1:
                            start = p_idx + 3
                            end = s_idx if s_idx > p_idx else len(custom_text)
                            prefix = custom_text[start:end].strip()
                        if s_idx != -1:
                            suffix = custom_text[s_idx + 3:].strip()

                    # 设置匹配规则为原始列名
                    if original_column and original_column != "自定义":
                        matching_rules[placeholder] = original_column
                        # 添加文本增加规则
                        text_addition_rules[placeholder] = {"prefix": prefix, "suffix": suffix}
                elif data_column:  # 普通匹配
                    matching_rules[placeholder] = data_column

        return matching_rules, text_addition_rules

    def batch_generate(self):
        """批量生成PPT或图片"""
        template_path = self.template_path_edit.text()
        data_path = self.data_path_edit.text()
        base_output_path = self.output_path_edit.text()

        if not template_path or not data_path or not base_output_path:
            QMessageBox.warning(self, "警告", "请先选择模板文件、数据源文件和输出目录")
            return

        # 先做校验，再创建输出文件夹（避免校验不通过时留下空的「贺报+时间戳」文件夹）
        if not self.auto_match_executed:
            QMessageBox.warning(self, "警告", "请先执行自动匹配")
            return

        if self.match_table.rowCount() == 0:
            QMessageBox.warning(self, "警告", "请先执行自动匹配")
            return

        # 收集匹配规则和文本增加规则
        matching_rules, text_addition_rules = self._collect_matching_rules_from_table()

        # 未绑定数据列的占位符会保留模板原文，先明确告知用户
        unbound = []
        for row in range(self.match_table.rowCount()):
            ph_item = self.match_table.item(row, 0)
            col_item = self.match_table.item(row, 1)
            ph = ph_item.text() if ph_item else ""
            col = col_item.text() if col_item else ""
            if ph and not col.strip():
                unbound.append(ph)
        if unbound:
            shown = "、".join(unbound[:8]) + ("…" if len(unbound) > 8 else "")
            reply = QMessageBox.question(
                self, "存在未绑定的占位符",
                f"以下 {len(unbound)} 个占位符未绑定数据列，将保留模板中的原文：\n\n"
                f"{shown}\n\n是否仍要继续生成？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return

        # 创建带时间戳的子文件夹
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        folder_name = f"贺报{timestamp}"
        output_path = os.path.join(base_output_path, folder_name)
        try:
            os.makedirs(output_path, exist_ok=True)
            logger.info(f"创建输出文件夹: {folder_name}")
        except Exception as e:
            QMessageBox.warning(self, "警告", f"创建输出文件夹失败: {str(e)}")
            return

        # 检查是否生成图片
        generate_images = self.direct_image_checkbox.isChecked()
        image_format = "PNG"  # 默认格式
        image_quality = 1.0  # 默认缩放倍数

        if generate_images:
            image_format = self.image_format_combo.currentText()
            # UI 文本转 float 缩放倍数，贯穿到导出引擎
            image_quality = self.settings_manager.quality_text_to_value(
                self.image_quality_combo.currentText()
            )
            logger.info(f"将生成{image_format}格式的图片文件，缩放倍数：{image_quality}")
        else:
            logger.info("将生成PPT文件")

        # 新的全量任务：清除失败行重试状态
        self.worker_thread.set_retry_rows(None)
        self._last_failed_rows = []

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
        self.cancel_btn.setEnabled(True)
        
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

    def _on_match_table_context_menu(self, pos):
        """
        匹配表格右键菜单信号适配器

        customContextMenuRequested 信号传递的是相对表格的 QPoint，
        需要转换成行号再交给 show_custom_menu 处理。
        如果右键没有命中具体行，则不做任何操作。
        """
        item = self.match_table.itemAt(pos)
        if item is None:
            return
        row = item.row()
        if row < 0:
            return
        # button 参数在 show_custom_menu 内部实际未被使用，传 0 占位
        self.show_custom_menu(row, 0)

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
                item = QTableWidgetItem(f"[{column_display}] {display_text}")
                # 同时存结构化数据，收集规则时优先用它（显示文本含空格也不会解析错）
                item.setData(Qt.UserRole, {
                    "column": original_column, "prefix": prefix, "suffix": suffix})
                self.match_table.setItem(row, 1, item)
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
                        # 重要：PPTGenerator 的数据存储在 data_reader.data（不是 generator.data）。
                        # 若缺少承保趸期数据列，直接用主窗口已处理好的 data_reader 替换，
                        # 确保后续批量生成使用包含承保趸期数据的完整数据。
                        generator_reader = getattr(ppt_generator, 'data_reader', None)
                        if generator_reader is None or generator_reader.data is None or \
                                "承保趸期数据" not in generator_reader.data.columns:
                            logger.info("为PPT生成器同步承保趸期数据...")
                            ppt_generator.data_reader = data_reader
                            ppt_generator.data_loaded = True

                        # 同步 ExactMatcher 的数据源，保证匹配规则基于最新数据
                        ppt_generator.exact_matcher.set_data(ppt_generator.data_reader.data)
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

    def on_generation_completed(self, summary):
        """生成完成后的处理（summary 为 worker 传入的结果摘要 dict）"""
        self.progress_label.setText("完成")

        # 兼容旧式纯成功计数参数
        if isinstance(summary, int):
            summary = {"success": summary, "failed": [], "total": summary,
                       "is_retry": False, "precheck": []}

        # 检查是否是图片生成模式
        generate_images = self.direct_image_checkbox.isChecked()
        file_type = "图片" if generate_images else "PPT"

        success = summary.get("success", 0)
        failed = summary.get("failed", [])
        degraded = summary.get("degraded", [])
        total = summary.get("total", success + len(failed))
        is_retry = summary.get("is_retry", False)
        precheck = summary.get("precheck", [])

        # 记录失败明细与本次输出目录，供"重试失败行"复用同一目录
        self._last_failed_rows = failed
        self._last_batch_output_path = self.worker_thread.output_path

        if failed:
            title = "重试完成" if is_retry else "生成完成（部分失败）"
            lines = [f"成功 {success} / {total}，失败 {len(failed)} 行：", ""]
            for f in failed[:5]:
                lines.append(f"  第 {f['row_number']} 行（{f['name']}）: {f['error'][:80]}")
            if len(failed) > 5:
                lines.append(f"  ……等共 {len(failed)} 行")
            lines.append("")
            lines.append("数据行的异常不影响其他行，可点击下方按钮只重试失败的行。")
            text = "\n".join(lines)
        else:
            title = "完成"
            prefix = "失败行重试完成，" if is_retry else ""
            text = f"{prefix}{file_type}批量生成完成！成功生成 {success} 个文件"

        # 降级提示：这些行图片导出失败，已改为输出 PPT（仍计入成功数）
        if degraded and generate_images:
            lines = [f"", f"⚠ 其中 {len(degraded)} 行图片导出失败，已降级输出为 PPT："]
            for f in degraded[:5]:
                lines.append(f"  第 {f['row_number']} 行（{f['name']}）: {f['error'][:80]}")
            if len(degraded) > 5:
                lines.append(f"  ……等共 {len(degraded)} 行")
            text += "\n".join(lines)

        # 预检提示：近似渲染效果知情告知（按问题码归并展示）
        if precheck:
            by_code = {}
            for issue in precheck:
                by_code.setdefault(issue["code"], []).append(issue)
            text += (f"\n\n⚠ 本模板有 {len(precheck)} 项近似渲染效果，"
                     f"导出图片与 PowerPoint 打开效果可能略有差异：")
            for code, issues in list(by_code.items())[:6]:
                first = issues[0]
                extra = f"（共 {len(issues)} 处）" if len(issues) > 1 else ""
                text += f"\n  · 第{first['slide']}页 {first['message']}{extra}"
            if len(by_code) > 6:
                text += f"\n  ……等 {len(by_code)} 类问题"
            text += "\n如需完全一致，可将图片导出引擎切换为 Office 高保真导出（配置项 processing.image_export_engine）。"

        retry_requested = False
        if failed:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Warning)
            box.setWindowTitle(title)
            box.setText(text)
            retry_btn = box.addButton("重试失败行", QMessageBox.AcceptRole)
            box.addButton("关闭", QMessageBox.RejectRole)
            box.exec()
            retry_requested = (box.clickedButton() is retry_btn)
        else:
            QMessageBox.information(self, title, text)

        # 重新初始化模板处理器（内部会关闭旧处理器并释放资源）
        try:
            if hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    ppt_generator._initialize_template_processor()

        except Exception as e:
            logger.warning(f"释放PPT生成器资源时出错: {e}")

        # 重新启用按钮
        self.auto_match_btn.setEnabled(True)
        self.batch_generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

        # 资源清理完成后才启动重试，避免与工作线程竞争模板处理器
        if retry_requested:
            self.retry_failed_rows()

    def retry_failed_rows(self):
        """仅重新生成上一次批量中失败的数据行（输出到同一目录）"""
        if not getattr(self, '_last_failed_rows', None):
            QMessageBox.information(self, "提示", "没有可重试的失败行")
            return
        if not getattr(self, '_last_batch_output_path', ''):
            QMessageBox.warning(self, "警告", "缺少上次输出目录信息，请重新执行批量生成")
            return

        template_path = self.template_path_edit.text()
        data_path = self.data_path_edit.text()
        if not template_path or not data_path:
            QMessageBox.warning(self, "警告", "模板或数据文件路径已变化，请重新执行批量生成")
            return

        matching_rules, text_addition_rules = self._collect_matching_rules_from_table()

        generate_images = self.direct_image_checkbox.isChecked()
        image_format = "PNG"
        image_quality = 1.0
        if generate_images:
            image_format = self.image_format_combo.currentText()
            image_quality = self.settings_manager.quality_text_to_value(
                self.image_quality_combo.currentText()
            )

        self.worker_thread.set_template_path(template_path)
        self.worker_thread.set_data_path(data_path)
        self.worker_thread.set_output_path(self._last_batch_output_path)
        self.worker_thread.set_matching_rules(matching_rules)
        self.worker_thread.set_text_addition_rules(text_addition_rules)
        self.worker_thread.set_image_generation_params(generate_images, image_format, image_quality)
        if hasattr(self, 'message_template_edit'):
            self.worker_thread.set_message_template(self.message_template_edit.toPlainText().strip())
        self.worker_thread.set_retry_rows([f["index"] for f in self._last_failed_rows])

        # 禁用按钮，防止重复点击
        self.auto_match_btn.setEnabled(False)
        self.batch_generate_btn.setEnabled(False)
        self.progress_label.setText("正在重试失败行...")

        logger.info(f"开始重试失败行：{len(self._last_failed_rows)} 行，输出目录: {self._last_batch_output_path}")
        self.worker_thread.start_batch_generate()
        self.cancel_btn.setEnabled(True)
        
    def show_error(self, message):
        """显示错误消息"""
        self.progress_label.setText("错误")
        QMessageBox.critical(self, "错误", message)

        # 即使出错也释放资源，避免文件锁定
        try:
            if hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    ppt_generator._initialize_template_processor()

        except Exception as e:
            logger.warning(f"释放PPT生成器资源时出错: {e}")

        # 重新启用按钮
        self.auto_match_btn.setEnabled(True)
        self.batch_generate_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)

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

    def show_recent_menu(self, file_type, edit_widget, setter_func):
        """弹出最近使用列表（模板/数据/输出三处"最近"按钮共用）"""
        recents = self.path_manager.get_recents(file_type)
        valid = [p for p in recents if os.path.exists(p)]
        if len(valid) != len(recents):
            self.path_manager.prune_recents(file_type, valid)

        menu = QMenu(self)
        if not valid:
            empty = menu.addAction("（暂无最近使用记录）")
            empty.setEnabled(False)
        else:
            for path in valid:
                label = path if file_type == 'output' else os.path.basename(path)
                action = menu.addAction(label)
                action.setToolTip(path)
                action.triggered.connect(
                    lambda checked, p=path: self._apply_recent_path(
                        file_type, p, edit_widget, setter_func))
        menu.exec(QCursor.pos())

    def _apply_recent_path(self, file_type, path, edit_widget, setter_func):
        """从最近列表选定路径：填入输入框、同步 worker、并置顶该记录"""
        if not os.path.exists(path):
            QMessageBox.warning(self, "路径无效", f"路径已不存在：\n{path}")
            return
        edit_widget.setText(path)
        if setter_func:
            setter_func(path)
        self.path_manager.add_recent(file_type, path)

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

        # 导出引擎下拉框同步到运行时配置（阻塞信号，避免加载时误触发保存）
        if hasattr(self, 'engine_combo') and self.engine_combo is not None:
            eng = str(self.config_manager.get('processing.image_export_engine', 'python') or 'python').lower()
            if eng in self._engine_values:
                self.engine_combo.blockSignals(True)
                self.engine_combo.setCurrentIndex(self._engine_values.index(eng))
                self.engine_combo.blockSignals(False)

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
        """窗口关闭事件处理：先请求取消并等待工作线程停下，再释放资源。"""
        try:
            # 批量进行中：请求协作式取消，等待线程在 current 行结束后停下。
            # quit() 对重写了 run() 的 QThread 无效，必须用取消标志 + wait。
            if hasattr(self, 'worker_thread') and self.worker_thread.isRunning():
                self.worker_thread.request_cancel()
                if not self.worker_thread.wait(8000):
                    logger.warning("窗口关闭：工作线程未在 8 秒内停止，进程退出将终止它")

            # 线程停下后再释放生成器资源，避免与正在进行的渲染竞争
            if hasattr(self.worker_thread, 'get_ppt_generator'):
                ppt_generator = self.worker_thread.get_ppt_generator()
                if ppt_generator:
                    ppt_generator.close()
                    logger.info("窗口关闭：已释放PPTX处理器资源")

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