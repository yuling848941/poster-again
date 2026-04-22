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
    from src.exact_matcher import ExactMatcher
    from src.ppt_generator import PPTGenerator
except ImportError:
    from gui.ppt_worker_thread import PPTWorkerThread
    from config_manager import ConfigManager
    from gui.widgets.add_text_dialog import AddTextDialog
    from gui.simple_config_dialog import SimpleConfigDialog, show_simple_config_dialog
    from gui.config_dialog import ConfigSaveLoadDialog
    from exact_matcher import ExactMatcher
    from ppt_generator import PPTGenerator


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PPT批量生成工具")
        self.setGeometry(100, 100, 1000, 700)
        
        # 初始化配置管理器
        self.config_manager = ConfigManager()

        # 初始化精确匹配器
        self.exact_matcher = ExactMatcher()

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

        # 添加"直接生成图片"复选框（在进度显示模块外面）
        self.direct_image_checkbox = QCheckBox("启用图片生成功能")
        self.direct_image_checkbox.setStyleSheet(CheckBoxStyles.check_box_style())
        self.direct_image_checkbox.stateChanged.connect(self.on_direct_image_checkbox_changed)
        layout.addWidget(self.direct_image_checkbox)

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
        layout.addWidget(self.office_refresh_btn)

        # Office 状态标签
        self.office_status_label = QLabel("")
        self.office_status_label.setStyleSheet("color: #666; font-size: 12px;")
        self.office_status_label.hide()
        layout.addWidget(self.office_status_label)

        # 添加图片格式选择下拉菜单（初始隐藏）
        image_format_layout = QHBoxLayout()
        image_format_label = QLabel("图片格式:")
        image_format_label.setFont(Fonts.label_font())
        self.image_format_combo = QComboBox()
        self.image_format_combo.addItems(["PNG", "JPG"])
        self.image_format_combo.setCurrentText("PNG")  # 默认选择PNG
        self.image_format_combo.setStyleSheet(ComboBoxStyles.combo_box_style())
        self.image_format_combo.setFont(Fonts.body_font())
        self.image_format_combo.setMinimumHeight(Dimensions.HEIGHT_INPUT)
        self.image_format_combo.currentTextChanged.connect(self.on_image_format_changed)
        self.image_format_widget = QWidget()
        self.image_format_widget.setLayout(image_format_layout)
        self.image_format_widget.hide()  # 初始隐藏
        image_format_layout.addWidget(image_format_label)
        image_format_layout.addWidget(self.image_format_combo)

        # 添加图片质量选择下拉菜单（初始隐藏）
        image_quality_layout = QHBoxLayout()
        image_quality_label = QLabel("图片质量:")
        image_quality_label.setFont(Fonts.label_font())
        self.image_quality_combo = QComboBox()
        self.image_quality_combo.addItems(["原始大小", "1.5倍", "2倍", "3倍", "4倍"])
        self.image_quality_combo.setCurrentText("原始大小")  # 默认选择原始大小
        self.image_quality_combo.setStyleSheet(ComboBoxStyles.combo_box_style())
        self.image_quality_combo.setFont(Fonts.body_font())
        self.image_quality_combo.setMinimumHeight(Dimensions.HEIGHT_INPUT)
        self.image_quality_combo.currentTextChanged.connect(self.on_image_quality_changed)
        self.image_quality_widget = QWidget()
        self.image_quality_widget.setLayout(image_quality_layout)
        self.image_quality_widget.hide()  # 初始隐藏
        image_quality_layout.addWidget(image_quality_label)
        image_quality_layout.addWidget(self.image_quality_combo)
        
        layout.addWidget(self.image_format_widget)
        layout.addWidget(self.image_quality_widget)

        # 添加弹性空间
        layout.addStretch()
        
        return panel
        
    def on_direct_image_checkbox_changed(self, state):
        """处理直接生成图片复选框状态变化"""
        # 在PySide6中，state是整数：0=未选中，2=选中
        if state == 2:  # Qt.Checked 的值是2
            # 显示图片格式和质量选择下拉菜单
            self.image_format_widget.show()
            # 显示刷新按钮和状态标签
            if hasattr(self, 'office_refresh_btn'):
                self.office_refresh_btn.show()
            if hasattr(self, 'office_status_label'):
                self.office_status_label.show()
                self._update_office_status_label()
            self.image_quality_widget.show()
            print("[LOG] 已启用直接生成图片选项")
            logger.info("已启用直接生成图片选项")
        else:  # 0=Qt.Unchecked 或其他值
            # 隐藏图片格式和质量选择下拉菜单
            self.image_format_widget.hide()
            self.image_quality_widget.hide()
            print("[LOG] 已禁用直接生成图片选项")
            logger.info("已禁用直接生成图片选项")

        # 使用防抖机制保存图片生成设置
        self._trigger_settings_save_debounced()

    def on_office_refresh_clicked(self):
        """处理 Office 检测按钮点击"""
        try:
            print("[LOG] 开始检测 Office...")
            logger.info("用户点击 Office 检测按钮")

            # 调用 Office 检测方法
            from src.core.detectors.office_suite_detector import OfficeSuiteDetector
            office_info = OfficeSuiteDetector.detect_office_suite(use_cache=False)

            if office_info:
                # 保存缓存
                self.config_manager.save_office_cache(office_info)
                print(f"[LOG] 检测到 Office: {office_info['name']}")
                logger.info(f"检测到 Office: {office_info['name']}")
            else:
                print("[LOG] 未检测到 Office")
                logger.warning("未检测到 Office")

            # 更新状态标签
            self._update_office_status_label()

        except Exception as e:
            error_msg = f"检测 Office 失败：{str(e)}"
            print(f"[LOG] {error_msg}")
            logger.error(error_msg, exc_info=True)
            self.office_status_label.setText("检测失败：" + str(e))

    def _update_office_status_label(self):
        """更新 Office 状态标签"""
        try:
            if not hasattr(self, 'office_status_label') or self.office_status_label is None:
                return

            # 尝试从缓存加载
            office_info = self.config_manager.load_office_cache()

            if office_info:
                status_text = f"已检测：{office_info['name']} ({office_info.get('version', '未知版本')})"
                self.office_status_label.setText(status_text)
                self.office_status_label.setStyleSheet("color: #2e7d32; font-size: 12px;")
            else:
                self.office_status_label.setText("未检测到 Office")
                self.office_status_label.setStyleSheet("color: #d32f2f; font-size: 12px;")

        except Exception as e:
            logger.error(f"更新 Office 状态失败：{str(e)}")
            self.office_status_label.setText("状态未知")
            self.office_status_label.setStyleSheet("color: #666; font-size: 12px;")


    def on_image_format_changed(self, format_text):
        """处理图片格式变化"""
        try:
            print(f"[LOG] 图片格式已更改为: {format_text}")
            logger.info(f"图片格式已更改为: {format_text}")

            # 使用防抖机制保存图片生成设置
            self._trigger_settings_save_debounced()

        except Exception as e:
            logger.error(f"处理图片格式变化失败: {str(e)}")
            print(f"[LOG] 处理图片格式变化时出错: {str(e)}")

    def on_image_quality_changed(self, quality_text):
        """处理图片质量变化"""
        try:
            print(f"[LOG] 图片质量已更改为: {quality_text}")
            logger.info(f"图片质量已更改为: {quality_text}")

            # 使用防抖机制保存图片生成设置
            self._trigger_settings_save_debounced()

        except Exception as e:
            logger.error(f"处理图片质量变化失败: {str(e)}")
            print(f"[LOG] 处理图片质量变化时出错: {str(e)}")

    def on_message_template_changed(self):
        """贺语模板内容变化时自动保存"""
        # 使用防抖机制保存
        self._trigger_settings_save_debounced()

    def load_message_template(self):
        """加载保存的贺语模板"""
        try:
            template = self.config_manager.get_message_template()
            if template and hasattr(self, 'message_template_edit'):
                self.message_template_edit.setPlainText(template)
                logger.info("已加载保存的贺语模板")
        except Exception as e:
            logger.error(f"加载贺语模板失败: {str(e)}")

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
                print(f"[LOG] 已选择模板文件: {file_path}")
                logger.info(f"已选择模板文件: {file_path}")
                self.worker_thread.set_template_path(file_path)

                # 自动保存路径到配置
                self.config_manager.update_last_paths(template_dir=file_path)

        except Exception as e:
            print(f"[LOG] 选择模板文件时出错: {str(e)}")
            logger.warning(f"选择模板文件失败: {str(e)}")
            
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
                print(f"[LOG] 已选择数据文件: {file_path}")
                logger.info(f"已选择数据文件: {file_path}")
                self.worker_thread.set_data_path(file_path)

                # 自动保存路径到配置
                self.config_manager.update_last_paths(data_dir=file_path)

        except Exception as e:
            print(f"[LOG] 选择数据文件时出错: {str(e)}")
            logger.warning(f"选择数据文件失败: {str(e)}")
            
    def browse_output_dir(self):
        """浏览并选择输出目录"""
        try:
            # 获取上次使用的输出目录作为起始目录
            start_dir = self.config_manager.get_start_dir_for_file_type('output')

            dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录", start_dir)
            if dir_path:
                self.output_path_edit.setText(dir_path)
                print(f"[LOG] 已选择输出目录: {dir_path}")
                logger.info(f"已选择输出目录: {dir_path}")
                self.worker_thread.set_output_path(dir_path)

                # 自动保存路径到配置
                self.config_manager.update_last_paths(output_dir=dir_path)

        except Exception as e:
            print(f"[LOG] 选择输出目录时出错: {str(e)}")
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
            print(f"[LOG] 创建输出文件夹: {folder_name}")
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
            print(f"[LOG] 将生成{image_format}格式的图片文件，质量：{image_quality}")
            logger.info(f"将生成{image_format}格式的图片文件，质量：{image_quality}")
        else:
            print("[LOG] 将生成PPT文件")
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
                print(f"[LOG] 已启用贺语生成功能")
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
        print(f"[LOG] {message}")
        logger.info(message)
        
    def populate_match_table(self, placeholders, data_columns, matching_rules=None):
        """填充匹配结果表格"""
        self.match_table.setRowCount(len(placeholders))

        for i, placeholder in enumerate(placeholders):
            # 占位符列
            self.match_table.setItem(i, 0, QTableWidgetItem(placeholder))

            # 数据列显示为只读文本项而非下拉框
            matched_column = ""

            # 优先使用传递过来的匹配规则（来自ExactMatcher的结果）
            if matching_rules and placeholder in matching_rules:
                matched_column = matching_rules[placeholder]
            else:
                # 如果没有匹配规则，使用简化的匹配逻辑
                # 移除占位符前缀
                placeholder_clean = placeholder
                for prefix in ['ph_', 'PH_', 'ph', 'PH']:
                    if placeholder.startswith(prefix):
                        placeholder_clean = placeholder[len(prefix):]
                        break

                # 优先精确匹配
                if placeholder_clean in data_columns:
                    matched_column = placeholder_clean
                else:
                    # 大小写不敏感匹配
                    for column in data_columns:
                        if column.lower() == placeholder_clean.lower():
                            matched_column = column
                            break

            self.match_table.setItem(i, 1, QTableWidgetItem(matched_column))

            # 操作按钮
            custom_btn = QPushButton("自定义")
            custom_btn.setStyleSheet(ButtonStyles.table_button_style())
            custom_btn.setFont(Fonts.button_font())
            custom_btn.clicked.connect(lambda _, r=i: self.show_custom_menu(r, custom_btn))
            self.match_table.setCellWidget(i, 2, custom_btn)

        # 重新启用按钮
        self.auto_match_btn.setEnabled(True)
        self.batch_generate_btn.setEnabled(True)
        self.config_btn.setEnabled(True)  # 【新增】启用配置管理按钮
        self.config_btn.setToolTip("")  # 【新增】清除工具提示

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
                print(f"[LOG] 已为占位符 '{placeholder}' 添加自定义文本: {display_text}"); logger.info(f"已为占位符 '{placeholder}' 添加自定义文本: {display_text}")

                # 重要：保存文本增加规则到主窗口的状态中
                if not hasattr(self, 'text_addition_rules'):
                    self.text_addition_rules = {}

                self.text_addition_rules[placeholder] = {
                    'prefix': prefix,
                    'suffix': suffix
                }

                print(f"[LOG] 已保存文本增加规则: {placeholder} -> 前缀:'{prefix}' 后缀:'{suffix}'"); logger.info(f"已保存文本增加规则: {placeholder} -> 前缀:'{prefix}' 后缀:'{suffix}'")
    
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
            print(f"[LOG] 占位符 '{placeholder_name}' 已与期趸数据列关联"); logger.info(f"占位符 '{placeholder_name}' 已与期趸数据列关联")
            
        except Exception as e:
            print(f"[LOG] 递交趸期数据失败: {str(e)}"); logger.info(f"递交趸期数据失败: {str(e)}")
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
                            print("[LOG] 为PPT生成器加载数据..."); logger.info("为PPT生成器加载数据...")
                            if not ppt_generator.load_data(excel_file):
                                print("[LOG] PPT生成器数据加载失败，将重新加载..."); logger.info("PPT生成器数据加载失败，将重新加载...")
                                # 如果load_data失败，尝试手动设置
                                ppt_generator.data = data_reader.data
                                ppt_generator.data_loaded = True

                        # 如果承保趸期数据列不存在，添加到ppt_generator的数据中
                        if "承保趸期数据" not in ppt_generator.data.columns:
                            print("[LOG] 为PPT生成器添加承保趸期数据列..."); logger.info("为PPT生成器添加承保趸期数据列...")
                            ppt_generator.data = data_reader.data.copy()
                except Exception as e:
                    print(f"[LOG] PPT生成器数据同步警告: {str(e)}"); logger.info(f"PPT生成器数据同步警告: {str(e)}")

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
                            print("[LOG] 用户取消了承保趸期数据输入，关联操作已取消"); logger.info("用户取消了承保趸期数据输入，关联操作已取消")
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
                        except:
                            pass

                if user_inputs:
                    self.chengbao_user_inputs = user_inputs
                    print(f"[LOG] 已处理承保趸期用户输入: {len(user_inputs)} 行"); logger.info(f"已处理承保趸期用户输入: {len(user_inputs)} 行")

            QMessageBox.information(self, "成功", f"已将占位符 '{placeholder_name}' 与承保趸期数据列关联")

            # 记录日志
            print(f"[LOG] 占位符 '{placeholder_name}' 已与承保趸期数据列关联"); logger.info(f"占位符 '{placeholder_name}' 已与承保趸期数据列关联")

        except Exception as e:
            print(f"[LOG] 承保趸期数据失败: {str(e)}"); logger.info(f"承保趸期数据失败: {str(e)}")
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
                print(f"[LOG] [配置管理] 承保趸期数据用户输入完成，共 {len(input_values)} 行"); logger.info(f"[配置管理] 承保趸期数据用户输入完成，共 {len(input_values)} 行")

                # 更新数据到PPT生成器
                if hasattr(self, 'worker_thread') and self.worker_thread:
                    if hasattr(self.worker_thread, 'get_ppt_generator'):
                        ppt_generator = self.worker_thread.get_ppt_generator()
                        if ppt_generator and hasattr(ppt_generator, 'data'):
                            # 将用户输入更新到承保趸期数据列
                            chengbao_column = data_reader.get_column("承保趸期数据")
                            if chengbao_column:
                                for row_index, years in input_values.items():
                                    chengbao_column[row_index] = f"{years}年交SFYP"
                                print("[LOG] 已更新承保趸期数据到PPT生成器"); logger.info("已更新承保趸期数据到PPT生成器")
            else:
                # 用户取消
                print("[LOG] 用户取消了承保趸期数据输入"); logger.info("用户取消了承保趸期数据输入")

        except Exception as e:
            print(f"[LOG] [配置管理] 承保趸期数据输入对话框失败: {str(e)}"); logger.info(f"[配置管理] 承保趸期数据输入对话框失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"承保趸期数据输入失败: {str(e)}")

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
            # 移除释放资源错误日志，简化输出
            pass

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
        """加载上次使用的路径"""
        try:
            # 从ConfigManager获取上次保存的路径
            paths = self.config_manager.load_last_paths()

            # 恢复模板文件路径
            if paths.get('template_dir') and self.config_manager.validate_path(paths['template_dir']):
                self.restore_path_to_edit(
                    paths['template_dir'],
                    self.template_path_edit,
                    self.worker_thread.set_template_path
                )

            # 恢复数据文件路径
            if paths.get('data_dir') and self.config_manager.validate_path(paths['data_dir']):
                self.restore_path_to_edit(
                    paths['data_dir'],
                    self.data_path_edit,
                    self.worker_thread.set_data_path
                )

            # 恢复输出目录路径
            if paths.get('output_dir') and self.config_manager.validate_path(paths['output_dir']):
                self.restore_path_to_edit(
                    paths['output_dir'],
                    self.output_path_edit,
                    self.worker_thread.set_output_path
                )

        except Exception as e:
            print(f"[LOG] 加载上次路径时出错: {str(e)}"); logger.info(f"加载上次路径时出错: {str(e)}")
            logger.warning(f"加载上次路径失败: {str(e)}")

    def restore_path_to_edit(self, path: str, edit_widget, setter_func):
        """
        将路径恢复到输入框并设置到工作线程

        Args:
            path: 要恢复的路径
            edit_widget: 输入框控件
            setter_func: 设置路径的方法
        """
        try:
            if path and edit_widget and setter_func:
                edit_widget.setText(path)
                setter_func(path)
                # 移除路径恢复日志，简化输出
        except Exception as e:
            print(f"[LOG] 恢复路径失败: {str(e)}"); logger.info(f"恢复路径失败: {str(e)}")
            logger.warning(f"恢复路径失败: {str(e)}")

    def load_image_generation_settings(self):
        """加载图片生成设置"""
        try:
            # 检查ConfigManager是否存在
            if not hasattr(self, 'config_manager') or self.config_manager is None:
                logger.error("ConfigManager未初始化，无法加载图片生成设置")
                return

            # 从ConfigManager获取上次保存的图片生成设置
            settings = self.config_manager.load_image_generation_settings()

            # 验证设置字典的有效性
            if not isinstance(settings, dict):
                logger.error("图片生成设置格式无效，应为字典类型")
                return

            # 恢复直接生成图片勾选状态
            if hasattr(self, 'direct_image_checkbox') and self.direct_image_checkbox is not None:
                if 'enabled' in settings and isinstance(settings['enabled'], bool):
                    self.direct_image_checkbox.setChecked(settings['enabled'])
                    logger.info(f"已恢复图片生成启用状态: {settings['enabled']}")
                else:
                    logger.warning("图片生成启用状态设置无效，使用默认值")

            # 恢复图片格式
            if hasattr(self, 'image_format_combo') and self.image_format_combo is not None:
                if 'format' in settings and settings['format'] in ['PNG', 'JPG']:
                    self.image_format_combo.setCurrentText(settings['format'])
                    logger.info(f"已恢复图片格式: {settings['format']}")
                else:
                    logger.warning("图片格式设置无效，使用默认值PNG")

            # 恢复图片质量
            if hasattr(self, 'image_quality_combo') and self.image_quality_combo is not None:
                if 'quality' in settings and isinstance(settings['quality'], (int, float)):
                    quality_text = self._quality_value_to_text(settings['quality'])
                    self.image_quality_combo.setCurrentText(quality_text)
                    logger.info(f"已恢复图片质量: {quality_text}")
                else:
                    logger.warning("图片质量设置无效，使用默认值")

        except Exception as e:
            error_msg = f"加载图片生成设置时出错: {str(e)}"
            logger.error(error_msg, exc_info=True)
            print(f"[LOG] {error_msg}")

    def save_image_generation_settings(self):
        """保存图片生成设置"""
        try:
            # 检查ConfigManager是否存在
            if not hasattr(self, 'config_manager') or self.config_manager is None:
                logger.error("ConfigManager未初始化，无法保存图片生成设置")
                return

            # 获取当前界面状态，添加安全检查
            enabled = False
            format = 'PNG'
            quality = 1.0

            if hasattr(self, 'direct_image_checkbox') and self.direct_image_checkbox is not None:
                enabled = self.direct_image_checkbox.isChecked()

            if hasattr(self, 'image_format_combo') and self.image_format_combo is not None:
                format = self.image_format_combo.currentText()
                if format not in ['PNG', 'JPG']:
                    logger.warning(f"无效的图片格式: {format}，使用默认值PNG")
                    format = 'PNG'

            if hasattr(self, 'image_quality_combo') and self.image_quality_combo is not None:
                quality_text = self.image_quality_combo.currentText()
                quality = self._quality_text_to_value(quality_text)

            # 通过ConfigManager保存
            self.config_manager.update_image_generation_settings(
                enabled=enabled,
                format=format,
                quality=quality
            )

            logger.info("图片生成设置保存成功")

        except Exception as e:
            error_msg = f"保存图片生成设置时出错: {str(e)}"
            logger.error(error_msg, exc_info=True)
            print(f"[LOG] {error_msg}")

    def _quality_value_to_text(self, quality: float) -> str:
        """将质量数值转换为对应的文本"""
        quality_map = {
            1.0: "原始大小",
            1.5: "1.5倍",
            2.0: "2倍",
            3.0: "3倍",
            4.0: "4倍"
        }
        return quality_map.get(quality, "原始大小")

    def _quality_text_to_value(self, text: str) -> float:
        """将质量文本转换为对应的数值"""
        text_map = {
            "原始大小": 1.0,
            "1.5倍": 1.5,
            "2倍": 2.0,
            "3倍": 3.0,
            "4倍": 4.0
        }
        return text_map.get(text, 1.0)

    def _save_settings_debounced(self):
        """防抖保存设置的实际执行方法（包括图片生成设置和贺语模板）"""
        try:
            # 保存图片生成设置
            self.save_image_generation_settings()

            # 保存贺语模板
            if hasattr(self, 'message_template_edit'):
                template = self.message_template_edit.toPlainText()
                self.config_manager.save_message_template(template)

        except Exception as e:
            logger.error(f"防抖保存设置失败: {str(e)}")

    def _trigger_settings_save_debounced(self):
        """触发防抖保存设置"""
        try:
            # 停止现有定时器并重启（实现防抖）
            self.settings_save_timer.stop()
            self.settings_save_timer.start(500)  # 500ms延迟
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

            # 同步匹配规则
            self.exact_matcher.clear_matching_rules()
            matched_count = 0

            if hasattr(self, 'match_table') and self.match_table is not None:
                for row in range(self.match_table.rowCount()):
                    placeholder_item = self.match_table.item(row, 0)
                    column_item = self.match_table.item(row, 1)
                    if placeholder_item and column_item:
                        placeholder = placeholder_item.text()
                        column_text = column_item.text()

                        # 跳过空值和"未匹配"
                        if not placeholder or not column_text or column_text == "未匹配":
                            continue

                        # 解析数据列名，处理自定义文本格式
                        data_column = column_text

                        # 检查是否是自定义文本格式：[原始列名] 前缀:xxx 后缀:xxx
                        if data_column.startswith("[") and "]" in data_column:
                            # 提取方括号中的原始列名
                            bracket_end = data_column.find("]")
                            data_column = data_column[1:bracket_end]

                        if data_column:
                            self.exact_matcher.set_matching_rule(placeholder, data_column)
                            matched_count += 1
                            logger.debug(f"同步匹配规则: {placeholder} -> {data_column}")

            logger.info(f"已同步匹配规则: {matched_count} 条")

            # 同步文本增加规则
            self.exact_matcher.text_addition_rules = {}
            text_rule_count = 0

            if hasattr(self, 'text_addition_rules'):
                for placeholder, rule in self.text_addition_rules.items():
                    if isinstance(rule, dict):
                        prefix = rule.get('prefix', '')
                        suffix = rule.get('suffix', '')
                        self.exact_matcher.set_text_addition_rule(placeholder, prefix, suffix)
                        text_rule_count += 1

            logger.info(f"已同步文本增加规则: {text_rule_count} 条")
            logger.info("匹配状态已成功同步到ExactMatcher")

        except Exception as e:
            logger.error(f"同步到ExactMatcher失败: {str(e)}", exc_info=True)

    def _sync_from_exact_matcher(self):
        """从ExactMatcher同步匹配状态回主界面"""
        try:
            logger.info("开始从ExactMatcher同步状态回主界面...")

            # 同步匹配规则
            matching_rules = self.exact_matcher.get_matching_rules()
            if matching_rules:
                # 只有当有新规则时才清空表格
                if hasattr(self, 'match_table') and self.match_table is not None:
                    self.match_table.setRowCount(0)
                    for placeholder, column in matching_rules.items():
                        self._add_match_row(placeholder, column)

                # 同步文本增加规则
                self.text_addition_rules = self.exact_matcher.get_all_text_addition_rules()

                # 更新日志
                print(f"[LOG] 已从配置加载 {len(matching_rules)} 条匹配规则")
                logger.info(f"已从配置加载 {len(matching_rules)} 条匹配规则")

                logger.info(f"已从ExactMatcher同步 {len(matching_rules)} 条匹配规则")
            else:
                logger.info("ExactMatcher中没有匹配规则，保持主界面状态不变")

            logger.info("匹配状态已从ExactMatcher同步回主界面")

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
        except:
            pass  # 析构函数中忽略所有异常


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())