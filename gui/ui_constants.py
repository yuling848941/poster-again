"""
UI样式常量定义
统一管理界面颜色、字体、尺寸等样式参数
"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

# 配色方案
class Colors:
    """现代化配色方案"""
    # 主色调
    PRIMARY = "#2563EB"  # 品牌蓝
    PRIMARY_HOVER = "#1D4ED8"  # 悬停状态
    PRIMARY_PRESSED = "#1E40AF"  # 按下状态

    # 强调色（渐变绿）
    SUCCESS = "#10B981"
    SUCCESS_HOVER = "#059669"
    SUCCESS_START = "#10B981"
    SUCCESS_END = "#059669"

    # 次要按钮色
    SECONDARY_BLUE = "#3B82F6"  # 自动匹配
    SECONDARY_BLUE_HOVER = "#2563EB"

    # 主要操作按钮（紫色渐变）
    PRIMARY_ACTION_START = "#8B5CF6"
    PRIMARY_ACTION_END = "#7C3AED"
    PRIMARY_ACTION_HOVER_START = "#7C3AED"
    PRIMARY_ACTION_HOVER_END = "#6D28D9"

    # 中性色
    GRAY_50 = "#F9FAFB"   # 背景色
    GRAY_100 = "#F3F4F6"  # 浅灰背景
    GRAY_200 = "#E5E7EB"  # 边框色
    GRAY_400 = "#9CA3AF"  # 占位符文字
    GRAY_600 = "#6B7280"  # 正文文字
    GRAY_800 = "#1F2937"  # 深色文字
    GRAY_900 = "#111827"  # 最深文字

    # 语义化色彩
    INFO = "#3B82F6"
    WARNING = "#F59E0B"
    ERROR = "#EF4444"

    # 表格配色
    TABLE_HEADER_BG = "#F9FAFB"
    TABLE_ROW_ODD = "#FFFFFF"
    TABLE_ROW_EVEN = "#F9FAFB"
    TABLE_HOVER = "#EFF6FF"
    TABLE_SELECTED = "#DBEAFE"

    # 阴影和透明度
    SHADOW_LIGHT = "rgba(0, 0, 0, 0.1)"
    SHADOW_MEDIUM = "rgba(0, 0, 0, 0.15)"
    SHADOW_STRONG = "rgba(0, 0, 0, 0.25)"

# 字体设置
class Fonts:
    """字体配置"""
    FAMILY = "Microsoft YaHei, Arial, sans-serif"

    # 字号
    SIZE_TITLE = 16      # 标题
    SIZE_LABEL = 13      # 标签
    SIZE_BODY = 11       # 正文
    SIZE_SMALL = 12      # 小字
    SIZE_BUTTON = 14     # 按钮文字

    # 字重（使用QFont.Weight枚举）
    from PySide6.QtGui import QFont
    WEIGHT_NORMAL = QFont.Weight.Normal
    WEIGHT_MEDIUM = QFont.Weight.Medium
    WEIGHT_BOLD = QFont.Weight.Bold
    WEIGHT_HEAVY = QFont.Weight.Black

    @staticmethod
    def title_font():
        """标题字体"""
        font = QFont()
        font.setFamily(Fonts.FAMILY)
        font.setPointSize(Fonts.SIZE_TITLE)
        font.setWeight(Fonts.WEIGHT_HEAVY)
        return font

    @staticmethod
    def body_font():
        """正文字体"""
        font = QFont()
        font.setFamily(Fonts.FAMILY)
        font.setPointSize(Fonts.SIZE_BODY)
        font.setWeight(Fonts.WEIGHT_NORMAL)
        return font

    @staticmethod
    def label_font():
        """标签字体"""
        font = QFont()
        font.setFamily(Fonts.FAMILY)
        font.setPointSize(Fonts.SIZE_LABEL)
        font.setWeight(Fonts.WEIGHT_MEDIUM)
        return font

    @staticmethod
    def button_font():
        """按钮字体"""
        font = QFont()
        font.setFamily(Fonts.FAMILY)
        font.setPointSize(Fonts.SIZE_BUTTON)
        font.setWeight(Fonts.WEIGHT_MEDIUM)
        return font

# 尺寸规范
class Dimensions:
    """尺寸规范"""
    # 圆角半径
    RADIUS_SMALL = 4
    RADIUS_MEDIUM = 6
    RADIUS_LARGE = 8

    # 间距
    SPACE_XS = 4
    SPACE_SM = 8
    SPACE_MD = 12
    SPACE_LG = 16
    SPACE_XL = 24

    # 内边距
    PADDING_INPUT = 8
    PADDING_BUTTON = 4
    PADDING_PANEL = 8
    PADDING_CELL = 8

    # 边框
    BORDER_THIN = 1
    BORDER_MEDIUM = 2

    # 高度
    HEIGHT_INPUT = 25
    HEIGHT_BUTTON = 25
    HEIGHT_PROGRESS = 8
    HEIGHT_COMBOBOX = 25

    # 最小尺寸
    MIN_BUTTON_WIDTH = 100
    MIN_WINDOW_WIDTH = 1000
    MIN_WINDOW_HEIGHT = 700

# 动画配置
class Animations:
    """动画效果配置"""
    # 时长（毫秒）
    DURATION_FAST = 150       # 快速过渡
    DURATION_NORMAL = 200     # 标准过渡
    DURATION_SLOW = 300       # 慢速过渡

    # 缩放比例
    SCALE_PRESSED = 0.98

    # 缓动函数（CSS风格）
    EASE_OUT = "ease-out"
    EASE_IN_OUT = "ease-in-out"

# 按钮样式
class ButtonStyles:
    """按钮样式常量"""
    @staticmethod
    def base_style():
        """基础按钮样式"""
        return f"""
            QPushButton {{
                background-color: white;
                border: none;
                border-radius: {Dimensions.RADIUS_LARGE}px;
                padding: {Dimensions.PADDING_BUTTON}px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BUTTON}px;
                font-weight: {Fonts.WEIGHT_MEDIUM};
                min-height: {Dimensions.HEIGHT_BUTTON}px;
            }}
        """

    @staticmethod
    def table_button_style():
        """表格中的按钮样式（紧凑型）"""
        return f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.SECONDARY_BLUE},
                    stop:1 {Colors.SECONDARY_BLUE_HOVER});
                color: white;
                border: none;
                border-radius: {Dimensions.RADIUS_SMALL}px;
                padding: 1px 4px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_SMALL}px;
                font-weight: {Fonts.WEIGHT_BOLD};
                min-height: 16px;
                margin: 0px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.SECONDARY_BLUE_HOVER},
                    stop:1 {Colors.PRIMARY});
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.PRIMARY},
                    stop:1 {Colors.PRIMARY_HOVER});
            }}
        """

    @staticmethod
    def primary_action_style():
        """主要操作按钮（批量生成）"""
        return f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.PRIMARY_ACTION_START},
                    stop:1 {Colors.PRIMARY_ACTION_END});
                color: white;
                border: none;
                border-radius: {Dimensions.RADIUS_LARGE}px;
                padding: {Dimensions.PADDING_BUTTON}px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BUTTON}px;
                font-weight: {Fonts.WEIGHT_BOLD};
                min-height: {Dimensions.HEIGHT_BUTTON}px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.PRIMARY_ACTION_HOVER_START},
                    stop:1 {Colors.PRIMARY_ACTION_HOVER_END});
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.PRIMARY_ACTION_HOVER_END},
                    stop:1 {Colors.PRIMARY_ACTION_HOVER_START});
            }}
            QPushButton:disabled {{
                background: {Colors.GRAY_200};
                color: {Colors.GRAY_400};
            }}
        """

    @staticmethod
    def success_style():
        """成功按钮（配置管理）"""
        return f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.SUCCESS_START},
                    stop:1 {Colors.SUCCESS_END});
                color: white;
                border: none;
                border-radius: {Dimensions.RADIUS_LARGE}px;
                padding: {Dimensions.PADDING_BUTTON}px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BUTTON}px;
                font-weight: {Fonts.WEIGHT_BOLD};
                min-height: {Dimensions.HEIGHT_BUTTON}px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.SUCCESS},
                    stop:1 {Colors.SUCCESS_HOVER});
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.SUCCESS_HOVER},
                    stop:1 {Colors.SUCCESS});
            }}
            QPushButton:disabled {{
                background: {Colors.GRAY_200};
                color: {Colors.GRAY_400};
            }}
        """

    @staticmethod
    def secondary_style():
        """次要按钮（自动匹配）"""
        return f"""
            QPushButton {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.SECONDARY_BLUE},
                    stop:1 {Colors.SECONDARY_BLUE_HOVER});
                color: white;
                border: none;
                border-radius: {Dimensions.RADIUS_LARGE}px;
                padding: {Dimensions.PADDING_BUTTON}px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BUTTON}px;
                font-weight: {Fonts.WEIGHT_BOLD};
                min-height: {Dimensions.HEIGHT_BUTTON}px;
            }}
            QPushButton:hover {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.SECONDARY_BLUE_HOVER},
                    stop:1 {Colors.PRIMARY});
            }}
            QPushButton:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {Colors.PRIMARY},
                    stop:1 {Colors.PRIMARY_HOVER});
            }}
            QPushButton:disabled {{
                background: {Colors.GRAY_200};
                color: {Colors.GRAY_400};
            }}
        """

# 输入框样式
class InputStyles:
    """输入框样式常量"""
    @staticmethod
    def line_edit_style():
        """QLineEdit 样式"""
        return f"""
            QLineEdit {{
                background-color: white;
                border: {Dimensions.BORDER_MEDIUM}px solid {Colors.GRAY_200};
                border-radius: {Dimensions.RADIUS_MEDIUM}px;
                padding: {Dimensions.PADDING_INPUT}px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BODY}px;
                color: {Colors.GRAY_800};
                selection-background-color: {Colors.PRIMARY};
                selection-color: white;
            }}
            QLineEdit:focus {{
                border: {Dimensions.BORDER_MEDIUM}px solid {Colors.PRIMARY};
                background-color: white;
            }}
            QLineEdit:hover {{
                border-color: {Colors.PRIMARY_HOVER};
            }}
            QLineEdit::placeholder {{
                color: {Colors.GRAY_400};
            }}
        """

# 面板样式
class PanelStyles:
    """面板样式常量"""
    @staticmethod
    def group_box_style():
        """QGroupBox 样式"""
        return f"""
            QGroupBox {{
                background-color: {Colors.GRAY_50};
                border: {Dimensions.BORDER_THIN}px solid {Colors.GRAY_200};
                border-radius: {Dimensions.RADIUS_LARGE}px;
                margin-top: 1.5em;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_LABEL}px;
                font-weight: {Fonts.WEIGHT_BOLD};
                color: {Colors.GRAY_800};
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 {Dimensions.PADDING_PANEL}px;
                background-color: {Colors.GRAY_50};
                color: {Colors.GRAY_800};
            }}
        """

    @staticmethod
    def frame_style():
        """QFrame 样式"""
        return f"""
            QFrame {{
                background-color: white;
                border-radius: {Dimensions.RADIUS_MEDIUM}px;
            }}
        """

# 表格样式
class TableStyles:
    """表格样式常量"""
    @staticmethod
    def table_style():
        """QTableWidget 样式"""
        return f"""
            QTableWidget {{
                background-color: {Colors.TABLE_ROW_ODD};
                border: {Dimensions.BORDER_THIN}px solid {Colors.GRAY_200};
                border-radius: {Dimensions.RADIUS_MEDIUM}px;
                gridline-color: {Colors.GRAY_200};
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BODY}px;
                color: {Colors.GRAY_800};
                selection-background-color: {Colors.TABLE_SELECTED};
                selection-color: {Colors.GRAY_800};
            }}
            QTableWidget::item {{
                padding: 4px;
                border: none;
                border-bottom: {Dimensions.BORDER_THIN}px solid {Colors.GRAY_200};
                min-height: 20px;
            }}
            QTableWidget::item:alternate {{
                background-color: {Colors.TABLE_ROW_EVEN};
            }}
            QTableWidget::item:selected {{
                background-color: {Colors.TABLE_SELECTED};
                border: {Dimensions.BORDER_MEDIUM}px solid {Colors.PRIMARY};
            }}
            QTableWidget::item:hover {{
                background-color: {Colors.TABLE_HOVER};
            }}
            QTableWidget QTableCornerButton::section {{
                background-color: {Colors.TABLE_HEADER_BG};
                border: {Dimensions.BORDER_THIN}px solid {Colors.GRAY_200};
            }}
            QHeaderView::section {{
                background-color: {Colors.TABLE_HEADER_BG};
                border: {Dimensions.BORDER_THIN}px solid {Colors.GRAY_200};
                padding: {Dimensions.PADDING_CELL}px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_LABEL}px;
                font-weight: {Fonts.WEIGHT_BOLD};
                color: {Colors.GRAY_800};
            }}
        """

# 下拉菜单样式
class ComboBoxStyles:
    """下拉菜单样式常量"""
    @staticmethod
    def combo_box_style():
        """QComboBox 样式"""
        return f"""
            QComboBox {{
                background-color: white;
                border: {Dimensions.BORDER_MEDIUM}px solid {Colors.GRAY_200};
                border-radius: {Dimensions.RADIUS_MEDIUM}px;
                padding: {Dimensions.PADDING_INPUT}px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BODY}px;
                font-weight: {Fonts.WEIGHT_MEDIUM};
                min-height: {Dimensions.HEIGHT_INPUT}px;
                color: {Colors.GRAY_800};
            }}
            QComboBox:hover {{
                border-color: {Colors.PRIMARY};
            }}
            QComboBox:focus {{
                border-color: {Colors.PRIMARY};
            }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: {Dimensions.BORDER_MEDIUM}px;
                border-left-color: {Colors.GRAY_200};
                border-left-style: solid;
                border-top-right-radius: {Dimensions.RADIUS_MEDIUM}px;
                border-bottom-right-radius: {Dimensions.RADIUS_MEDIUM}px;
                background-color: {Colors.GRAY_100};
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid {Colors.GRAY_600};
                margin: 2px;
            }}
            QComboBox QAbstractItemView {{
                background-color: white;
                border: {Dimensions.BORDER_THIN}px solid {Colors.GRAY_200};
                border-radius: {Dimensions.RADIUS_MEDIUM}px;
                selection-background-color: {Colors.SUCCESS};
                selection-color: white;
                padding: {Dimensions.SPACE_XS}px;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BODY}px;
            }}
        """

# 进度条样式
class ProgressStyles:
    """进度条样式常量"""
    @staticmethod
    def progress_bar_style():
        """QProgressBar 样式"""
        return f"""
            QProgressBar {{
                background-color: {Colors.GRAY_200};
                border: none;
                border-radius: {Dimensions.RADIUS_SMALL}px;
                text-align: center;
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_SMALL}px;
                font-weight: {Fonts.WEIGHT_MEDIUM};
                color: {Colors.GRAY_800};
                height: {Dimensions.HEIGHT_PROGRESS}px;
            }}
            QProgressBar::chunk {{
                background-color: {Colors.PRIMARY};
                border-radius: {Dimensions.RADIUS_SMALL}px;
            }}
        """

# 日志区域样式
class LogStyles:
    """日志区域样式常量"""
    @staticmethod
    def text_edit_style():
        """QTextEdit 样式"""
        return f"""
            QTextEdit {{
                background-color: {Colors.GRAY_50};
                border: {Dimensions.BORDER_THIN}px solid {Colors.GRAY_200};
                border-radius: {Dimensions.RADIUS_MEDIUM}px;
                padding: {Dimensions.PADDING_INPUT}px;
                font-family: 'Consolas, Monaco, Courier New, monospace';
                font-size: {Fonts.SIZE_SMALL}px;
                color: {Colors.GRAY_800};
                selection-background-color: {Colors.PRIMARY};
                selection-color: white;
            }}
            QTextEdit QScrollBar:vertical {{
                background: transparent;
                width: 12px;
                border-radius: 6px;
            }}
            QTextEdit QScrollBar::handle:vertical {{
                background: {Colors.GRAY_400};
                border-radius: 6px;
                min-height: 20px;
            }}
            QTextEdit QScrollBar::handle:vertical:hover {{
                background: {Colors.GRAY_600};
            }}
            QTextEdit QScrollBar::add-line:vertical,
            QTextEdit QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
        """

# 标题样式
class TitleStyles:
    """标题样式常量"""
    @staticmethod
    def main_title_style():
        """主标题样式"""
        return f"""
            font-family: {Fonts.FAMILY};
            font-size: {Fonts.SIZE_TITLE}px;
            font-weight: {Fonts.WEIGHT_HEAVY};
            color: {Colors.PRIMARY};
        """

# 复选框样式
class CheckBoxStyles:
    """复选框样式常量"""
    @staticmethod
    def check_box_style():
        """QCheckBox 样式"""
        return f"""
            QCheckBox {{
                font-family: {Fonts.FAMILY};
                font-size: {Fonts.SIZE_BODY}px;
                color: {Colors.GRAY_800};
                spacing: 8px;
            }}
            QCheckBox::indicator {{
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: {Dimensions.BORDER_MEDIUM}px solid {Colors.GRAY_200};
                background-color: white;
            }}
            QCheckBox::indicator:hover {{
                border-color: {Colors.PRIMARY};
            }}
            QCheckBox::indicator:checked {{
                background-color: {Colors.SUCCESS};
                border-color: {Colors.SUCCESS};
            }}
        """
