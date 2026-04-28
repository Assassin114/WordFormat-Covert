"""
模板配置Tab
配置模板各项格式
"""

from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QScrollArea,
    QLabel, QPushButton, QGroupBox, QFormLayout,
    QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox,
    QCheckBox, QFileDialog, QMessageBox, QListWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QTabWidget,
    QColorDialog, QButtonGroup, QRadioButton, QTextEdit,
    QGridLayout, QFrame
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFontDatabase

from ..models.template_config import (
    create_default_template, TemplateConfig,
    FontConfig, ParagraphConfig, HeadingConfig, CaptionConfig,
    HeaderFooterConfig, CoverConfig, TableStyleConfig, TableFontConfig,
    PageNumberFormat, PageNumberPosition, PageNumberAlignment,
    PrintMode
)


# 中文标准字号
CHINESE_FONT_SIZES = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
}
FONT_SIZE_NAMES = list(CHINESE_FONT_SIZES.keys())

# 行距类型
LINE_SPACING_TYPES = ["固定值", "最小值", "单倍行距", "多倍行距"]
LINE_SPACING_CODES = ["fixed", "atLeast", "single", "multiple"]

# 页码格式
PAGE_NUMBER_FORMATS = ["阿拉伯数字", "小写罗马数字", "大写罗马数字", "小写字母", "大写字母"]
PAGE_NUMBER_FORMAT_CODES = ["arabic", "roman_low", "roman_up", "letter_low", "letter_up"]


class TemplateConfigWidget(QWidget):
    """
    模板配置界面
    """
    
    def __init__(self):
        super().__init__()
        self.template = create_default_template()
        self.template_path = None
        self._modified = False
        self._system_fonts = []
        
        self._init_ui()
        self._load_system_fonts()
    
    def _init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)
        
        # 创建标签页
        tabs = QTabWidget()
        
        # Tab 1: 基础配置
        tab_basic = self._create_basic_tab()
        tabs.addTab(tab_basic, "基础配置")
        
        # Tab 2: 标题格式
        tab_heading = self._create_heading_tab()
        tabs.addTab(tab_heading, "标题格式")
        
        # Tab 3: 正文格式
        tab_body = self._create_body_tab()
        tabs.addTab(tab_body, "正文格式")
        
        # Tab 4: 表格格式
        tab_table = self._create_table_tab()
        tabs.addTab(tab_table, "表格格式")
        
        # Tab 5: 题注格式
        tab_caption = self._create_caption_tab()
        tabs.addTab(tab_caption, "题注格式")
        
        # Tab 6: 页眉页脚
        tab_hf = self._create_header_footer_tab()
        tabs.addTab(tab_hf, "页眉页脚")
        
        # Tab 7: 封面字段
        tab_cover = self._create_cover_tab()
        tabs.addTab(tab_cover, "封面字段")
        
        layout.addWidget(tabs)
        
        # 底部按钮
        btn_layout = QHBoxLayout()
        self.btn_new = QPushButton("新建")
        self.btn_new.clicked.connect(self.new_template)
        btn_layout.addWidget(self.btn_new)
        
        self.btn_load = QPushButton("加载模板")
        self.btn_load.clicked.connect(self.load_template)
        btn_layout.addWidget(self.btn_load)
        
        self.btn_save = QPushButton("保存模板")
        self.btn_save.clicked.connect(self.save_template)
        btn_layout.addWidget(self.btn_save)
        
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
    
    def _create_basic_tab(self) -> QWidget:
        """创建基础配置标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # ===== 打印模式 =====
        print_group = QGroupBox("打印模式")
        print_layout = QHBoxLayout()
        
        self.print_mode_group = QButtonGroup()
        self.radio_single = QRadioButton("单面打印")
        self.radio_duplex = QRadioButton("双面打印")
        self.print_mode_group.addButton(self.radio_single, 0)
        self.print_mode_group.addButton(self.radio_duplex, 1)
        self.radio_single.setChecked(True)
        
        print_layout.addWidget(self.radio_single)
        print_layout.addWidget(self.radio_duplex)
        print_layout.addStretch()
        
        print_group.setLayout(print_layout)
        scroll_layout.addWidget(print_group)
        
        # ===== 页眉页脚全局设置 =====
        hf_global_group = QGroupBox("页眉页脚全局设置")
        hf_global_layout = QVBoxLayout()
        
        self.cb_first_diff = QCheckBox("首页不同")
        self.cb_odd_even = QCheckBox("奇偶页不同")
        
        hf_global_layout.addWidget(self.cb_first_diff)
        hf_global_layout.addWidget(self.cb_odd_even)
        
        hf_global_group.setLayout(hf_global_layout)
        scroll_layout.addWidget(hf_global_group)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        
        return widget
    
    def _create_heading_tab(self) -> QWidget:
        """创建标题格式标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        self.heading_editors = []
        
        for level in range(1, 6):
            editor = HeadingLevelEditor(level)
            scroll_layout.addWidget(editor)
            self.heading_editors.append(editor)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        
        return widget
    
    def _create_body_tab(self) -> QWidget:
        """创建正文格式标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # ===== 正文字体 =====
        font_group = QGroupBox("正文字体")
        font_layout = QGridLayout()
        
        # 中文字体
        font_layout.addWidget(QLabel("中文字体:"), 0, 0)
        self.body_font_cn = QComboBox()
        self.body_font_cn.setEditable(True)
        font_layout.addWidget(self.body_font_cn, 0, 1)
        
        # 英文字体
        font_layout.addWidget(QLabel("英文字体:"), 0, 2)
        self.body_font_en = QComboBox()
        self.body_font_en.setEditable(True)
        font_layout.addWidget(self.body_font_en, 0, 3)
        
        # 字号
        font_layout.addWidget(QLabel("字号:"), 1, 0)
        self.body_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.body_size.addItem(name)
        self.body_size.setCurrentText("五号")
        font_layout.addWidget(self.body_size, 1, 1)
        
        # 颜色
        font_layout.addWidget(QLabel("颜色:"), 1, 2)
        self.body_color = QPushButton("选择颜色")
        self.body_color.setStyleSheet("background-color: #000000; color: white;")
        self.body_color.clicked.connect(lambda: self._select_color(self.body_color))
        font_layout.addWidget(self.body_color, 1, 3)
        
        # 加粗/斜体
        self.body_bold = QCheckBox("加粗")
        self.body_italic = QCheckBox("斜体")
        font_layout.addWidget(self.body_bold, 2, 1)
        font_layout.addWidget(self.body_italic, 2, 2)
        
        font_group.setLayout(font_layout)
        layout.addWidget(font_group)
        
        # ===== 正文段落 =====
        para_group = QGroupBox("段落格式")
        para_layout = QGridLayout()
        
        # 行距类型
        para_layout.addWidget(QLabel("行距类型:"), 0, 0)
        self.body_spacing_type = QComboBox()
        for t in LINE_SPACING_TYPES:
            self.body_spacing_type.addItem(t)
        self.body_spacing_type.currentIndexChanged.connect(self._on_spacing_type_changed)
        para_layout.addWidget(self.body_spacing_type, 0, 1)
        
        # 行距值
        para_layout.addWidget(QLabel("行距值:"), 0, 2)
        self.body_spacing_value = QDoubleSpinBox()
        self.body_spacing_value.setRange(1, 100)
        self.body_spacing_value.setValue(1.5)
        self.body_spacing_value.setDecimals(1)
        para_layout.addWidget(self.body_spacing_value, 0, 3)
        
        # 段前间距
        para_layout.addWidget(QLabel("段前间距(pt):"), 1, 0)
        self.body_space_before = QDoubleSpinBox()
        self.body_space_before.setRange(0, 100)
        self.body_space_before.setValue(0)
        para_layout.addWidget(self.body_space_before, 1, 1)
        
        # 段后间距
        para_layout.addWidget(QLabel("段后间距(pt):"), 1, 2)
        self.body_space_after = QDoubleSpinBox()
        self.body_space_after.setRange(0, 100)
        self.body_space_after.setValue(0)
        para_layout.addWidget(self.body_space_after, 1, 3)
        
        # 首行缩进
        para_layout.addWidget(QLabel("首行缩进(字符):"), 2, 0)
        self.body_indent = QDoubleSpinBox()
        self.body_indent.setRange(0, 10)
        self.body_indent.setValue(2)
        para_layout.addWidget(self.body_indent, 2, 1)
        
        # 对齐方式
        para_layout.addWidget(QLabel("对齐方式:"), 2, 2)
        self.body_align = QComboBox()
        self.body_align.addItems(["左对齐", "居中", "右对齐", "两端对齐"])
        para_layout.addWidget(self.body_align, 2, 3)
        
        para_group.setLayout(para_layout)
        layout.addWidget(para_group)
        
        layout.addStretch()
        
        return widget
    
    def _create_table_tab(self) -> QWidget:
        """创建表格格式标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # ===== 基础表格式 =====
        regular_group = QGroupBox("基础表（常规表格）")
        regular_layout = QVBoxLayout()
        
        # 表头字体
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("表头字体:"))
        self.regular_header_font = QComboBox()
        self.regular_header_font.setEditable(True)
        h_layout.addWidget(self.regular_header_font)
        h_layout.addWidget(QLabel("字号:"))
        self.regular_header_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.regular_header_size.addItem(name)
        h_layout.addWidget(self.regular_header_size)
        self.regular_header_bold = QCheckBox("加粗")
        h_layout.addWidget(self.regular_header_bold)
        h_layout.addStretch()
        regular_layout.addLayout(h_layout)
        
        # 表头背景色
        h_layout2 = QHBoxLayout()
        h_layout2.addWidget(QLabel("表头背景色:"))
        self.regular_header_bg = QPushButton("选择颜色")
        self.regular_header_bg.setStyleSheet("background-color: #E8E8E8;")
        self.regular_header_bg.clicked.connect(lambda: self._select_color(self.regular_header_bg))
        h_layout2.addWidget(self.regular_header_bg)
        h_layout2.addWidget(QLabel("对齐:"))
        self.regular_header_align = QComboBox()
        self.regular_header_align.addItems(["左对齐", "居中", "右对齐"])
        h_layout2.addWidget(self.regular_header_align)
        h_layout2.addStretch()
        regular_layout.addLayout(h_layout2)
        
        # 正文字体
        h_layout3 = QHBoxLayout()
        h_layout3.addWidget(QLabel("正文字体:"))
        self.regular_body_font = QComboBox()
        self.regular_body_font.setEditable(True)
        h_layout3.addWidget(self.regular_body_font)
        h_layout3.addWidget(QLabel("字号:"))
        self.regular_body_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.regular_body_size.addItem(name)
        h_layout3.addWidget(self.regular_body_size)
        h_layout3.addStretch()
        regular_layout.addLayout(h_layout3)
        
        regular_group.setLayout(regular_layout)
        layout.addWidget(regular_group)
        
        # ===== 记录表格式 =====
        record_group = QGroupBox("记录表（测试/试验记录表）")
        record_layout = QVBoxLayout()
        
        # 表头字体
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("表头字体:"))
        self.record_header_font = QComboBox()
        self.record_header_font.setEditable(True)
        h_layout.addWidget(self.record_header_font)
        h_layout.addWidget(QLabel("字号:"))
        self.record_header_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.record_header_size.addItem(name)
        h_layout.addWidget(self.record_header_size)
        self.record_header_bold = QCheckBox("加粗")
        h_layout.addWidget(self.record_header_bold)
        h_layout.addStretch()
        record_layout.addLayout(h_layout)
        
        # 正文字体
        h_layout3 = QHBoxLayout()
        h_layout3.addWidget(QLabel("正文字体:"))
        self.record_body_font = QComboBox()
        self.record_body_font.setEditable(True)
        h_layout3.addWidget(self.record_body_font)
        h_layout3.addWidget(QLabel("字号:"))
        self.record_body_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.record_body_size.addItem(name)
        h_layout3.addWidget(self.record_body_size)
        h_layout3.addStretch()
        record_layout.addLayout(h_layout3)
        
        record_group.setLayout(record_layout)
        layout.addWidget(record_group)
        
        layout.addStretch()
        
        return widget
    
    def _create_caption_tab(self) -> QWidget:
        """创建题注格式标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 图序
        fig_group = QGroupBox("图序")
        fig_layout = QGridLayout()
        
        fig_layout.addWidget(QLabel("前缀:"), 0, 0)
        self.fig_prefix = QLineEdit("图")
        fig_layout.addWidget(self.fig_prefix, 0, 1)
        
        fig_layout.addWidget(QLabel("位置:"), 0, 2)
        self.fig_position = QComboBox()
        self.fig_position.addItems(["下方", "上方"])
        fig_layout.addWidget(self.fig_position, 0, 3)
        
        fig_layout.addWidget(QLabel("字体:"), 1, 0)
        self.fig_font = QComboBox()
        self.fig_font.setEditable(True)
        fig_layout.addWidget(self.fig_font, 1, 1)
        
        fig_layout.addWidget(QLabel("字号:"), 1, 2)
        self.fig_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.fig_size.addItem(name)
        fig_layout.addWidget(self.fig_size, 1, 3)
        
        fig_group.setLayout(fig_layout)
        layout.addWidget(fig_group)
        
        # 表序
        table_group = QGroupBox("表序")
        table_layout = QGridLayout()
        
        table_layout.addWidget(QLabel("前缀:"), 0, 0)
        self.table_prefix = QLineEdit("表")
        table_layout.addWidget(self.table_prefix, 0, 1)
        
        table_layout.addWidget(QLabel("位置:"), 0, 2)
        self.table_position = QComboBox()
        self.table_position.addItems(["上方", "下方"])
        table_layout.addWidget(self.table_position, 0, 3)
        
        table_layout.addWidget(QLabel("字体:"), 1, 0)
        self.table_font = QComboBox()
        self.table_font.setEditable(True)
        table_layout.addWidget(self.table_font, 1, 1)
        
        table_layout.addWidget(QLabel("字号:"), 1, 2)
        self.table_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.table_size.addItem(name)
        table_layout.addWidget(self.table_size, 1, 3)
        
        table_group.setLayout(table_layout)
        layout.addWidget(table_group)
        
        layout.addStretch()
        
        return widget
    
    def _create_header_footer_tab(self) -> QWidget:
        """创建页眉页脚标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # ===== 第1类：封面页 =====
        cover_hf_group = QGroupBox("第1类：封面页 / 签署页 / 修改记录页")
        cover_hf_layout = QVBoxLayout()
        
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("页眉文本:"))
        self.cover_hf_text = QLineEdit("文档标题")
        h_layout.addWidget(self.cover_hf_text)
        cover_hf_layout.addLayout(h_layout)
        
        cb_layout = QHBoxLayout()
        self.cover_show_header = QCheckBox("显示页眉")
        self.cover_show_footer = QCheckBox("显示页脚")
        self.cover_show_footer.setChecked(False)
        cb_layout.addWidget(self.cover_show_header)
        cb_layout.addWidget(self.cover_show_footer)
        cb_layout.addStretch()
        cover_hf_layout.addLayout(cb_layout)
        
        cover_hf_group.setLayout(cover_hf_layout)
        scroll_layout.addWidget(cover_hf_group)
        
        # ===== 第2类：目录页 =====
        toc_hf_group = QGroupBox("第2类：目录页")
        toc_hf_layout = QVBoxLayout()
        
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("页眉文本:"))
        self.toc_hf_text = QLineEdit("目录")
        h_layout.addWidget(self.toc_hf_text)
        toc_hf_layout.addLayout(h_layout)
        
        cb_layout = QHBoxLayout()
        self.toc_show_header = QCheckBox("显示页眉")
        self.toc_show_header.setChecked(True)
        self.toc_show_footer = QCheckBox("显示页脚")
        self.toc_show_footer.setChecked(True)
        cb_layout.addWidget(self.toc_show_header)
        cb_layout.addWidget(self.toc_show_footer)
        cb_layout.addStretch()
        toc_hf_layout.addLayout(cb_layout)
        
        # 页码设置
        pg_layout = QHBoxLayout()
        pg_layout.addWidget(QLabel("页码格式:"))
        self.toc_page_format = QComboBox()
        for fmt in PAGE_NUMBER_FORMATS:
            self.toc_page_format.addItem(fmt)
        pg_layout.addWidget(self.toc_page_format)
        
        pg_layout.addWidget(QLabel("页码位置:"))
        self.toc_page_position = QComboBox()
        self.toc_page_position.addItems(["页眉", "页脚"])
        pg_layout.addWidget(self.toc_page_position)
        
        pg_layout.addWidget(QLabel("对齐:"))
        self.toc_page_align = QComboBox()
        self.toc_page_align.addItems(["左", "中", "右"])
        pg_layout.addWidget(self.toc_page_align)
        pg_layout.addStretch()
        toc_hf_layout.addLayout(pg_layout)
        
        # 页码字体
        font_layout = QHBoxLayout()
        font_layout.addWidget(QLabel("页码字体:"))
        self.toc_page_font = QComboBox()
        self.toc_page_font.setEditable(True)
        font_layout.addWidget(self.toc_page_font)
        font_layout.addWidget(QLabel("字号:"))
        self.toc_page_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.toc_page_size.addItem(name)
        font_layout.addWidget(self.toc_page_size)
        self.toc_page_bold = QCheckBox("加粗")
        font_layout.addWidget(self.toc_page_bold)
        font_layout.addStretch()
        toc_hf_layout.addLayout(font_layout)
        
        toc_hf_group.setLayout(toc_hf_layout)
        scroll_layout.addWidget(toc_hf_group)
        
        # ===== 第3类：正文部分 =====
        body_hf_group = QGroupBox("第3类：正文部分")
        body_hf_layout = QVBoxLayout()
        
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("页眉文本:"))
        self.body_hf_text = QLineEdit("正文")
        h_layout.addWidget(self.body_hf_text)
        body_hf_layout.addLayout(h_layout)
        
        cb_layout = QHBoxLayout()
        self.body_show_header = QCheckBox("显示页眉")
        self.body_show_header.setChecked(True)
        self.body_show_footer = QCheckBox("显示页脚")
        self.body_show_footer.setChecked(True)
        cb_layout.addWidget(self.body_show_header)
        cb_layout.addWidget(self.body_show_footer)
        cb_layout.addStretch()
        body_hf_layout.addLayout(cb_layout)
        
        # 页码设置
        pg_layout = QHBoxLayout()
        pg_layout.addWidget(QLabel("页码格式:"))
        self.body_page_format = QComboBox()
        for fmt in PAGE_NUMBER_FORMATS:
            self.body_page_format.addItem(fmt)
        pg_layout.addWidget(self.body_page_format)
        
        pg_layout.addWidget(QLabel("页码位置:"))
        self.body_page_position = QComboBox()
        self.body_page_position.addItems(["页眉", "页脚"])
        pg_layout.addWidget(self.body_page_position)
        
        pg_layout.addWidget(QLabel("对齐:"))
        self.body_page_align = QComboBox()
        self.body_page_align.addItems(["左", "中", "右"])
        pg_layout.addWidget(self.body_page_align)
        pg_layout.addStretch()
        body_hf_layout.addLayout(pg_layout)
        
        # 页码字体
        font_layout = QHBoxLayout()
        font_layout.addWidget(QLabel("页码字体:"))
        self.body_page_font = QComboBox()
        self.body_page_font.setEditable(True)
        self.body_page_font.addItem("Times New Roman")
        font_layout.addWidget(self.body_page_font)
        font_layout.addWidget(QLabel("字号:"))
        self.body_page_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.body_page_size.addItem(name)
        font_layout.addWidget(self.body_page_size)
        self.body_page_bold = QCheckBox("加粗")
        font_layout.addWidget(self.body_page_bold)
        font_layout.addStretch()
        body_hf_layout.addLayout(font_layout)
        
        body_hf_group.setLayout(body_hf_layout)
        scroll_layout.addWidget(body_hf_group)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        
        return widget
    
    def _create_cover_tab(self) -> QWidget:
        """创建封面字段标签页"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        label = QLabel("封面字段编辑（用于替换模板中的占位符）")
        label.setStyleSheet("color: gray; padding: 5px;")
        layout.addWidget(label)
        
        self.cover_fields_table = QTableWidget()
        self.cover_fields_table.setColumnCount(2)
        self.cover_fields_table.setHorizontalHeaderLabels(["字段名", "值"])
        self.cover_fields_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.cover_fields_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._init_cover_fields_table()
        layout.addWidget(self.cover_fields_table)
        
        # 按钮行
        btn_row = QHBoxLayout()
        btn_add_field = QPushButton("添加字段")
        btn_add_field.clicked.connect(self._add_cover_field)
        btn_row.addWidget(btn_add_field)
        
        btn_remove_field = QPushButton("删除选中")
        btn_remove_field.clicked.connect(self._remove_cover_field)
        btn_row.addWidget(btn_remove_field)
        
        btn_row.addStretch()
        layout.addLayout(btn_row)
        
        # 预定义字段
        preset_group = QGroupBox("插入预定义字段")
        preset_layout = QHBoxLayout()
        
        presets = [
            ("合同编号", "contract_no"),
            ("密级", "classification"),
            ("年份", "year"),
            ("单位", "organization"),
            ("日期", "date"),
            ("拟制", "drafter"),
            ("校对", "reviewer"),
            ("标审", "approver"),
            ("审核", "checker"),
            ("批准", "final_approver"),
        ]
        
        for label_text, key in presets:
            btn = QPushButton(label_text)
            btn.clicked.connect(lambda checked, k=key: self._insert_preset_field(k))
            preset_layout.addWidget(btn)
        
        preset_layout.addStretch()
        preset_group.setLayout(preset_layout)
        layout.addWidget(preset_group)
        
        return widget
    
    def _init_cover_fields_table(self):
        """初始化封面字段表格"""
        standard_fields = [
            ("contract_no", "合同编号"),
            ("classification", "密级"),
            ("year", "年份"),
            ("organization", "单位"),
            ("date", "日期"),
            ("drafter", "拟制人"),
            ("drafter_time", "拟制时间"),
            ("reviewer", "校对人"),
            ("reviewer_time", "校对时间"),
            ("approver", "标审人"),
            ("approver_time", "标审时间"),
            ("checker", "审核人"),
            ("checker_time", "审核时间"),
            ("final_approver", "批准人"),
            ("final_approver_time", "批准时间"),
        ]
        
        self.cover_fields_table.setRowCount(len(standard_fields))
        for i, (key, desc) in enumerate(standard_fields):
            key_item = QTableWidgetItem(key)
            key_item.setFlags(key_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.cover_fields_table.setItem(i, 0, key_item)
            value_item = QTableWidgetItem("")
            value_item.setText(self.template.cover.fields.get(key, ""))
            self.cover_fields_table.setItem(i, 1, value_item)
    
    def _load_system_fonts(self):
        """从系统加载字体列表"""
        self._system_fonts = QFontDatabase.families()
        
        # 更新所有字体下拉框
        font_combos = [
            self.body_font_cn, self.body_font_en,
            self.regular_header_font, self.regular_body_font,
            self.record_header_font, self.record_body_font,
            self.fig_font, self.table_font,
            self.toc_page_font, self.body_page_font,
        ]
        
        for combo in font_combos:
            combo.addItems(self._system_fonts)
            # 默认设置
            if combo == self.body_font_cn:
                combo.setCurrentText("宋体")
            elif combo == self.body_font_en:
                combo.setCurrentText("Times New Roman")
            elif combo == self.regular_header_font:
                combo.setCurrentText("宋体")
            elif combo == self.regular_body_font:
                combo.setCurrentText("宋体")
            elif combo == self.record_header_font:
                combo.setCurrentText("宋体")
            elif combo == self.record_body_font:
                combo.setCurrentText("宋体")
            elif combo == self.fig_font:
                combo.setCurrentText("宋体")
            elif combo == self.table_font:
                combo.setCurrentText("宋体")
            elif combo == self.toc_page_font:
                combo.setCurrentText("Times New Roman")
            elif combo == self.body_page_font:
                combo.setCurrentText("Times New Roman")
    
    def _select_color(self, button: QPushButton):
        """选择颜色"""
        color = QColorDialog.getColor()
        if color.isValid():
            button.setStyleSheet(f"background-color: {color.name()}; color: white;")
    
    def _on_spacing_type_changed(self, index: int):
        """行距类型变更"""
        if index == 0:  # 固定值
            self.body_spacing_value.setRange(10, 100)
            self.body_spacing_value.setValue(20)
        elif index == 1:  # 最小值
            self.body_spacing_value.setRange(10, 100)
            self.body_spacing_value.setValue(15)
        elif index == 2:  # 单倍行距
            self.body_spacing_value.setValue(1.0)
            self.body_spacing_value.setEnabled(False)
        else:  # 多倍行距
            self.body_spacing_value.setRange(0.5, 3.0)
            self.body_spacing_value.setValue(1.5)
            self.body_spacing_value.setEnabled(True)
    
    def _add_cover_field(self):
        """添加封面字段"""
        row = self.cover_fields_table.rowCount()
        self.cover_fields_table.insertRow(row)
        self.cover_fields_table.setItem(row, 0, QTableWidgetItem(""))
        self.cover_fields_table.setItem(row, 1, QTableWidgetItem(""))
    
    def _remove_cover_field(self):
        """删除选中的封面字段"""
        current_row = self.cover_fields_table.currentRow()
        if current_row >= 0:
            self.cover_fields_table.removeRow(current_row)
    
    def _insert_preset_field(self, key: str):
        """插入预定义字段"""
        for row in range(self.cover_fields_table.rowCount()):
            if self.cover_fields_table.item(row, 0).text() == key:
                self.cover_fields_table.selectRow(row)
                return
        
        row = self.cover_fields_table.rowCount()
        self.cover_fields_table.insertRow(row)
        key_item = QTableWidgetItem(key)
        key_item.setFlags(key_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.cover_fields_table.setItem(row, 0, key_item)
        self.cover_fields_table.setItem(row, 1, QTableWidgetItem(""))
        self.cover_fields_table.selectRow(row)
    
    def new_template(self):
        """新建模板"""
        self.template = create_default_template()
        self.template_path = None
        self._modified = False
    
    def load_template(self):
        """加载模板"""
        file, _ = QFileDialog.getOpenFileName(
            self, "加载模板", "", "JSON文件 (*.json);;所有文件 (*.*)"
        )
        
        if file:
            try:
                from ..templates.template_io import load_template
                self.template = load_template(file)
                self.template_path = file
                self._sync_ui_from_template()
                self._modified = False
                QMessageBox.information(self, "成功", f"已加载模板:\n{file}")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"加载模板失败:\n{str(e)}")
    
    def save_template(self):
        """保存模板"""
        file, _ = QFileDialog.getSaveFileName(
            self, "保存模板", "template.json", "JSON文件 (*.json);;所有文件 (*.*)"
        )
        
        if file:
            try:
                from ..templates.template_io import save_template
                self._sync_template_from_ui()
                save_template(self.template, file)
                self.template_path = file
                self._modified = False
                QMessageBox.information(self, "成功", f"已保存模板:\n{file}")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"保存模板失败:\n{str(e)}")
    
    def get_template(self) -> TemplateConfig:
        """获取当前模板"""
        self._sync_template_from_ui()
        return self.template
    
    def _sync_template_from_ui(self):
        """同步UI到模板对象"""
        # 打印模式
        if self.print_mode_group.checkedId() == 0:
            self.template.print_mode = PrintMode.SINGLE
        else:
            self.template.print_mode = PrintMode.DUPLEX
        
        # 页眉页脚全局
        self.template.header_footer.different_first_page = self.cb_first_diff.isChecked()
        self.template.header_footer.different_odd_even = self.cb_odd_even.isChecked()
        
        # 同步标题
        for i, editor in enumerate(self.heading_editors):
            if i < len(self.template.headings):
                h = self.template.headings[i]
                h.font.name = editor.font_name
                h.font.size = editor.font_size
                h.font.bold = editor.font_bold
                h.paragraph.line_spacing = editor.line_spacing
                h.paragraph.space_before = editor.space_before
                h.paragraph.space_after = editor.space_after
        
        # 同步正文
        self.template.body.font.name = self.body_font_cn.currentText()
        self.template.body.font.size = CHINESE_FONT_SIZES.get(self.body_size.currentText(), 10.5)
        self.template.body.font.bold = self.body_bold.isChecked()
        self.template.body.font.italic = self.body_italic.isChecked()
        
        # 行距
        idx = self.body_spacing_type.currentIndex()
        self.template.body.line_spacing_type = LINE_SPACING_CODES[idx]
        self.template.body.line_spacing = self.body_spacing_value.value()
        
        # 封面字段
        self.template.cover.fields = {}
        for row in range(self.cover_fields_table.rowCount()):
            key_item = self.cover_fields_table.item(row, 0)
            value_item = self.cover_fields_table.item(row, 1)
            if key_item and key_item.text().strip():
                key = key_item.text().strip()
                value = value_item.text() if value_item else ""
                self.template.cover.fields[key] = value
    
    def _sync_ui_from_template(self):
        """同步模板对象到UI"""
        # 打印模式
        if self.template.print_mode == PrintMode.DUPLEX:
            self.radio_duplex.setChecked(True)
        else:
            self.radio_single.setChecked(True)
        
        # 页眉页脚全局
        self.cb_first_diff.setChecked(self.template.header_footer.different_first_page)
        self.cb_odd_even.setChecked(self.template.header_footer.different_odd_even)
        
        # 封面字段
        for row in range(self.cover_fields_table.rowCount()):
            key = self.cover_fields_table.item(row, 0).text()
            if key in self.template.cover.fields:
                self.cover_fields_table.item(row, 1).setText(self.template.cover.fields[key])


class HeadingLevelEditor(QWidget):
    """单级标题格式编辑器"""
    
    value_changed = pyqtSignal()
    
    def __init__(self, level: int):
        super().__init__()
        self.level = level
        self._font_name = "黑体"
        self._font_size = 16
        self._font_bold = True
        self._line_spacing = 1.5
        self._space_before = 12
        self._space_after = 6
        
        self._init_ui()
    
    def _init_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 5, 0, 5)
        
        level_label = QLabel(f"{self.level}级:")
        level_label.setMinimumWidth(30)
        layout.addWidget(level_label)
        
        # 字体
        self.font_combo = QComboBox()
        self.font_combo.addItems(["黑体", "楷体", "宋体", "微软雅黑", "仿宋"])
        self.font_combo.setCurrentText(self._font_name)
        self.font_combo.currentTextChanged.connect(self._on_changed)
        layout.addWidget(QLabel("字体:"))
        layout.addWidget(self.font_combo)
        
        # 字号
        self.size_combo = QComboBox()
        for name in FONT_SIZE_NAMES:
            self.size_combo.addItem(name)
        self.size_combo.setCurrentText(self._get_chinese_size_name(self._font_size))
        self.size_combo.currentTextChanged.connect(self._on_size_changed)
        layout.addWidget(QLabel("字号:"))
        layout.addWidget(self.size_combo)
        
        # 加粗
        self.bold_check = QCheckBox("加粗")
        self.bold_check.setChecked(self._font_bold)
        self.bold_check.stateChanged.connect(self._on_changed)
        layout.addWidget(self.bold_check)
        
        # 行距
        self.spacing_spin = QDoubleSpinBox()
        self.spacing_spin.setRange(1.0, 3.0)
        self.spacing_spin.setValue(self._line_spacing)
        self.spacing_spin.setSingleStep(0.1)
        self.spacing_spin.valueChanged.connect(self._on_changed)
        layout.addWidget(QLabel("行距:"))
        layout.addWidget(self.spacing_spin)
        
        layout.addStretch()
    
    def _on_changed(self):
        self._font_name = self.font_combo.currentText()
        self._font_size = self._get_font_size_from_chinese(self.size_combo.currentText())
        self._font_bold = self.bold_check.isChecked()
        self._line_spacing = self.spacing_spin.value()
        self.value_changed.emit()
    
    def _on_size_changed(self):
        self._font_size = self._get_font_size_from_chinese(self.size_combo.currentText())
        self.value_changed.emit()
    
    @staticmethod
    def _get_font_size_from_chinese(name: str) -> float:
        return CHINESE_FONT_SIZES.get(name, 12)
    
    @staticmethod
    def _get_chinese_size_name(size: float) -> str:
        for name, s in CHINESE_FONT_SIZES.items():
            if abs(s - size) < 0.5:
                return name
        return "小四"
    
    @property
    def font_name(self): return self._font_name
    @font_name.setter
    def font_name(self, value):
        self._font_name = value
        self.font_combo.setCurrentText(value)
    
    @property
    def font_size(self): return self._font_size
    @font_size.setter
    def font_size(self, value):
        self._font_size = value
        self.size_combo.setCurrentText(self._get_chinese_size_name(value))
    
    @property
    def font_bold(self): return self._font_bold
    @font_bold.setter
    def font_bold(self, value):
        self._font_bold = value
        self.bold_check.setChecked(value)
    
    @property
    def line_spacing(self): return self._line_spacing
    @line_spacing.setter
    def line_spacing(self, value):
        self._line_spacing = value
        self.spacing_spin.setValue(value)
    
    @property
    def space_before(self): return self._space_before
    @space_before.setter
    def space_before(self, value): self._space_before = value
    
    @property
    def space_after(self): return self._space_after
    @space_after.setter
    def space_after(self, value): self._space_after = value
