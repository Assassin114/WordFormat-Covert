"""
Template Config Widget - Left/Right Layout
"""
from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QScrollArea,
    QLabel, QPushButton, QGroupBox, QFormLayout,
    QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox,
    QCheckBox, QFileDialog, QMessageBox,
    QListWidget, QStackedWidget,
    QColorDialog, QButtonGroup, QRadioButton,
    QGridLayout, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from ..models.template_config import (
    create_default_template, TemplateConfig
)

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
LINE_SPACING_TYPES = ["固定值", "最小值", "单倍行距", "多倍行距"]


class TemplateConfigWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.template = create_default_template()
        self.template_path = None
        self._init_ui()
        self._connect_signals()
    
    def _init_ui(self):
        layout = QHBoxLayout(self)
        left_panel = self._create_left_panel()
        layout.addWidget(left_panel, 1)
        self.right_stack = QStackedWidget()
        self._create_config_pages()
        layout.addWidget(self.right_stack, 3)
        btn_layout = QHBoxLayout()
        for txt in ["新建", "加载模板", "保存模板"]:
            btn = QPushButton(txt)
            btn_layout.addWidget(btn)
        btn_layout.addStretch()
        main_layout = QVBoxLayout()
        main_layout.addLayout(layout, 1)
        main_layout.addLayout(btn_layout)
        self.setLayout(main_layout)
    
    def _create_left_panel(self) -> QWidget:
        widget = QFrame()
        widget.setFrameStyle(QFrame.Shape.StyledPanel)
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel("<b>配置项目</b>"))
        self.config_list = QListWidget()
        items = [
            ("heading1", "标题1"), ("heading2", "标题2"), ("heading3", "标题3"),
            ("heading4", "标题4"), ("heading5", "标题5"), ("body", "正文"),
            ("table_regular", "表格-基础表"), ("table_record", "表格-记录表"),
            ("caption_figure", "题注-图序"), ("caption_table", "题注-表序"),
            ("header_all", "页眉页脚"), ("cover_fields", "封面字段"),
        ]
        for item_id, item_name in items:
            item = QListWidgetItem(item_name)
            item.setData(Qt.ItemDataRole.UserRole, item_id)
            self.config_list.addItem(item)
        self.config_list.setCurrentRow(0)
        layout.addWidget(self.config_list)
        return widget
    
    def _create_config_pages(self):
        self.heading_page = self._create_heading_page()
        self.right_stack.addWidget(self.heading_page)
        self.body_page = self._create_body_page()
        self.right_stack.addWidget(self.body_page)
        self.table_page = self._create_table_page()
        self.right_stack.addWidget(self.table_page)
        self.caption_page = self._create_caption_page()
        self.right_stack.addWidget(self.caption_page)
        self.header_page = self._create_header_page()
        self.right_stack.addWidget(self.header_page)
        self.cover_page = self._create_cover_page()
        self.right_stack.addWidget(self.cover_page)
    
    def _create_heading_page(self) -> QWidget:
        return self._create_text_style_page("标题")
    
    def _create_body_page(self) -> QWidget:
        return self._create_text_style_page("正文")
    
    def _create_text_style_page(self, title: str) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        font_group = QGroupBox(f"{title}字体")
        font_layout = QGridLayout()
        font_layout.addWidget(QLabel("中文字体:"), 0, 0)
        font_cn = QComboBox()
        font_cn.setEditable(True)
        font_layout.addWidget(font_cn, 0, 1)
        font_layout.addWidget(QLabel("英文字体:"), 0, 2)
        font_en = QComboBox()
        font_en.setEditable(True)
        font_layout.addWidget(font_en, 0, 3)
        font_layout.addWidget(QLabel("字号:"), 1, 0)
        size_combo = QComboBox()
        for name in FONT_SIZE_NAMES:
            size_combo.addItem(name)
        font_layout.addWidget(size_combo, 1, 1)
        font_layout.addWidget(QLabel("颜色:"), 1, 2)
        color_btn = QPushButton("选择")
        color_btn.setStyleSheet("background-color: #000; color: white;")
        font_layout.addWidget(color_btn, 1, 3)
        bold_cb = QCheckBox("加粗")
        italic_cb = QCheckBox("斜体")
        font_layout.addWidget(bold_cb, 2, 1)
        font_layout.addWidget(italic_cb, 2, 2)
        font_group.setLayout(font_layout)
        scroll_layout.addWidget(font_group)
        
        para_group = QGroupBox("段落设置")
        para_layout = QGridLayout()
        para_layout.addWidget(QLabel("行距类型:"), 0, 0)
        spacing_type = QComboBox()
        for t in LINE_SPACING_TYPES:
            spacing_type.addItem(t)
        para_layout.addWidget(spacing_type, 0, 1)
        para_layout.addWidget(QLabel("行距值:"), 0, 2)
        spacing_val = QDoubleSpinBox()
        spacing_val.setRange(0.5, 100)
        spacing_val.setValue(1.5)
        para_layout.addWidget(spacing_val, 0, 3)
        para_layout.addWidget(QLabel("段前(pt):"), 1, 0)
        space_before = QDoubleSpinBox()
        space_before.setRange(0, 100)
        para_layout.addWidget(space_before, 1, 1)
        para_layout.addWidget(QLabel("段后(pt):"), 1, 2)
        space_after = QDoubleSpinBox()
        space_after.setRange(0, 100)
        para_layout.addWidget(space_after, 1, 3)
        para_layout.addWidget(QLabel("首行缩进:"), 2, 0)
        indent = QDoubleSpinBox()
        indent.setRange(0, 10)
        indent.setValue(2)
        para_layout.addWidget(indent, 2, 1)
        para_layout.addWidget(QLabel("对齐:"), 2, 2)
        align = QComboBox()
        align.addItems(["左对齐", "居中", "右对齐", "两端对齐"])
        para_layout.addWidget(align, 2, 3)
        para_group.setLayout(para_layout)
        scroll_layout.addWidget(para_group)
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        return widget
    
    def _create_table_page(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        type_layout = QHBoxLayout()
        self.table_type_group = QButtonGroup()
        rb1 = QRadioButton("基础表")
        rb2 = QRadioButton("记录表")
        self.table_type_group.addButton(rb1, 0)
        self.table_type_group.addButton(rb2, 1)
        rb1.setChecked(True)
        type_layout.addWidget(rb1)
        type_layout.addWidget(rb2)
        type_layout.addStretch()
        layout.addLayout(type_layout)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        header_group = QGroupBox("表头字体")
        header_layout = QGridLayout()
        header_layout.addWidget(QLabel("字体:"), 0, 0)
        h_font = QComboBox()
        h_font.setEditable(True)
        header_layout.addWidget(h_font, 0, 1)
        header_layout.addWidget(QLabel("字号:"), 0, 2)
        h_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            h_size.addItem(name)
        header_layout.addWidget(h_size, 0, 3)
        h_bold = QCheckBox("加粗")
        header_layout.addWidget(h_bold, 1, 1)
        header_layout.addWidget(QLabel("背景色:"), 1, 2)
        h_bg = QPushButton("选择")
        h_bg.setStyleSheet("background-color: #E8E8E8;")
        header_layout.addWidget(h_bg, 1, 3)
        header_group.setLayout(header_layout)
        scroll_layout.addWidget(header_group)
        body_group = QGroupBox("正文字体")
        body_layout = QGridLayout()
        body_layout.addWidget(QLabel("字体:"), 0, 0)
        b_font = QComboBox()
        b_font.setEditable(True)
        body_layout.addWidget(b_font, 0, 1)
        body_layout.addWidget(QLabel("字号:"), 0, 2)
        b_size = QComboBox()
        for name in FONT_SIZE_NAMES:
            b_size.addItem(name)
        body_layout.addWidget(b_size, 0, 3)
        body_group.setLayout(body_layout)
        scroll_layout.addWidget(body_group)
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        return widget
    
    def _create_caption_page(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        fig_group = QGroupBox("图序")
        fig_layout = QGridLayout()
        fig_layout.addWidget(QLabel("前缀:"), 0, 0)
        fig_prefix = QLineEdit("图")
        fig_layout.addWidget(fig_prefix, 0, 1)
        fig_layout.addWidget(QLabel("位置:"), 0, 2)
        fig_pos = QComboBox()
        fig_pos.addItems(["下方", "上方"])
        fig_layout.addWidget(fig_pos, 0, 3)
        fig_group.setLayout(fig_layout)
        scroll_layout.addWidget(fig_group)
        table_group = QGroupBox("表序")
        table_layout = QGridLayout()
        table_layout.addWidget(QLabel("前缀:"), 0, 0)
        tbl_prefix = QLineEdit("表")
        table_layout.addWidget(tbl_prefix, 0, 1)
        table_layout.addWidget(QLabel("位置:"), 0, 2)
        tbl_pos = QComboBox()
        tbl_pos.addItems(["上方", "下方"])
        table_layout.addWidget(tbl_pos, 0, 3)
        table_group.setLayout(table_layout)
        scroll_layout.addWidget(table_group)
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        return widget
    
    def _create_header_page(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        print_group = QGroupBox("打印模式")
        print_layout = QHBoxLayout()
        self.print_group = QButtonGroup()
        rb_single = QRadioButton("单面打印")
        rb_duplex = QRadioButton("双面打印")
        self.print_group.addButton(rb_single, 0)
        self.print_group.addButton(rb_duplex, 1)
        rb_single.setChecked(True)
        print_layout.addWidget(rb_single)
        print_layout.addWidget(rb_duplex)
        print_layout.addStretch()
        print_group.setLayout(print_layout)
        scroll_layout.addWidget(print_group)
        hf_group = QGroupBox("页眉页脚设置")
        hf_layout = QVBoxLayout()
        cb1 = QCheckBox("首页不同")
        cb2 = QCheckBox("奇偶页不同")
        hf_layout.addWidget(cb1)
        hf_layout.addWidget(cb2)
        hf_group.setLayout(hf_layout)
        scroll_layout.addWidget(hf_group)
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        return widget
    
    def _create_cover_page(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel("封面字段（占位符 → 值）"))
        self.cover_table = QTableWidget(15, 2)
        self.cover_table.setHorizontalHeaderLabels(["字段名", "值"])
        fields = ["contract_no", "classification", "year", "organization",
                  "date", "drafter", "reviewer", "approver", "checker",
                  "final_approver", "drafter_time", "reviewer_time",
                  "approver_time", "checker_time", "final_approver_time"]
        for i, f in enumerate(fields):
            item = QTableWidgetItem(f)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.cover_table.setItem(i, 0, item)
        layout.addWidget(self.cover_table)
        return widget
    
    def _connect_signals(self):
        self.config_list.currentRowChanged.connect(self._on_item_changed)
    
    def _on_item_changed(self, row):
        item = self.config_list.item(row)
        item_id = item.data(Qt.ItemDataRole.UserRole)
        if item_id and item_id.startswith("heading"):
            self.right_stack.setCurrentWidget(self.heading_page)
        elif item_id == "body":
            self.right_stack.setCurrentWidget(self.body_page)
        elif item_id.startswith("table_"):
            self.right_stack.setCurrentWidget(self.table_page)
        elif item_id.startswith("caption_"):
            self.right_stack.setCurrentWidget(self.caption_page)
        elif item_id == "header_all":
            self.right_stack.setCurrentWidget(self.header_page)
        elif item_id == "cover_fields":
            self.right_stack.setCurrentWidget(self.cover_page)
    
    def new_template(self):
        self.template = create_default_template()
        self.template_path = None
    
    def load_template(self):
        file, _ = QFileDialog.getOpenFileName(self, "加载模板", "", "JSON文件 (*.json)")
        if file:
            try:
                from ..templates.template_io import load_template
                self.template = load_template(file)
                self.template_path = file
                QMessageBox.information(self, "成功", f"已加载:\n{file}")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"加载失败:\n{str(e)}")
    
    def save_template(self):
        file, _ = QFileDialog.getSaveFileName(self, "保存模板", "template.json", "JSON文件 (*.json)")
        if file:
            try:
                from ..templates.template_io import save_template
                save_template(self.template, file)
                self.template_path = file
                QMessageBox.information(self, "成功", f"已保存:\n{file}")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"保存失败:\n{str(e)}")
    
    def get_template(self) -> TemplateConfig:
        return self.template
