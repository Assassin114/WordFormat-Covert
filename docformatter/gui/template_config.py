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
    QTableWidget, QTableWidgetItem, QHeaderView, QTabWidget
)
from PyQt6.QtCore import Qt, pyqtSignal

from ..models.template_config import create_default_template, TemplateConfig
from ..templates.template_io import load_template, save_template


class TemplateConfigWidget(QWidget):
    """
    模板配置界面
    """
    
    def __init__(self):
        super().__init__()
        self.template = create_default_template()
        self.template_path = None
        self._modified = False
        
        self._init_ui()
    
    def _init_ui(self):
        """初始化UI"""
        layout = QVBoxLayout(self)
        
        # 创建标签页
        tabs = QTabWidget()
        
        # Tab 1: 基础配置
        tab_basic = QWidget()
        tab_basic_layout = QVBoxLayout(tab_basic)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # ===== 封面配置 =====
        cover_group = QGroupBox("封面配置")
        cover_layout = QVBoxLayout()
        
        self.cover_enabled = QCheckBox("启用封面")
        self.cover_enabled.setChecked(True)
        cover_layout.addWidget(self.cover_enabled)
        
        cover_btn_layout = QHBoxLayout()
        self.btn_select_cover = QPushButton("选择封面模板")
        self.btn_select_cover.clicked.connect(self._select_cover_template)
        cover_btn_layout.addWidget(self.btn_select_cover)
        
        self.cover_path_label = QLabel("未选择")
        self.cover_path_label.setStyleSheet("color: gray;")
        cover_btn_layout.addWidget(self.cover_path_label)
        cover_btn_layout.addStretch()
        
        cover_layout.addLayout(cover_btn_layout)
        
        # 封面字段编辑按钮
        self.btn_edit_cover_fields = QPushButton("编辑封面字段...")
        self.btn_edit_cover_fields.clicked.connect(self._edit_cover_fields)
        cover_layout.addWidget(self.btn_edit_cover_fields)
        
        cover_group.setLayout(cover_layout)
        scroll_layout.addWidget(cover_group)
        
        # ===== 签署页配置 =====
        sig_group = QGroupBox("签署页配置")
        sig_layout = QVBoxLayout()
        
        self.sig_enabled = QCheckBox("启用签署页")
        self.sig_enabled.setChecked(True)
        sig_layout.addWidget(self.sig_enabled)
        
        sig_btn_layout = QHBoxLayout()
        self.btn_select_sig = QPushButton("选择签署页模板")
        self.btn_select_sig.clicked.connect(self._select_sig_template)
        sig_btn_layout.addWidget(self.btn_select_sig)
        
        self.sig_path_label = QLabel("未选择")
        self.sig_path_label.setStyleSheet("color: gray;")
        sig_btn_layout.addWidget(self.sig_path_label)
        sig_btn_layout.addStretch()
        
        sig_layout.addLayout(sig_btn_layout)
        
        sig_group.setLayout(sig_layout)
        scroll_layout.addWidget(sig_group)
        
        # ===== 标题格式 =====
        heading_group = QGroupBox("标题格式 (1-5级)")
        heading_layout = QVBoxLayout()
        
        self.heading_editors = []
        for level in range(1, 6):
            editor = HeadingLevelEditor(level)
            editor.value_changed.connect(self._on_heading_changed)
            heading_layout.addWidget(editor)
            self.heading_editors.append(editor)
        
        heading_group.setLayout(heading_layout)
        scroll_layout.addWidget(heading_group)
        
        # ===== 正文格式 =====
        body_group = QGroupBox("正文格式")
        body_layout = QFormLayout()
        
        self.body_font = QLineEdit("宋体")
        body_layout.addRow("字体:", self.body_font)
        
        self.body_size = QDoubleSpinBox()
        self.body_size.setRange(9, 72)
        self.body_size.setValue(10.5)
        self.body_size.setSuffix(" pt")
        body_layout.addRow("字号:", self.body_size)
        
        self.body_spacing = QDoubleSpinBox()
        self.body_spacing.setRange(0.5, 4.0)
        self.body_spacing.setValue(1.5)
        self.body_spacing.setSingleStep(0.1)
        body_layout.addRow("行距:", self.body_spacing)
        
        self.body_indent = QDoubleSpinBox()
        self.body_indent.setRange(0, 10)
        self.body_indent.setValue(2)
        self.body_indent.setSuffix(" 字符")
        body_layout.addRow("首行缩进:", self.body_indent)
        
        body_group.setLayout(body_layout)
        scroll_layout.addWidget(body_group)
        
        # ===== 题注格式 =====
        caption_group = QGroupBox("题注格式")
        caption_layout = QFormLayout()
        
        h_layout = QHBoxLayout()
        h_layout.addWidget(QLabel("图序前缀:"))
        self.caption_figure = QLineEdit("图")
        h_layout.addWidget(self.caption_figure)
        h_layout.addWidget(QLabel("位置:"))
        self.caption_figure_pos = QComboBox()
        self.caption_figure_pos.addItems(["下方", "上方"])
        h_layout.addWidget(self.caption_figure_pos)
        h_layout.addStretch()
        caption_layout.addRow(h_layout)
        
        t_layout = QHBoxLayout()
        t_layout.addWidget(QLabel("表序前缀:"))
        self.caption_table = QLineEdit("表")
        t_layout.addWidget(self.caption_table)
        t_layout.addWidget(QLabel("位置:"))
        self.caption_table_pos = QComboBox()
        self.caption_table_pos.addItems(["上方", "下方"])
        t_layout.addWidget(self.caption_table_pos)
        t_layout.addStretch()
        caption_layout.addRow(t_layout)
        
        caption_group.setLayout(caption_layout)
        scroll_layout.addWidget(caption_group)
        
        # ===== 页眉页脚 =====
        hf_group = QGroupBox("页眉页脚")
        hf_layout = QVBoxLayout()
        
        self.hf_first_diff = QCheckBox("首页不同")
        hf_layout.addWidget(self.hf_first_diff)
        
        self.hf_odd_even = QCheckBox("奇偶页不同")
        hf_layout.addWidget(self.hf_odd_even)
        
        hf_group.setLayout(hf_layout)
        scroll_layout.addWidget(hf_group)
        
        # ===== 打印模式 =====
        print_group = QGroupBox("打印模式")
        print_layout = QHBoxLayout()
        
        self.print_single = QCheckBox("单面打印")
        self.print_single.setChecked(True)
        print_layout.addWidget(self.print_single)
        
        self.print_duplex = QCheckBox("双面打印")
        print_layout.addWidget(self.print_duplex)
        
        print_group.setLayout(print_layout)
        scroll_layout.addWidget(print_group)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        tab_basic_layout.addWidget(scroll)
        
        # Tab 2: 封面字段
        tab_cover = QWidget()
        tab_cover_layout = QVBoxLayout(tab_cover)
        
        cover_fields_label = QLabel("封面字段编辑（用于替换模板中的占位符）")
        cover_fields_label.setStyleSheet("color: gray; padding: 5px;")
        tab_cover_layout.addWidget(cover_fields_label)
        
        # 封面字段表格
        self.cover_fields_table = QTableWidget()
        self.cover_fields_table.setColumnCount(2)
        self.cover_fields_table.setHorizontalHeaderLabels(["字段名", "值"])
        self.cover_fields_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.cover_fields_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._init_cover_fields_table()
        tab_cover_layout.addWidget(self.cover_fields_table)
        
        # 按钮行
        btn_row = QHBoxLayout()
        btn_add_field = QPushButton("添加字段")
        btn_add_field.clicked.connect(self._add_cover_field)
        btn_row.addWidget(btn_add_field)
        
        btn_remove_field = QPushButton("删除选中")
        btn_remove_field.clicked.connect(self._remove_cover_field)
        btn_row.addWidget(btn_remove_field)
        
        btn_row.addStretch()
        tab_cover_layout.addLayout(btn_row)
        
        # 预定义字段按钮
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
        
        for label, key in presets:
            btn = QPushButton(label)
            btn.clicked.connect(lambda checked, k=key: self._insert_preset_field(k))
            preset_layout.addWidget(btn)
        
        preset_layout.addStretch()
        preset_group.setLayout(preset_layout)
        tab_cover_layout.addWidget(preset_group)
        
        # 添加标签页
        tabs.addTab(tab_basic, "基础配置")
        tabs.addTab(tab_cover, "封面字段")
        
        layout.addWidget(tabs)
        
        # ===== 底部按钮 =====
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
    
    def _init_cover_fields_table(self):
        """初始化封面字段表格"""
        # 标准封面字段
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
            key_item.setFlags(key_item.flags() & ~Qt.ItemFlag.ItemIsEditable)  # 只读
            self.cover_fields_table.setItem(i, 0, key_item)
            
            value_item = QTableWidgetItem("")
            value_item.setText(self.template.cover.fields.get(key, ""))
            self.cover_fields_table.setItem(i, 1, value_item)
    
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
        # 查找是否已存在
        for row in range(self.cover_fields_table.rowCount()):
            if self.cover_fields_table.item(row, 0).text() == key:
                self.cover_fields_table.selectRow(row)
                return
        
        # 添加新行
        row = self.cover_fields_table.rowCount()
        self.cover_fields_table.insertRow(row)
        key_item = QTableWidgetItem(key)
        key_item.setFlags(key_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        self.cover_fields_table.setItem(row, 0, key_item)
        self.cover_fields_table.setItem(row, 1, QTableWidgetItem(""))
        self.cover_fields_table.selectRow(row)
    
    def _edit_cover_fields(self):
        """编辑封面字段（切换到字段Tab）"""
        # 找到父窗口的tab_widget并切换
        parent = self.parent()
        while parent and not hasattr(parent, 'tab_widget'):
            parent = parent.parent()
        if parent and hasattr(parent, 'tab_widget'):
            parent.tab_widget.setCurrentIndex(1)  # 切换到封面字段Tab
    
    def _select_cover_template(self):
        """选择封面模板"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "选择封面模板",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )
        
        if file:
            self.template.cover.template_path = file
            self.cover_path_label.setText(Path(file).name)
            self._modified = True
    
    def _select_sig_template(self):
        """选择签署页模板"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "选择签署页模板",
            "",
            "Word文档 (*.docx);;所有文件 (*.*)"
        )
        
        if file:
            self.template.signature.template_path = file
            self.sig_path_label.setText(Path(file).name)
            self._modified = True
    
    def _on_heading_changed(self):
        """标题格式变更"""
        self._modified = True
        self._sync_template()
    
    def _sync_template(self):
        """同步UI到模板对象"""
        # 同步标题格式
        for i, editor in enumerate(self.heading_editors):
            h = self.template.headings[i]
            h.font.name = editor.font_name
            h.font.size = editor.font_size
            h.font.bold = editor.font_bold
            h.paragraph.line_spacing = editor.line_spacing
            h.paragraph.space_before = editor.space_before
            h.paragraph.space_after = editor.space_after
        
        # 同步正文格式
        self.template.body.font.name = self.body_font.text()
        self.template.body.font.size = self.body_size.value()
        self.template.body.line_spacing = self.body_spacing.value()
        self.template.body.first_line_indent = self.body_indent.value() * 10
    
    def _sync_cover_fields_to_template(self):
        """同步封面字段到模板"""
        self.template.cover.fields = {}
        for row in range(self.cover_fields_table.rowCount()):
            key_item = self.cover_fields_table.item(row, 0)
            value_item = self.cover_fields_table.item(row, 1)
            if key_item and key_item.text().strip():
                key = key_item.text().strip()
                value = value_item.text() if value_item else ""
                self.template.cover.fields[key] = value
    
    def new_template(self):
        """新建模板"""
        self.template = create_default_template()
        self.template_path = None
        self._sync_ui_from_template()
        self._modified = False
    
    def load_template(self):
        """加载模板"""
        file, _ = QFileDialog.getOpenFileName(
            self,
            "加载模板",
            "",
            "JSON文件 (*.json);;所有文件 (*.*)"
        )
        
        if file:
            try:
                self.template = load_template(file)
                self.template_path = file
                self._sync_ui_from_template()
                self._modified = False
                QMessageBox.information(self, "成功", f"已加载模板:\n{file}")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"加载模板失败:\n{str(e)}")
    
    def save_template(self):
        """保存模板"""
        self._sync_template()
        self._sync_cover_fields_to_template()
        
        file, _ = QFileDialog.getSaveFileName(
            self,
            "保存模板",
            "template.json",
            "JSON文件 (*.json);;所有文件 (*.*)"
        )
        
        if file:
            try:
                save_template(self.template, file)
                self.template_path = file
                self._modified = False
                QMessageBox.information(self, "成功", f"已保存模板:\n{file}")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"保存模板失败:\n{str(e)}")
    
    def get_template(self) -> TemplateConfig:
        """获取当前模板"""
        self._sync_template()
        self._sync_cover_fields_to_template()
        return self.template
    
    def _sync_ui_from_template(self):
        """同步模板对象到UI"""
        # 同步封面路径
        if self.template.cover.template_path:
            self.cover_path_label.setText(Path(self.template.cover.template_path).name)
        
        # 同步签署页路径
        if self.template.signature.template_path:
            self.sig_path_label.setText(Path(self.template.signature.template_path).name)
        
        # 同步封面字段
        for row in range(self.cover_fields_table.rowCount()):
            key = self.cover_fields_table.item(row, 0).text()
            if key and key in self.template.cover.fields:
                self.cover_fields_table.item(row, 1).setText(self.template.cover.fields[key])
        
        # 同步标题格式
        for i, editor in enumerate(self.heading_editors):
            if i < len(self.template.headings):
                h = self.template.headings[i]
                editor.font_name = h.font.name
                editor.font_size = h.font.size
                editor.font_bold = h.font.bold
                editor.line_spacing = h.paragraph.line_spacing
                editor.space_before = h.paragraph.space_before
                editor.space_after = h.paragraph.space_after
        
        # 同步正文格式
        self.body_font.setText(self.template.body.font.name)
        self.body_size.setValue(self.template.body.font.size)
        self.body_spacing.setValue(self.template.body.line_spacing)
        self.body_indent.setValue(self.template.body.first_line_indent / 10)


class HeadingLevelEditor(QWidget):
    """单级标题格式编辑器"""
    
    value_changed = pyqtSignal()
    
    def __init__(self, level: int):
        super().__init__()
        self.level = level
        self._font_name = "黑体"
        self._font_size = 16
        self._font_bold = False
        self._line_spacing = 1.5
        self._space_before = 12
        self._space_after = 6
        
        self._init_ui()
    
    def _init_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 5, 0, 5)
        
        # 级别标签
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
        self.size_spin = QDoubleSpinBox()
        self.size_spin.setRange(9, 48)
        self.size_spin.setValue(self._font_size)
        self.size_spin.setSuffix(" pt")
        self.size_spin.valueChanged.connect(self._on_changed)
        layout.addWidget(QLabel("字号:"))
        layout.addWidget(self.size_spin)
        
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
        """值变更"""
        self._font_name = self.font_combo.currentText()
        self._font_size = self.size_spin.value()
        self._font_bold = self.bold_check.isChecked()
        self._line_spacing = self.spacing_spin.value()
        self.value_changed.emit()
    
    @property
    def font_name(self):
        return self._font_name
    
    @font_name.setter
    def font_name(self, value):
        self._font_name = value
        self.font_combo.setCurrentText(value)
    
    @property
    def font_size(self):
        return self._font_size
    
    @font_size.setter
    def font_size(self, value):
        self._font_size = value
        self.size_spin.setValue(value)
    
    @property
    def font_bold(self):
        return self._font_bold
    
    @font_bold.setter
    def font_bold(self, value):
        self._font_bold = value
        self.bold_check.setChecked(value)
    
    @property
    def line_spacing(self):
        return self._line_spacing
    
    @line_spacing.setter
    def line_spacing(self, value):
        self._line_spacing = value
        self.spacing_spin.setValue(value)
    
    @property
    def space_before(self):
        return self._space_before
    
    @space_before.setter
    def space_before(self, value):
        self._space_before = value
    
    @property
    def space_after(self):
        return self._space_after
    
    @space_after.setter
    def space_after(self, value):
        self._space_after = value
