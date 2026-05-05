"""
模板配置 Tab — 左侧项目列表 + 中间配置面板 + 右侧预览
"""

from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QScrollArea,
    QLabel, QPushButton, QGroupBox, QGridLayout,
    QLineEdit, QDoubleSpinBox, QComboBox,
    QCheckBox, QFileDialog, QMessageBox,
    QListWidget, QListWidgetItem, QStackedWidget,
    QColorDialog, QButtonGroup, QRadioButton,
    QFrame, QTableWidget, QTableWidgetItem, QHeaderView,
    QTextEdit, QSplitter, QSpinBox,
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFontDatabase

from ..models.template_config import (
    TemplateConfig, FontConfig, ParagraphConfig,
    create_default_template, PageNumberFormat,
    PageNumberPosition, PageNumberAlignment, EquationFormat,
    PrintMode, NumberingMode,
)
from ..templates.template_manager import TemplateManager
from ..utils.placeholder_scanner import scan_placeholders


# ══════════════════════════════════════════════════════════════
# 常量
# ══════════════════════════════════════════════════════════════

CHINESE_FONT_SIZES = {
    "初号": 42, "小初": 36, "一号": 26, "小一": 24,
    "二号": 22, "小二": 18, "三号": 16, "小三": 15,
    "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
}
FONT_SIZE_NAMES = list(CHINESE_FONT_SIZES.keys())
LINE_SPACING_TYPES = ["多倍行距", "单倍行距", "固定值", "最小值"]
ALIGNMENT_MAP = {"左对齐": "left", "居中": "center", "右对齐": "right", "两端对齐": "both"}
ALIGNMENT_MAP_REV = {"left": "左对齐", "center": "居中", "right": "右对齐", "both": "两端对齐"}
PAGE_NUM_FORMAT_MAP = {"阿拉伯数字": "arabic", "大写罗马": "roman_up", "小写罗马": "roman_low", "大写字母": "letter_up", "小写字母": "letter_low"}
PAGE_NUM_FORMAT_MAP_REV = {v: k for k, v in PAGE_NUM_FORMAT_MAP.items()}
PAGE_NUM_POS_MAP = {"页脚": "footer", "页眉": "header"}
PAGE_NUM_POS_MAP_REV = {v: k for k, v in PAGE_NUM_POS_MAP.items()}
PAGE_NUM_ALIGN_MAP = {"左对齐": "left", "居中": "center", "右对齐": "right"}
PAGE_NUM_ALIGN_MAP_REV = {v: k for k, v in PAGE_NUM_ALIGN_MAP.items()}

PAPER_PRESETS = {
    "A4": (210, 297),
    "A3": (297, 420),
    "A5": (148, 210),
    "B5": (176, 250),
    "Letter": (215.9, 279.4),
    "Legal": (215.9, 355.6),
    "自定义": None,
}

NUMBER_STYLE_MAP = {
    "无编号": "none",
    "一、二、三": "chinese",
    "1. / 2.": "decimal",
    "A. / B.": "upper_letter",
    "a. / b.": "lower_letter",
    "I. / II.": "upper_roman",
    "i. / ii.": "lower_roman",
}
NUMBER_STYLE_MAP_REV = {v: k for k, v in NUMBER_STYLE_MAP.items()}
NUMBER_STYLE_KEYS = list(NUMBER_STYLE_MAP.keys())


# ══════════════════════════════════════════════════════════════
# 字体缓存
# ══════════════════════════════════════════════════════════════

def _get_system_fonts():
    """获取系统字体列表（缓存）"""
    if not hasattr(_get_system_fonts, "_cache"):
        _get_system_fonts._cache = sorted(set(QFontDatabase.families()))
    return _get_system_fonts._cache


def _populate_font_combo(combo: QComboBox, default_cn=True):
    """填充字体下拉框，常用中文字体前置"""
    fonts = _get_system_fonts()
    combo.addItems(fonts)


# ══════════════════════════════════════════════════════════════
# 通用组件
# ══════════════════════════════════════════════════════════════

class _FontParagraphGroup(QGroupBox):
    """字体+段落合并编辑组件 — 一个 QGroupBox 内包含两者"""
    changed = pyqtSignal()

    def __init__(self, title: str = "字体与段落"):
        super().__init__(title)
        layout = QGridLayout(self)

        # ── 字体行 ──
        font_sep = QLabel("<b>字体</b>")
        layout.addWidget(font_sep, 0, 0, 1, 5)

        layout.addWidget(QLabel("中文字体:"), 1, 0)
        self.cn_font = QComboBox()
        self.cn_font.setEditable(True)
        self.cn_font.setMinimumWidth(120)
        _populate_font_combo(self.cn_font)
        layout.addWidget(self.cn_font, 1, 1)

        layout.addWidget(QLabel("英文字体:"), 1, 2)
        self.en_font = QComboBox()
        self.en_font.setEditable(True)
        self.en_font.setMinimumWidth(120)
        _populate_font_combo(self.en_font, default_cn=False)
        layout.addWidget(self.en_font, 1, 3)

        layout.addWidget(QLabel("字号:"), 2, 0)
        self.size_combo = QComboBox()
        self.size_combo.setMaximumWidth(80)
        for name in FONT_SIZE_NAMES:
            self.size_combo.addItem(name)
        layout.addWidget(self.size_combo, 2, 1)

        layout.addWidget(QLabel("颜色:"), 2, 2)
        self.color_btn = QPushButton("选择")
        self.color_btn.setMaximumWidth(60)
        self.color_btn.clicked.connect(self._pick_color)
        layout.addWidget(self.color_btn, 2, 3)

        self.bold_cb = QCheckBox("加粗")
        self.italic_cb = QCheckBox("斜体")
        layout.addWidget(self.bold_cb, 3, 1)
        layout.addWidget(self.italic_cb, 3, 2)

        # ── 段落行 ──
        para_sep = QLabel("<b>段落</b>")
        layout.addWidget(para_sep, 4, 0, 1, 5)

        layout.addWidget(QLabel("行距类型:"), 5, 0)
        self.spacing_type = QComboBox()
        self.spacing_type.setMaximumWidth(90)
        for t in LINE_SPACING_TYPES:
            self.spacing_type.addItem(t)
        layout.addWidget(self.spacing_type, 5, 1)

        layout.addWidget(QLabel("行距值:"), 5, 2)
        self.spacing_val = QDoubleSpinBox()
        self.spacing_val.setRange(0.5, 200)
        self.spacing_val.setValue(1.5)
        layout.addWidget(self.spacing_val, 5, 3)
        self.spacing_unit = QLabel("倍")
        layout.addWidget(self.spacing_unit, 5, 4)

        layout.addWidget(QLabel("段前(磅):"), 6, 0)
        self.space_before = QDoubleSpinBox()
        self.space_before.setRange(0, 1000)
        layout.addWidget(self.space_before, 6, 1)

        layout.addWidget(QLabel("段后(磅):"), 6, 2)
        self.space_after = QDoubleSpinBox()
        self.space_after.setRange(0, 1000)
        layout.addWidget(self.space_after, 6, 3)

        layout.addWidget(QLabel("首行缩进(字符):"), 7, 0)
        self.indent = QDoubleSpinBox()
        self.indent.setRange(0, 20)
        self.indent.setValue(2)
        layout.addWidget(self.indent, 7, 1)

        layout.addWidget(QLabel("对齐方式:"), 7, 2)
        self.align = QComboBox()
        self.align.addItems(list(ALIGNMENT_MAP.keys()))
        layout.addWidget(self.align, 7, 3)

        self.spacing_type.currentIndexChanged.connect(self._on_type_changed)
        self._current_color = "#000000"
        self.color_btn.setStyleSheet("background-color: #000000;")

        # 所有控件变更 → 发射 changed 信号
        for w in [self.cn_font, self.en_font, self.size_combo, self.spacing_type, self.align]:
            w.currentTextChanged.connect(self._emit_changed)
        for w in [self.spacing_val, self.space_before, self.space_after, self.indent]:
            w.valueChanged.connect(self._emit_changed)
        self.bold_cb.toggled.connect(self._emit_changed)
        self.italic_cb.toggled.connect(self._emit_changed)

    def _emit_changed(self, *_):
        self.changed.emit()

    def _pick_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self._current_color = color.name()
            self.color_btn.setStyleSheet(f"background-color: {self._current_color};")
            self._emit_changed()

    def _on_type_changed(self, idx):
        tp = LINE_SPACING_TYPES[idx]
        if tp == "单倍行距":
            self.spacing_val.setEnabled(False)
            self.spacing_unit.setText("")
        else:
            self.spacing_val.setEnabled(True)
            if tp == "多倍行距":
                self.spacing_unit.setText("倍")
            elif tp in ("固定值", "最小值"):
                self.spacing_unit.setText("磅")

    def _spacing_type_key(self) -> str:
        tp = self.spacing_type.currentText()
        return {"多倍行距": "multiple", "单倍行距": "single", "固定值": "fixed", "最小值": "atLeast"}[tp]

    def get_font_config(self) -> FontConfig:
        size_name = self.size_combo.currentText()
        return FontConfig(
            cn_name=self.cn_font.currentText() or "宋体",
            en_name=self.en_font.currentText() or "Times New Roman",
            size=CHINESE_FONT_SIZES.get(size_name, 10.5),
            size_name=size_name,
            bold=self.bold_cb.isChecked(),
            italic=self.italic_cb.isChecked(),
            color=self._current_color,
        )

    def set_font_config(self, f: FontConfig):
        self.cn_font.setCurrentText(f.cn_name)
        self.en_font.setCurrentText(f.en_name)
        idx = self.size_combo.findText(f.size_name)
        if idx >= 0:
            self.size_combo.setCurrentIndex(idx)
        self.bold_cb.setChecked(f.bold)
        self.italic_cb.setChecked(f.italic)
        self._current_color = f.color
        self.color_btn.setStyleSheet(f"background-color: {f.color};")

    def get_paragraph_config(self) -> ParagraphConfig:
        val = self.spacing_val.value()
        tp = self._spacing_type_key()
        return ParagraphConfig(
            line_spacing=val,
            line_spacing_type=tp,
            line_spacing_fixed=val if tp == "fixed" else 20,
            line_spacing_min=val if tp == "atLeast" else 15,
            space_before=self.space_before.value(),
            space_after=self.space_after.value(),
            first_line_indent=self.indent.value(),
            alignment=ALIGNMENT_MAP.get(self.align.currentText(), "left"),
        )

    def set_paragraph_config(self, p: ParagraphConfig):
        tp_map = {"multiple": 0, "single": 1, "fixed": 2, "atLeast": 3}
        self.spacing_type.setCurrentIndex(tp_map.get(p.line_spacing_type, 0))
        val = p.line_spacing
        if p.line_spacing_type == "fixed":
            val = p.line_spacing_fixed
        elif p.line_spacing_type == "atLeast":
            val = p.line_spacing_min
        self.spacing_val.setValue(val)
        self.space_before.setValue(p.space_before)
        self.space_after.setValue(p.space_after)
        self.indent.setValue(p.first_line_indent)
        idx = self.align.findText(ALIGNMENT_MAP_REV.get(p.alignment, "左对齐"))
        if idx >= 0:
            self.align.setCurrentIndex(idx)

    def preview_html(self, scale: float = 1.0) -> str:
        """生成该字体/段落配置的 HTML 预览"""
        f = self.get_font_config()
        p = self.get_paragraph_config()
        align = ALIGNMENT_MAP_REV.get(p.alignment, "左对齐")
        weight = "bold" if f.bold else "normal"
        style = "italic" if f.italic else "normal"
        indent = f"text-indent: {p.first_line_indent * 16 * scale}px;" if p.first_line_indent else ""
        line_h = p.line_spacing if p.line_spacing_type == "multiple" else 1.5
        return f"""<div style="font-family: '{f.cn_name}', '{f.en_name}', serif;
                    font-size: {f.size * scale:.0f}pt; font-weight: {weight}; font-style: {style};
                    color: {f.color}; text-align: {align};
                    line-height: {line_h}; {indent}">
                    这是一段示例文本，用于预览字体和段落格式效果。<br>
                    The quick brown fox jumps over the lazy dog. 1234567890
                    </div>"""


# ══════════════════════════════════════════════════════════════
# 各配置页面
# ══════════════════════════════════════════════════════════════

class _PageConfigPage(QWidget):
    """页面设置 — 纸张预设 + 页边距 + 装订线 + 页码"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        w = QWidget()
        sl = QVBoxLayout(w)

        # 纸张预设
        pg1 = QGroupBox("纸张")
        g1 = QGridLayout(pg1)
        g1.addWidget(QLabel("预设:"), 0, 0)
        self.preset_combo = QComboBox()
        self.preset_combo.addItems(list(PAPER_PRESETS.keys()))
        g1.addWidget(self.preset_combo, 0, 1)
        g1.addWidget(QLabel("宽度(mm):"), 0, 2)
        self.page_width = QDoubleSpinBox()
        self.page_width.setRange(50, 1000)
        self.page_width.setValue(210)
        g1.addWidget(self.page_width, 0, 3)
        g1.addWidget(QLabel("高度(mm):"), 0, 4)
        self.page_height = QDoubleSpinBox()
        self.page_height.setRange(50, 1000)
        self.page_height.setValue(297)
        g1.addWidget(self.page_height, 0, 5)
        g1.addWidget(QLabel("方向:"), 1, 0)
        self.orientation_cb = QComboBox()
        self.orientation_cb.addItems(["纵向", "横向"])
        g1.addWidget(self.orientation_cb, 1, 1)
        self.preset_combo.currentIndexChanged.connect(self._on_preset_changed)
        sl.addWidget(pg1)

        # 页边距
        pg2 = QGroupBox("页边距")
        g2 = QGridLayout(pg2)
        g2.addWidget(QLabel("上(mm):"), 0, 0)
        self.margin_top = QDoubleSpinBox()
        self.margin_top.setRange(0, 100)
        self.margin_top.setValue(37)
        g2.addWidget(self.margin_top, 0, 1)
        g2.addWidget(QLabel("下(mm):"), 0, 2)
        self.margin_bottom = QDoubleSpinBox()
        self.margin_bottom.setRange(0, 100)
        self.margin_bottom.setValue(35)
        g2.addWidget(self.margin_bottom, 0, 3)
        g2.addWidget(QLabel("左(mm):"), 1, 0)
        self.margin_left = QDoubleSpinBox()
        self.margin_left.setRange(0, 100)
        self.margin_left.setValue(28)
        g2.addWidget(self.margin_left, 1, 1)
        g2.addWidget(QLabel("右(mm):"), 1, 2)
        self.margin_right = QDoubleSpinBox()
        self.margin_right.setRange(0, 100)
        self.margin_right.setValue(26)
        g2.addWidget(self.margin_right, 1, 3)
        sl.addWidget(pg2)

        # 装订线
        pg3 = QGroupBox("装订线")
        g3 = QGridLayout(pg3)
        g3.addWidget(QLabel("宽度(mm):"), 0, 0)
        self.gutter = QDoubleSpinBox()
        self.gutter.setRange(0, 50)
        self.gutter.setValue(0)
        g3.addWidget(self.gutter, 0, 1)
        g3.addWidget(QLabel("位置:"), 0, 2)
        self.gutter_pos = QComboBox()
        self.gutter_pos.addItems(["左", "上"])
        g3.addWidget(self.gutter_pos, 0, 3)
        sl.addWidget(pg3)

        # 页码
        pg4 = QGroupBox("页码")
        g4 = QGridLayout(pg4)
        g4.addWidget(QLabel("起始页码:"), 0, 0)
        self.page_num_start = QSpinBox()
        self.page_num_start.setRange(0, 9999)
        self.page_num_start.setValue(1)
        g4.addWidget(self.page_num_start, 0, 1)
        sl.addWidget(pg4)
        sl.addStretch()

        scroll.setWidget(w)
        layout.addWidget(scroll)

    def _on_preset_changed(self, idx):
        key = list(PAPER_PRESETS.keys())[idx]
        if key == "自定义":
            return
        w, h = PAPER_PRESETS[key]
        self.page_width.setValue(w)
        self.page_height.setValue(h)

    def load_config(self, t: TemplateConfig):
        pc = t.page
        self.page_width.setValue(pc.width)
        self.page_height.setValue(pc.height)
        # 匹配预设
        matched = False
        for i, (key, dims) in enumerate(PAPER_PRESETS.items()):
            if dims and abs(dims[0] - pc.width) < 0.1 and abs(dims[1] - pc.height) < 0.1:
                self.preset_combo.setCurrentIndex(i)
                matched = True
                break
        if not matched:
            self.preset_combo.setCurrentText("自定义")
        self.orientation_cb.setCurrentText("纵向" if pc.orientation == "portrait" else "横向")
        self.margin_top.setValue(pc.margin_top)
        self.margin_bottom.setValue(pc.margin_bottom)
        self.margin_left.setValue(pc.margin_left)
        self.margin_right.setValue(pc.margin_right)
        self.gutter.setValue(pc.gutter)
        self.gutter_pos.setCurrentText("左" if pc.gutter_position == "left" else "上")
        self.page_num_start.setValue(pc.page_number_start)

    def save_config(self, t: TemplateConfig):
        pc = t.page
        pc.width = self.page_width.value()
        pc.height = self.page_height.value()
        pc.orientation = "portrait" if self.orientation_cb.currentText() == "纵向" else "landscape"
        pc.margin_top = self.margin_top.value()
        pc.margin_bottom = self.margin_bottom.value()
        pc.margin_left = self.margin_left.value()
        pc.margin_right = self.margin_right.value()
        pc.gutter = self.gutter.value()
        pc.gutter_position = "left" if self.gutter_pos.currentText() == "左" else "top"
        pc.page_number_start = int(self.page_num_start.value())

    def preview_html(self, scale: float = 1.0) -> str:
        w = self.page_width.value()
        h = self.page_height.value()
        orientation = self.orientation_cb.currentText()
        mt = self.margin_top.value()
        mb = self.margin_bottom.value()
        ml = self.margin_left.value()
        mr = self.margin_right.value()
        gutter = self.gutter.value()
        return (
            f"<div style='font-size: {10 * scale:.0f}pt; line-height: 1.6;'>"
            f"<b>页面设置预览</b><br>"
            f"纸张: {w:.0f}×{h:.0f}mm {orientation}<br>"
            f"页边距: 上{mt:.0f} 下{mb:.0f} 左{ml:.0f} 右{mr:.0f}mm<br>"
            f"装订线: {gutter:.0f}mm<br>"
            f"起始页码: {self.page_num_start.value()}"
            f"</div>"
        )


class _HeaderFooterPage(QWidget):
    """页眉页脚 — 打印模式 + 公式格式 + 三类页面页眉页脚"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        w = QWidget()
        sl = QVBoxLayout(w)

        # 打印模式 + 公式格式
        top = QHBoxLayout()
        pg = QGroupBox("打印模式")
        pl = QHBoxLayout()
        self.print_group = QButtonGroup()
        rb_s = QRadioButton("单面打印")
        rb_d = QRadioButton("双面打印")
        self.print_group.addButton(rb_s, 0)
        self.print_group.addButton(rb_d, 1)
        rb_s.setChecked(True)
        pl.addWidget(rb_s)
        pl.addWidget(rb_d)
        pl.addStretch()
        pg.setLayout(pl)
        top.addWidget(pg)

        eqg = QGroupBox("公式格式")
        eql = QHBoxLayout()
        eql.addWidget(QLabel("公式处理:"))
        self.eq_format = QComboBox()
        self.eq_format.addItems(["保留原格式", "强制转为Word可编辑公式", "渲染为图片"])
        eql.addWidget(self.eq_format)
        eql.addStretch()
        eqg.setLayout(eql)
        top.addWidget(eqg)
        sl.addLayout(top)

        # 全局选项
        sg = QGroupBox("全局选项")
        sgl = QHBoxLayout()
        self.diff_first_cb = QCheckBox("首页不同")
        self.diff_odd_even_cb = QCheckBox("奇偶页不同")
        sgl.addWidget(self.diff_first_cb)
        sgl.addWidget(self.diff_odd_even_cb)
        sgl.addStretch()
        sg.setLayout(sgl)
        sl.addWidget(sg)

        # 三类页面的页眉页脚
        self._hf_widgets = {}
        for key, label in [
            ("cover", "第1类 — 封面/签署页/修改记录"),
            ("toc", "第2类 — 目录"),
            ("body", "第3类 — 正文"),
        ]:
            g = QGroupBox(label)
            gl = QGridLayout(g)

            gl.addWidget(QLabel("页眉文本:"), 0, 0)
            header_edit = QLineEdit()
            gl.addWidget(header_edit, 0, 1)
            show_header_cb = QCheckBox("显示页眉")
            gl.addWidget(show_header_cb, 1, 0)

            show_footer_cb = QCheckBox("显示页脚")
            gl.addWidget(show_footer_cb, 2, 0)

            wmap = {
                "header_text": header_edit,
                "show_header": show_header_cb,
                "show_footer": show_footer_cb,
            }

            if key == "cover":
                gl.addWidget(QLabel("页脚：无页码"), 3, 0, 1, 2)
            else:
                show_pn_cb = QCheckBox("显示页码")
                gl.addWidget(show_pn_cb, 3, 0)
                wmap["show_page_number"] = show_pn_cb

                gl.addWidget(QLabel("页码格式:"), 4, 0)
                pn_fmt = QComboBox()
                pn_fmt.addItems(list(PAGE_NUM_FORMAT_MAP.keys()))
                gl.addWidget(pn_fmt, 4, 1)
                wmap["pn_format"] = pn_fmt

                gl.addWidget(QLabel("页码位置:"), 5, 0)
                pn_pos = QComboBox()
                pn_pos.addItems(list(PAGE_NUM_POS_MAP.keys()))
                gl.addWidget(pn_pos, 5, 1)
                wmap["pn_pos"] = pn_pos

                gl.addWidget(QLabel("页码对齐:"), 6, 0)
                pn_align = QComboBox()
                pn_align.addItems(list(PAGE_NUM_ALIGN_MAP.keys()))
                gl.addWidget(pn_align, 6, 1)
                wmap["pn_align"] = pn_align

                pn_font = _FontParagraphGroup("页码字体与段落")
                gl.addWidget(pn_font, 7, 0, 1, 2)
                wmap["pn_font"] = pn_font

            sl.addWidget(g)
            self._hf_widgets[key] = wmap

        sl.addStretch()
        scroll.setWidget(w)
        layout.addWidget(scroll)

    def load_config(self, t: TemplateConfig):
        self.print_group.buttons()[0].setChecked(t.print_mode.value == "single")
        self.print_group.buttons()[1].setChecked(t.print_mode.value == "duplex")
        self.diff_first_cb.setChecked(t.header_footer.different_first_page)
        self.diff_odd_even_cb.setChecked(t.header_footer.different_odd_even)
        eq_map = {"keep": 0, "omml": 1, "image": 2}
        self.eq_format.setCurrentIndex(eq_map.get(t.equation_format.value, 0))

        for key in ["cover", "toc", "body"]:
            hp = getattr(t.header_footer, f"{key}_page")
            wm = self._hf_widgets[key]
            wm["header_text"].setText(hp.header_text)
            wm["show_header"].setChecked(hp.show_header)
            wm["show_footer"].setChecked(hp.show_footer)
            if key != "cover":
                wm["show_page_number"].setChecked(hp.show_page_number)
                wm["pn_format"].setCurrentText(PAGE_NUM_FORMAT_MAP_REV.get(hp.page_number_format.value, "大写罗马"))
                wm["pn_pos"].setCurrentText(PAGE_NUM_POS_MAP_REV.get(hp.page_number_position.value, "页脚"))
                wm["pn_align"].setCurrentText(PAGE_NUM_ALIGN_MAP_REV.get(hp.page_number_alignment.value, "居中"))
                wm["pn_font"].set_font_config(hp.page_number_font)
                wm["pn_font"].set_paragraph_config(hp.page_number_paragraph)

    def save_config(self, t: TemplateConfig):
        t.print_mode = PrintMode.SINGLE if self.print_group.checkedId() == 0 else PrintMode.DUPLEX
        t.header_footer.different_first_page = self.diff_first_cb.isChecked()
        t.header_footer.different_odd_even = self.diff_odd_even_cb.isChecked()
        eq_map = {0: "keep", 1: "omml", 2: "image"}
        t.equation_format = EquationFormat(eq_map[self.eq_format.currentIndex()])

        for key in ["cover", "toc", "body"]:
            hp = getattr(t.header_footer, f"{key}_page")
            wm = self._hf_widgets[key]
            hp.header_text = wm["header_text"].text()
            hp.show_header = wm["show_header"].isChecked()
            hp.show_footer = wm["show_footer"].isChecked()
            if key != "cover":
                hp.show_page_number = wm["show_page_number"].isChecked()
                hp.page_number_format = PageNumberFormat(PAGE_NUM_FORMAT_MAP.get(wm["pn_format"].currentText(), "roman_up"))
                hp.page_number_position = PageNumberPosition(PAGE_NUM_POS_MAP.get(wm["pn_pos"].currentText(), "footer"))
                hp.page_number_alignment = PageNumberAlignment(PAGE_NUM_ALIGN_MAP.get(wm["pn_align"].currentText(), "center"))
                hp.page_number_font = wm["pn_font"].get_font_config()
                hp.page_number_paragraph = wm["pn_font"].get_paragraph_config()

    def preview_html(self, scale: float = 1.0) -> str:
        lines = [f"<div style='font-size: {10 * scale:.0f}pt; line-height: 1.6;'>"]
        lines.append("<b>页眉页脚预览</b><br>")
        for key, label in [("cover", "封面/签署/修改"), ("toc", "目录"), ("body", "正文")]:
            wm = self._hf_widgets[key]
            hdr = wm["header_text"].text() or "(无)"
            lines.append(f"{label}: 页眉=\"{hdr}\" &nbsp;")
            lines.append(f"页眉{'显示' if wm['show_header'].isChecked() else '隐藏'} &nbsp;")
            lines.append(f"页脚{'显示' if wm['show_footer'].isChecked() else '隐藏'}<br>")
        lines.append("</div>")
        return "".join(lines)


class _CoverPage(QWidget):
    """封面 — 模板选择 + 占位符映射 + 字体段落"""

    def __init__(self):
        super().__init__()
        self.detected_title = ""
        layout = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        w = QWidget()
        sl = QVBoxLayout(w)

        self.enabled_cb = QCheckBox("启用封面")
        sl.addWidget(self.enabled_cb)

        tp_layout = QHBoxLayout()
        tp_layout.addWidget(QLabel("模板文件:"))
        self.tp_label = QLabel("未选择")
        btn = QPushButton("选择")
        btn.clicked.connect(self._select_template)
        tp_layout.addWidget(self.tp_label)
        tp_layout.addWidget(btn)
        sl.addWidget(QLabel("选择封面模板 .docx 后自动扫描占位符"))
        sl.addLayout(tp_layout)

        sl.addWidget(QLabel("字段映射（占位符 → 实际值）:"))
        self.field_table = QTableWidget(0, 2)
        self.field_table.setHorizontalHeaderLabels(["占位符", "值"])
        self.field_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        sl.addWidget(self.field_table)

        self.heading_fp = _FontParagraphGroup("标题字体与段落")
        sl.addWidget(self.heading_fp)
        sl.addStretch()

        scroll.setWidget(w)
        layout.addWidget(scroll)
        self._template_path = None

    def _select_template(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择封面模板", "", "Word文档 (*.docx)")
        if file:
            self._template_path = file
            self.tp_label.setText(Path(file).name)
            placeholders = scan_placeholders(file)
            self.field_table.setRowCount(len(placeholders))
            for i, ph in enumerate(placeholders):
                self.field_table.setItem(i, 0, QTableWidgetItem(ph))
                if self.field_table.item(i, 1) is None:
                    self.field_table.setItem(i, 1, QTableWidgetItem(""))
            # 自动检测标题：取第一个占位符作为标题候选
            if placeholders:
                self.detected_title = placeholders[0]
            else:
                self.detected_title = Path(file).stem

    def load_config(self, t: TemplateConfig):
        self.enabled_cb.setChecked(t.cover.enabled)
        self._template_path = t.cover.template_path
        self.tp_label.setText(Path(t.cover.template_path).name if t.cover.template_path else "未选择")
        keys = list(t.cover.fields.keys())
        self.field_table.setRowCount(len(keys))
        for i, k in enumerate(keys):
            self.field_table.setItem(i, 0, QTableWidgetItem(k))
            self.field_table.setItem(i, 1, QTableWidgetItem(t.cover.fields.get(k, "")))
        self.heading_fp.set_font_config(t.cover.heading_font)
        self.heading_fp.set_paragraph_config(t.cover.heading_paragraph)

    def save_config(self, t: TemplateConfig):
        t.cover.enabled = self.enabled_cb.isChecked()
        t.cover.template_path = self._template_path
        fields = {}
        for i in range(self.field_table.rowCount()):
            key_item = self.field_table.item(i, 0)
            val_item = self.field_table.item(i, 1)
            if key_item and key_item.text():
                fields[key_item.text()] = val_item.text() if val_item else ""
        t.cover.fields = fields
        t.cover.heading_font = self.heading_fp.get_font_config()
        t.cover.heading_paragraph = self.heading_fp.get_paragraph_config()

    def preview_html(self, scale: float = 1.0) -> str:
        f = self.heading_fp.get_font_config()
        weight = "bold" if f.bold else "normal"
        return (
            f"<div style='font-family: \"{f.cn_name}\"; font-size: {f.size * scale:.0f}pt;"
            f"font-weight: {weight}; color: {f.color};"
            f"text-align: center; padding: 40px 0;'>"
            f"{f.cn_name} {f.size_name}<br>封面标题预览"
            f"</div>"
        )


class _SignaturePage(QWidget):
    """签署页 — 模板 + 标题 + 字体段落"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        w = QWidget()
        sl = QVBoxLayout(w)

        self.enabled_cb = QCheckBox("启用签署页")
        sl.addWidget(self.enabled_cb)

        tp_layout = QHBoxLayout()
        tp_layout.addWidget(QLabel("模板文件:"))
        self.tp_label = QLabel("未选择")
        btn = QPushButton("选择")
        btn.clicked.connect(self._select_template)
        tp_layout.addWidget(self.tp_label)
        tp_layout.addWidget(btn)
        sl.addLayout(tp_layout)

        sl.addWidget(QLabel("标题文本（可从封面自动读取）:"))
        self.title_edit = QLineEdit("签署页")
        sl.addWidget(self.title_edit)

        self.heading_fp = _FontParagraphGroup("标题字体与段落")
        sl.addWidget(self.heading_fp)
        sl.addStretch()

        scroll.setWidget(w)
        layout.addWidget(scroll)
        self._template_path = None

    def _select_template(self):
        file, _ = QFileDialog.getOpenFileName(self, "选择签署页模板", "", "Word文档 (*.docx)")
        if file:
            self._template_path = file
            self.tp_label.setText(Path(file).name)

    def set_detected_title(self, title: str):
        """从封面自动检测的标题"""
        if title and self.title_edit.text() == "签署页":
            self.title_edit.setText(title)

    def load_config(self, t: TemplateConfig):
        self.enabled_cb.setChecked(t.signature.enabled)
        self._template_path = t.signature.template_path
        self.tp_label.setText(Path(t.signature.template_path).name if t.signature.template_path else "未选择")
        self.title_edit.setText(t.signature.title)
        self.heading_fp.set_font_config(t.signature.heading_font)
        self.heading_fp.set_paragraph_config(t.signature.heading_paragraph)

    def save_config(self, t: TemplateConfig):
        t.signature.enabled = self.enabled_cb.isChecked()
        t.signature.template_path = self._template_path
        t.signature.title = self.title_edit.text()
        t.signature.heading_font = self.heading_fp.get_font_config()
        t.signature.heading_paragraph = self.heading_fp.get_paragraph_config()

    def preview_html(self, scale: float = 1.0) -> str:
        f = self.heading_fp.get_font_config()
        title = self.title_edit.text() or "签署页"
        weight = "bold" if f.bold else "normal"
        return (
            f"<div style='font-family: \"{f.cn_name}\"; font-size: {f.size * scale:.0f}pt;"
            f"font-weight: {weight}; color: {f.color};"
            f"text-align: center; padding: 30px 0;'>"
            f"{title}"
            f"</div>"
        )


class _RevisionPage(QWidget):
    """修改记录页 — 标题 + 表头 + 正文"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        w = QWidget()
        sl = QVBoxLayout(w)

        self.enabled_cb = QCheckBox("启用修改记录页")
        sl.addWidget(self.enabled_cb)

        self.heading_fp = _FontParagraphGroup("标题字体与段落")
        sl.addWidget(self.heading_fp)
        self.th_fp = _FontParagraphGroup("表格表头字体与段落")
        sl.addWidget(self.th_fp)
        self.tb_fp = _FontParagraphGroup("表格正文字体与段落")
        sl.addWidget(self.tb_fp)
        sl.addStretch()

        scroll.setWidget(w)
        layout.addWidget(scroll)

    def load_config(self, t: TemplateConfig):
        self.enabled_cb.setChecked(t.revision.enabled)
        self.heading_fp.set_font_config(t.revision.heading_font)
        self.heading_fp.set_paragraph_config(t.revision.heading_paragraph)
        self.th_fp.set_font_config(t.revision.table_header_font)
        self.th_fp.set_paragraph_config(t.revision.table_header_paragraph)
        self.tb_fp.set_font_config(t.revision.table_body_font)
        self.tb_fp.set_paragraph_config(t.revision.table_body_paragraph)

    def save_config(self, t: TemplateConfig):
        t.revision.enabled = self.enabled_cb.isChecked()
        t.revision.heading_font = self.heading_fp.get_font_config()
        t.revision.heading_paragraph = self.heading_fp.get_paragraph_config()
        t.revision.table_header_font = self.th_fp.get_font_config()
        t.revision.table_header_paragraph = self.th_fp.get_paragraph_config()
        t.revision.table_body_font = self.tb_fp.get_font_config()
        t.revision.table_body_paragraph = self.tb_fp.get_paragraph_config()

    def preview_html(self, scale: float = 1.0) -> str:
        hf = self.heading_fp.get_font_config()
        weight = "bold" if hf.bold else "normal"
        return (
            f"<div style='font-family: \"{hf.cn_name}\"; font-size: {hf.size * scale:.0f}pt;"
            f"font-weight: {weight}; text-align: center;'>"
            f"修改记录<br>"
            f"<table style='width:100%; border-collapse: collapse; font-size: {10 * scale:.0f}pt; margin-top: 10px;'>"
            f"<tr style='background: #E8E8E8;'><th style='border:1px solid #ccc;padding:4px;'>版本</th>"
            f"<th style='border:1px solid #ccc;padding:4px;'>日期</th><th style='border:1px solid #ccc;padding:4px;'>说明</th></tr>"
            f"<tr><td style='border:1px solid #ccc;padding:4px;'>V1.0</td>"
            f"<td style='border:1px solid #ccc;padding:4px;'>2026-01-01</td>"
            f"<td style='border:1px solid #ccc;padding:4px;'>初版</td></tr>"
            f"</table></div>"
        )


class _TOCPage(QWidget):
    """目录页"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        w = QWidget()
        sl = QVBoxLayout(w)

        self.enabled_cb = QCheckBox("启用目录页")
        sl.addWidget(self.enabled_cb)

        self.heading_fp = _FontParagraphGroup("标题字体与段落")
        sl.addWidget(self.heading_fp)
        self.entry_fp = _FontParagraphGroup("条目字体与段落")
        sl.addWidget(self.entry_fp)

        indent_layout = QHBoxLayout()
        indent_layout.addWidget(QLabel("每级缩进(字符):"))
        self.entry_indent = QDoubleSpinBox()
        self.entry_indent.setRange(0, 5)
        indent_layout.addWidget(self.entry_indent)
        indent_layout.addStretch()
        sl.addLayout(indent_layout)
        sl.addStretch()

        scroll.setWidget(w)
        layout.addWidget(scroll)

    def load_config(self, t: TemplateConfig):
        self.enabled_cb.setChecked(t.toc.enabled)
        self.heading_fp.set_font_config(t.toc.heading_font)
        self.heading_fp.set_paragraph_config(t.toc.heading_paragraph)
        self.entry_fp.set_font_config(t.toc.entry_font)
        self.entry_fp.set_paragraph_config(t.toc.entry_paragraph)
        self.entry_indent.setValue(t.toc.entry_indent)

    def save_config(self, t: TemplateConfig):
        t.toc.enabled = self.enabled_cb.isChecked()
        t.toc.heading_font = self.heading_fp.get_font_config()
        t.toc.heading_paragraph = self.heading_fp.get_paragraph_config()
        t.toc.entry_font = self.entry_fp.get_font_config()
        t.toc.entry_paragraph = self.entry_fp.get_paragraph_config()
        t.toc.entry_indent = self.entry_indent.value()

    def preview_html(self, scale: float = 1.0) -> str:
        hf = self.heading_fp.get_font_config()
        ef = self.entry_fp.get_font_config()
        return (
            f"<div>"
            f"<div style='font-family: \"{hf.cn_name}\"; font-size: {hf.size * scale:.0f}pt;"
            f"font-weight: bold; text-align: center;'>目 录</div>"
            f"<div style='font-family: \"{ef.cn_name}\"; font-size: {ef.size * scale:.0f}pt; line-height: 1.8;'>"
            f"<div>&nbsp;&nbsp;1. 第一章 概述</div>"
            f"<div style='margin-left: {self.entry_indent.value() * 16}px;'>1.1 背景</div>"
            f"<div style='margin-left: {self.entry_indent.value() * 16}px;'>1.2 范围</div>"
            f"<div>&nbsp;&nbsp;2. 第二章 方案</div>"
            f"</div></div>"
        )


class _HeadingTablePage(QWidget):
    """多级列表配置 — 表格展示所有级别 + 编号预览 + 选中行编辑字体段落"""

    changed = pyqtSignal()

    # 编号预览辅助
    _STYLE_SEP = {"none": "", "chinese": "、", "decimal": ".", "upper_letter": ".",
                  "lower_letter": ".", "upper_roman": ".", "lower_roman": "."}
    _STYLE_SAMPLE = {"none": "", "chinese": "一", "decimal": "1", "upper_letter": "A",
                     "lower_letter": "a", "upper_roman": "I", "lower_roman": "i"}

    def __init__(self):
        super().__init__()
        self._current_row = -1
        layout = QVBoxLayout(self)

        # 操作按钮
        btn_layout = QHBoxLayout()
        self.add_btn = QPushButton("＋ 增加级别")
        self.add_btn.clicked.connect(self._add_level)
        btn_layout.addWidget(self.add_btn)
        self.del_btn = QPushButton("－ 删除最后级别")
        self.del_btn.clicked.connect(self._del_level)
        btn_layout.addWidget(self.del_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # 多级列表表格: 级别 | 编号格式 | 包含上级 | 预览
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["级别", "编号格式", "包含上级", "编号预览"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.itemClicked.connect(self._on_row_clicked)
        layout.addWidget(self.table, 1)

        # 选中级别的字体段落编辑
        self.selected_fp = _FontParagraphGroup("选中级别 — 字体与段落")
        self.selected_fp.changed.connect(self._on_fp_changed)
        layout.addWidget(self.selected_fp)

        self._headings = []

    # ── 编号预览算法 ──

    def _build_number_preview(self, i: int) -> str:
        """构建第 i 级的多级编号预览，如 '一、1.' 或 '1.1.'"""
        h = self._headings[i]
        if h.number_style == "none":
            return "(无)"
        parts = []
        for j in range(i, -1, -1):
            hj = self._headings[j]
            if hj.number_style == "none":
                continue
            sample = self._STYLE_SAMPLE.get(hj.number_style, "")
            sep = self._STYLE_SEP.get(hj.number_style, "")
            parts.insert(0, sample + sep)
            # j 的某个下级（j+1 到 i）multi=False 则断链，不再往上追溯
            if j < i and any(not self._headings[k].number_multi for k in range(j + 1, i + 1)):
                parts.pop(0)
                break
        return "".join(parts)

    def _refresh_preview_column(self, from_row: int = 0):
        """刷新 from_row 及之后所有行的编号预览列"""
        for i in range(from_row, len(self._headings)):
            self.table.setItem(i, 3, QTableWidgetItem(self._build_number_preview(i)))

    # ── 行选择 ──

    def _on_row_clicked(self, item):
        self._select_row(item.row())

    def _select_row(self, row):
        if row == self._current_row:
            return
        self._save_current_row()
        self._current_row = row
        self.table.selectRow(row)
        if 0 <= row < len(self._headings):
            self.selected_fp.blockSignals(True)
            self.selected_fp.set_font_config(self._headings[row].font)
            self.selected_fp.set_paragraph_config(self._headings[row].paragraph)
            self.selected_fp.blockSignals(False)

    def _save_current_row(self):
        if 0 <= self._current_row < len(self._headings):
            h = self._headings[self._current_row]
            h.font = self.selected_fp.get_font_config()
            h.paragraph = self.selected_fp.get_paragraph_config()

    def _on_fp_changed(self):
        """字体段落组变化 → 保存到当前行数据"""
        if 0 <= self._current_row < len(self._headings):
            h = self._headings[self._current_row]
            h.font = self.selected_fp.get_font_config()
            h.paragraph = self.selected_fp.get_paragraph_config()

    # ── 增删级别 ──

    def _add_level(self):
        from ..models.template_config import HeadingConfig
        level = len(self._headings) + 1
        h = HeadingConfig(
            level=level, number_style="none", number_multi=False,
            font=FontConfig(cn_name="宋体", size=10.5, size_name="五号"),
            paragraph=ParagraphConfig(line_spacing=1.5),
        )
        self._headings.append(h)
        self._refresh_table()
        self.table.selectRow(level - 1)
        self._current_row = level - 1
        self.selected_fp.blockSignals(True)
        self.selected_fp.set_font_config(h.font)
        self.selected_fp.set_paragraph_config(h.paragraph)
        self.selected_fp.blockSignals(False)

    def _del_level(self):
        if not self._headings:
            return
        self._save_current_row()
        self._headings.pop()
        if self._current_row >= len(self._headings):
            self._current_row = len(self._headings) - 1
        self._refresh_table()
        if self._current_row >= 0:
            self.table.selectRow(self._current_row)
            self.selected_fp.blockSignals(True)
            self.selected_fp.set_font_config(self._headings[self._current_row].font)
            self.selected_fp.set_paragraph_config(self._headings[self._current_row].paragraph)
            self.selected_fp.blockSignals(False)

    # ── 表格构建 ──

    def _refresh_table(self):
        self.table.setRowCount(len(self._headings))
        for i in range(len(self._headings)):
            self._refresh_table_row(i)
        self._refresh_preview_column()

    def _refresh_table_row(self, i):
        h = self._headings[i]
        self.table.setItem(i, 0, QTableWidgetItem(str(h.level)))

        # 编号格式 — QComboBox（列1）
        style_combo = QComboBox()
        style_combo.addItems(NUMBER_STYLE_KEYS)
        style_combo.currentTextChanged.connect(lambda txt, r=i: self._on_style_changed(r, txt))
        style_combo.blockSignals(True)
        style_combo.setCurrentText(NUMBER_STYLE_MAP_REV.get(h.number_style, "无编号"))
        style_combo.blockSignals(False)
        self.table.setCellWidget(i, 1, style_combo)

        # 包含上级 — QCheckBox
        multi_cb = QCheckBox("包含上级")
        multi_cb.setToolTip("勾选后编号会带上前面级别的序号，如 1.1、一、(一)")
        multi_cb.toggled.connect(lambda checked, r=i: self._on_multi_changed(r, checked))
        multi_cb.blockSignals(True)
        multi_cb.setChecked(h.number_multi)
        multi_cb.blockSignals(False)
        self.table.setCellWidget(i, 2, multi_cb)

    # ── 交互回调 ──

    def _on_style_changed(self, row, text):
        if 0 <= row < len(self._headings):
            self._headings[row].number_style = NUMBER_STYLE_MAP.get(text, "none")
            self._refresh_preview_column(row)
            self._select_row(row)
            self.changed.emit()

    def _on_multi_changed(self, row, checked):
        if 0 <= row < len(self._headings):
            self._headings[row].number_multi = checked
            self._refresh_preview_column(row)
            self._select_row(row)
            self.changed.emit()

    # ── 模板 I/O ──

    def load_config(self, t: TemplateConfig):
        self._headings = t.headings[:]
        self._current_row = -1
        self._refresh_table()
        if self._headings:
            self.table.selectRow(0)
            self._current_row = 0
            self.selected_fp.blockSignals(True)
            self.selected_fp.set_font_config(self._headings[0].font)
            self.selected_fp.set_paragraph_config(self._headings[0].paragraph)
            self.selected_fp.blockSignals(False)

    def save_config(self, t: TemplateConfig):
        self._save_current_row()
        t.headings = self._headings[:]

    # ── 预览 ──

    def preview_html(self, scale: float = 1.0) -> str:
        parts = ["<div style='line-height: 1.8;'>"]
        for h in self._headings:
            f = h.font
            weight = "bold" if f.bold else "normal"
            num = self._build_number_preview(h.level - 1)
            prefix = f"{num} " if num and num != "(无)" else ""
            parts.append(
                f"<div style='font-family: \"{f.cn_name}\"; font-size: {f.size * scale:.0f}pt;"
                f"font-weight: {weight}; color: {f.color};"
                f"padding: 4px 0;'>"
                f"{prefix}标题{h.level} 示例文本"
                f"</div>"
            )
        parts.append("</div>")
        return "".join(parts)


class _BodyPage(QWidget):
    """正文配置"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        self.fp = _FontParagraphGroup("正文字体与段落")
        layout.addWidget(self.fp)
        layout.addStretch()

    def load_config(self, t: TemplateConfig):
        self.fp.set_font_config(t.body.font)
        self.fp.set_paragraph_config(t.body.paragraph)

    def save_config(self, t: TemplateConfig):
        t.body.font = self.fp.get_font_config()
        t.body.paragraph = self.fp.get_paragraph_config()

    def preview_html(self, scale: float = 1.0) -> str:
        return self.fp.preview_html()


class _TablePage(QWidget):
    """表格配置 — 基础表/记录表切换 + 表头/正文"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
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

        self.header_fp = _FontParagraphGroup("表头字体与段落")
        layout.addWidget(self.header_fp)

        bg_layout = QHBoxLayout()
        bg_layout.addWidget(QLabel("表头背景色:"))
        self.bg_btn = QPushButton("选择")
        self.bg_btn.clicked.connect(self._pick_bg)
        bg_layout.addWidget(self.bg_btn)
        bg_layout.addStretch()
        layout.addLayout(bg_layout)

        self.body_fp = _FontParagraphGroup("正文字体与段落")
        layout.addWidget(self.body_fp)
        layout.addStretch()
        self._current_bg = "#E8E8E8"

    def _pick_bg(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self._current_bg = color.name()
            self.bg_btn.setStyleSheet(f"background-color: {self._current_bg};")

    def load_config(self, t: TemplateConfig):
        idx = self.table_type_group.checkedId()
        ts = t.table_font.regular_table if idx == 0 else t.table_font.record_table
        self.header_fp.set_font_config(ts.header_font)
        self.header_fp.set_paragraph_config(ts.header_paragraph)
        self._current_bg = ts.header_bg_color
        self.bg_btn.setStyleSheet(f"background-color: {self._current_bg};")
        self.body_fp.set_font_config(ts.body_font)
        self.body_fp.set_paragraph_config(ts.body_paragraph)

    def save_config(self, t: TemplateConfig):
        idx = self.table_type_group.checkedId()
        ts = t.table_font.regular_table if idx == 0 else t.table_font.record_table
        ts.header_font = self.header_fp.get_font_config()
        ts.header_paragraph = self.header_fp.get_paragraph_config()
        ts.header_bg_color = self._current_bg
        ts.body_font = self.body_fp.get_font_config()
        ts.body_paragraph = self.body_fp.get_paragraph_config()

    def preview_html(self, scale: float = 1.0) -> str:
        hf = self.header_fp.get_font_config()
        bf = self.body_fp.get_font_config()
        return (
            f"<table style='width:100%; border-collapse: collapse;'>"
            f"<tr style='background: {self._current_bg};'>"
            f"<th style='border:1px solid #aaa;padding:4px;font-family:\"{hf.cn_name}\";"
            f"font-size:{hf.size * scale:.0f}pt;font-weight:bold;'>列A</th>"
            f"<th style='border:1px solid #aaa;padding:4px;font-family:\"{hf.cn_name}\";"
            f"font-size:{hf.size * scale:.0f}pt;font-weight:bold;'>列B</th></tr>"
            f"<tr><td style='border:1px solid #aaa;padding:4px;font-family:\"{bf.cn_name}\";"
            f"font-size:{bf.size * scale:.0f}pt;'>值1</td>"
            f"<td style='border:1px solid #aaa;padding:4px;font-family:\"{bf.cn_name}\";"
            f"font-size:{bf.size * scale:.0f}pt;'>值2</td></tr>"
            f"</table>"
        )


class _CaptionPage(QWidget):
    """题注配置"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        w = QWidget()
        sl = QVBoxLayout(w)

        # 图序
        g1 = QGroupBox("图序")
        g1l = QGridLayout(g1)
        g1l.addWidget(QLabel("前缀:"), 0, 0)
        self.fig_prefix = QLineEdit("图")
        g1l.addWidget(self.fig_prefix, 0, 1)
        g1l.addWidget(QLabel("位置:"), 0, 2)
        self.fig_pos = QComboBox()
        self.fig_pos.addItems(["下方", "上方"])
        g1l.addWidget(self.fig_pos, 0, 3)
        sl.addWidget(g1)

        # 表序
        g2 = QGroupBox("表序")
        g2l = QGridLayout(g2)
        g2l.addWidget(QLabel("前缀:"), 0, 0)
        self.tbl_prefix = QLineEdit("表")
        g2l.addWidget(self.tbl_prefix, 0, 1)
        g2l.addWidget(QLabel("位置:"), 0, 2)
        self.tbl_pos = QComboBox()
        self.tbl_pos.addItems(["上方", "下方"])
        g2l.addWidget(self.tbl_pos, 0, 3)
        sl.addWidget(g2)

        # 公式序
        g3 = QGroupBox("公式序")
        g3l = QGridLayout(g3)
        g3l.addWidget(QLabel("前缀:"), 0, 0)
        self.eq_prefix = QLineEdit("公式")
        g3l.addWidget(self.eq_prefix, 0, 1)
        g3l.addWidget(QLabel("编号模式:"), 0, 2)
        self.num_mode = QComboBox()
        self.num_mode.addItems(["全局编号", "按章节编号"])
        g3l.addWidget(self.num_mode, 0, 3)
        sl.addWidget(g3)

        self.cap_fp = _FontParagraphGroup("题注字体与段落")
        sl.addWidget(self.cap_fp)
        sl.addStretch()

        scroll.setWidget(w)
        layout.addWidget(scroll)

    def load_config(self, t: TemplateConfig):
        self.fig_prefix.setText(t.caption.figure_prefix)
        self.fig_pos.setCurrentText("下方" if t.caption.position_figure == "below" else "上方")
        self.tbl_prefix.setText(t.caption.table_prefix)
        self.tbl_pos.setCurrentText("上方" if t.caption.position_table == "above" else "下方")
        self.eq_prefix.setText(t.caption.equation_prefix)
        self.num_mode.setCurrentText("全局编号" if t.caption.numbering_mode == NumberingMode.GLOBAL else "按章节编号")
        self.cap_fp.set_font_config(t.caption.font)
        self.cap_fp.set_paragraph_config(t.caption.paragraph)

    def save_config(self, t: TemplateConfig):
        t.caption.figure_prefix = self.fig_prefix.text()
        t.caption.position_figure = "below" if self.fig_pos.currentText() == "下方" else "above"
        t.caption.table_prefix = self.tbl_prefix.text()
        t.caption.position_table = "above" if self.tbl_pos.currentText() == "上方" else "below"
        t.caption.equation_prefix = self.eq_prefix.text()
        t.caption.numbering_mode = NumberingMode.GLOBAL if self.num_mode.currentText() == "全局编号" else NumberingMode.PER_TYPE
        t.caption.font = self.cap_fp.get_font_config()
        t.caption.paragraph = self.cap_fp.get_paragraph_config()

    def preview_html(self, scale: float = 1.0) -> str:
        f = self.cap_fp.get_font_config()
        return (
            f"<div style='font-family: \"{f.cn_name}\"; font-size: {f.size * scale:.0f}pt;"
            f"text-align: center; padding: 8px;'>"
            f"{self.fig_prefix.text()} 1 示例题注 &nbsp;|&nbsp; "
            f"{self.tbl_prefix.text()} 1 示例表注 &nbsp;|&nbsp; "
            f"{self.eq_prefix.text()} (1)"
            f"</div>"
        )


class _CodeBlockPage(QWidget):
    """代码块配置"""

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        self.fp = _FontParagraphGroup("代码字体与段落")
        layout.addWidget(self.fp)

        bg_layout = QHBoxLayout()
        bg_layout.addWidget(QLabel("背景色:"))
        self.bg_btn = QPushButton("选择")
        self.bg_btn.clicked.connect(self._pick_bg)
        bg_layout.addWidget(self.bg_btn)
        bg_layout.addStretch()
        layout.addLayout(bg_layout)
        layout.addStretch()
        self._current_bg = "#F0F0F0"

    def _pick_bg(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self._current_bg = color.name()
            self.bg_btn.setStyleSheet(f"background-color: {self._current_bg};")

    def load_config(self, t: TemplateConfig):
        self.fp.set_font_config(t.code_block.font)
        self.fp.set_paragraph_config(t.code_block.paragraph)
        self._current_bg = t.code_block.bg_color
        self.bg_btn.setStyleSheet(f"background-color: {self._current_bg};")

    def save_config(self, t: TemplateConfig):
        t.code_block.font = self.fp.get_font_config()
        t.code_block.paragraph = self.fp.get_paragraph_config()
        t.code_block.bg_color = self._current_bg

    def preview_html(self, scale: float = 1.0) -> str:
        f = self.fp.get_font_config()
        return (
            f"<pre style='font-family: \"{f.en_name}\", monospace; font-size: {f.size * scale:.0f}pt;"
            f"background: {self._current_bg}; padding: 12px; border-radius: 4px;'>"
            f"def hello():<br>    print(\"Hello World\")"
            f"</pre>"
        )


# ══════════════════════════════════════════════════════════════
# 预览面板
# ══════════════════════════════════════════════════════════════

class _PreviewPanel(QFrame):
    """独立预览面板 — 根据页面设置动态调整比例和字体缩放"""

    PREVIEW_WIDTH = 200

    def __init__(self):
        super().__init__()
        self.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.setMaximumWidth(self.PREVIEW_WIDTH + 40)
        layout = QVBoxLayout(self)

        self._title = QLabel("<b>风格预览</b> (A4)")
        self._title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self._title)

        self.preview_area = QTextEdit()
        self.preview_area.setReadOnly(True)
        self.preview_area.setFixedWidth(self.PREVIEW_WIDTH)
        self.preview_area.setStyleSheet(
            "QTextEdit { background: white; border: 1px solid #aaa; padding: 12px; }"
        )
        layout.addWidget(self.preview_area, 0, Qt.AlignmentFlag.AlignCenter)

        hint = QLabel("<i>HTML 渲染，非 Word 所见即所得</i>")
        hint.setWordWrap(True)
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(hint)
        layout.addStretch()

        self._scale = 1.0
        self.set_page_size(210, 297)  # 默认 A4

    def set_page_size(self, width_mm: float, height_mm: float):
        """根据页面尺寸调整预览区域宽高比"""
        is_landscape = width_mm > height_mm
        pw = float(width_mm) if not is_landscape else float(height_mm)
        ph = float(height_mm) if not is_landscape else float(width_mm)
        ratio = ph / pw
        panel_h = max(60, int(self.PREVIEW_WIDTH * ratio))
        self.preview_area.setFixedHeight(panel_h)

        paper_names = {210: "A4", 297: "A3", 148: "A5", 176: "B5", 215.9: "Letter", 215.9: "Legal"}
        w_key = round(width_mm, 1)
        name = paper_names.get(w_key, f"{width_mm:.0f}×{height_mm:.0f}")
        orient = "横向" if is_landscape else "纵向"
        self._title.setText(f"<b>风格预览</b> ({name} {orient})")

        self._scale = self.PREVIEW_WIDTH / max(width_mm, height_mm) if is_landscape else self.PREVIEW_WIDTH / width_mm

    def show_preview(self, html: str):
        self.preview_area.setHtml(html)

    @property
    def scale(self) -> float:
        return self._scale


# ══════════════════════════════════════════════════════════════
# 主模板配置控件
# ══════════════════════════════════════════════════════════════

class TemplateConfigWidget(QWidget):
    """模板配置 Tab — 顶部工具栏 + 左侧列表 + 中间配置 + 右侧预览"""

    # 左侧列表项定义：(id, 显示名, 页面类)
    PAGE_DEFS = [
        ("page", "页面设置", _PageConfigPage),
        ("header_footer", "页眉页脚", _HeaderFooterPage),
        ("cover", "封面", _CoverPage),
        ("signature", "签署页", _SignaturePage),
        ("revision", "修改记录页", _RevisionPage),
        ("toc", "目录页", _TOCPage),
        ("heading", "标题", _HeadingTablePage),
        ("body", "正文", _BodyPage),
        ("table", "表格", _TablePage),
        ("caption", "题注", _CaptionPage),
        ("code_block", "代码块", _CodeBlockPage),
    ]

    def __init__(self):
        super().__init__()
        self.template = create_default_template()
        self.template_path = None
        self.tm = TemplateManager()
        self._current_item_id = None
        self._pages = {}
        self._init_ui()
        self._connect_signals()
        self._refresh_template_combo()

    def _init_ui(self):
        main_layout = QVBoxLayout(self)

        # ── 顶部工具栏 ──
        toolbar = QHBoxLayout()
        self.button_new = QPushButton("新建模板")
        self.button_new.setMinimumHeight(32)
        toolbar.addWidget(self.button_new)

        toolbar.addWidget(QLabel(" 模板:"))
        self.template_combo = QComboBox()
        self.template_combo.setMinimumWidth(180)
        self.template_combo.currentIndexChanged.connect(self._on_template_select)
        toolbar.addWidget(self.template_combo)

        self.button_save = QPushButton("保存")
        self.button_save.setMinimumHeight(32)
        toolbar.addWidget(self.button_save)

        self.button_save_as = QPushButton("另存为")
        self.button_save_as.setMinimumHeight(32)
        toolbar.addWidget(self.button_save_as)

        self.button_export = QPushButton("导出")
        self.button_export.setMinimumHeight(32)
        toolbar.addWidget(self.button_export)

        self.button_import = QPushButton("导入")
        self.button_import.setMinimumHeight(32)
        toolbar.addWidget(self.button_import)

        toolbar.addStretch()
        main_layout.addLayout(toolbar)

        # ── 元信息 ──
        meta_bar = QHBoxLayout()
        meta_bar.addWidget(QLabel("名称:"))
        self.meta_name = QLineEdit()
        self.meta_name.setPlaceholderText("模板名称")
        self.meta_name.setMaximumWidth(150)
        meta_bar.addWidget(self.meta_name)

        meta_bar.addWidget(QLabel("标签:"))
        self.meta_tags = QLineEdit()
        self.meta_tags.setPlaceholderText("逗号分隔")
        self.meta_tags.setMaximumWidth(150)
        meta_bar.addWidget(self.meta_tags)

        meta_bar.addWidget(QLabel("来源文档:"))
        self.meta_docs = QLineEdit()
        self.meta_docs.setPlaceholderText("逗号分隔")
        self.meta_docs.setMaximumWidth(180)
        meta_bar.addWidget(self.meta_docs)

        meta_bar.addStretch()
        main_layout.addLayout(meta_bar)

        # ── 左 / 中 / 右 三栏 ──
        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_panel = self._create_left_panel()
        splitter.addWidget(left_panel)

        self.pages_stack = QStackedWidget()
        for item_id, _, page_cls in self.PAGE_DEFS:
            page = page_cls()
            self._pages[item_id] = page
            self.pages_stack.addWidget(page)
        splitter.addWidget(self.pages_stack)

        self.preview = _PreviewPanel()
        splitter.addWidget(self.preview)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setStretchFactor(2, 1)
        main_layout.addWidget(splitter, 1)

    def _create_left_panel(self) -> QWidget:
        widget = QFrame()
        widget.setFrameStyle(QFrame.Shape.StyledPanel)
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel("<b>配置项目</b>"))
        self.config_list = QListWidget()

        for item_id, item_name, _ in self.PAGE_DEFS:
            list_item = QListWidgetItem(item_name)
            list_item.setData(Qt.ItemDataRole.UserRole, item_id)
            self.config_list.addItem(list_item)

        self.config_list.setCurrentRow(0)
        layout.addWidget(self.config_list)
        return widget

    def _connect_signals(self):
        self.config_list.currentRowChanged.connect(self._on_item_changed)
        self.button_new.clicked.connect(self.new_template)
        self.button_save.clicked.connect(self.save_template)
        self.button_save_as.clicked.connect(self._save_as)
        self.button_export.clicked.connect(self._export_template)
        self.button_import.clicked.connect(self._import_template)
        # 手动触发初始页加载
        initial_row = self.config_list.currentRow()
        if initial_row >= 0:
            self._on_item_changed(initial_row)

    def _on_item_changed(self, row):
        # 保存当前页面数据
        if self._current_item_id and self._current_item_id in self._pages:
            self._pages[self._current_item_id].save_config(self.template)

        item = self.config_list.item(row)
        if item is None:
            return
        item_id = item.data(Qt.ItemDataRole.UserRole)
        if item_id not in self._pages:
            return

        self._current_item_id = item_id
        page = self._pages[item_id]
        self.pages_stack.setCurrentWidget(page)
        page.load_config(self.template)

        # 封面→签署页标题联动
        if item_id == "signature" and "cover" in self._pages:
            cover_title = self._pages["cover"].detected_title
            if cover_title:
                self._pages["signature"].set_detected_title(cover_title)

        # 重新绑定预览更新信号
        self._reconnect_preview(page)

        # 更新预览
        self._update_preview()

    def _reconnect_preview(self, page):
        """断开旧预览信号，连接当前页面的控件变更信号"""
        # 断开旧连接
        if hasattr(self, '_preview_conns'):
            for signal, slot in self._preview_conns:
                try:
                    signal.disconnect(slot)
                except TypeError:
                    pass
        self._preview_conns = []

        # 连接所有 _FontParagraphGroup 的 changed 信号
        for child in page.findChildren(_FontParagraphGroup):
            child.changed.connect(self._update_preview)
            self._preview_conns.append((child.changed, self._update_preview))

        # 连接页面特有控件
        page_id = self._current_item_id
        if page_id == "page":
            for w in [page.preset_combo, page.page_width, page.page_height,
                       page.orientation_cb, page.margin_top, page.margin_bottom,
                       page.margin_left, page.margin_right, page.gutter,
                       page.gutter_pos, page.page_num_start]:
                self._connect_widget(w)
        elif page_id == "header_footer":
            for rb in page.print_group.buttons():
                self._connect_widget(rb)
            self._connect_widget(page.eq_format)
            self._connect_widget(page.diff_first_cb)
            self._connect_widget(page.diff_odd_even_cb)
        elif page_id == "cover":
            self._connect_widget(page.enabled_cb)
        elif page_id == "heading":
            page.table.itemSelectionChanged.connect(self._update_preview)
            self._preview_conns.append((page.table.itemSelectionChanged, self._update_preview))
            page.changed.connect(self._update_preview)
            self._preview_conns.append((page.changed, self._update_preview))
        elif page_id == "table":
            for rb in page.table_type_group.buttons():
                self._connect_widget(rb)
            self._connect_widget(page.bg_btn)
        elif page_id == "caption":
            for w in [page.fig_prefix, page.fig_pos, page.tbl_prefix,
                       page.tbl_pos, page.eq_prefix, page.num_mode]:
                self._connect_widget(w)
        elif page_id == "code_block":
            page.bg_btn.clicked.connect(self._update_preview)
            self._preview_conns.append((page.bg_btn.clicked, self._update_preview))

    def _connect_widget(self, w):
        """连接通用控件的变更信号到预览更新"""
        sig = None
        if hasattr(w, 'currentTextChanged'):
            sig = w.currentTextChanged
        elif hasattr(w, 'valueChanged'):
            sig = w.valueChanged
        elif hasattr(w, 'toggled'):
            sig = w.toggled
        elif hasattr(w, 'clicked'):
            sig = w.clicked
        elif hasattr(w, 'textChanged'):
            sig = w.textChanged
        if sig is not None:
            sig.connect(self._update_preview)
            self._preview_conns.append((sig, self._update_preview))

    def _update_preview(self, *_):
        """更新预览面板"""
        if self._current_item_id and self._current_item_id in self._pages:
            page = self._pages[self._current_item_id]
            if hasattr(page, 'preview_html'):
                page.save_config(self.template)
                self.preview.set_page_size(self.template.page.width, self.template.page.height)
                html = page.preview_html(scale=self.preview.scale)
                self.preview.show_preview(html)

    def _load_all_pages(self):
        for page in self._pages.values():
            page.load_config(self.template)
        # 刷新当前页面预览
        if self._current_item_id and self._current_item_id in self._pages:
            page = self._pages[self._current_item_id]
            self._reconnect_preview(page)
            self._update_preview()

    # ── 模板管理 ──

    def _refresh_template_combo(self):
        """刷新模板下拉列表"""
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        self._templates = self.tm.list_templates()
        matched = -1
        for i, t in enumerate(self._templates):
            suffix = " [内置]" if t["source"] == "builtin" else ""
            self.template_combo.addItem(f"{t['name']}{suffix}")
            if self.template_path and t["path"] == self.template_path:
                matched = i
        self.template_combo.blockSignals(False)
        if matched >= 0:
            self.template_combo.setCurrentIndex(matched)
        elif self._templates and self.template_path is None:
            self.template_combo.setCurrentIndex(0)
            self._on_template_select(0)

    def _on_template_select(self, index):
        if index < 0 or index >= len(self._templates):
            return
        t = self._templates[index]
        try:
            self.template = self.tm.load_template(t["path"])
            self.template_path = t["path"]
            self._load_all_pages()
            self._load_meta_into_ui(t["meta"])
        except Exception as e:
            QMessageBox.warning(self, "失败", f"加载模板失败:\n{str(e)}")

    def _load_meta_into_ui(self, meta: dict):
        self.meta_name.setText(meta.get("name", ""))
        self.meta_tags.setText(", ".join(meta.get("tags", [])))
        self.meta_docs.setText(", ".join(meta.get("source_docs", [])))

    def _collect_meta(self) -> dict:
        from ..templates.template_manager import read_meta
        meta = {}
        if self.template_path:
            try:
                meta = read_meta(self.template_path)
            except Exception:
                pass
        meta["name"] = self.meta_name.text() or "未命名模板"
        meta["tags"] = [t.strip() for t in self.meta_tags.text().split(",") if t.strip()]
        meta["source_docs"] = [d.strip() for d in self.meta_docs.text().split(",") if d.strip()]
        return meta

    def new_template(self):
        self.template = create_default_template()
        self.template_path = None
        self._templates = []
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        self.template_combo.blockSignals(False)
        self.meta_name.clear()
        self.meta_tags.clear()
        self.meta_docs.clear()
        self._load_all_pages()

    def save_template(self):
        if self._current_item_id and self._current_item_id in self._pages:
            self._pages[self._current_item_id].save_config(self.template)

        if self.template_path and Path(self.template_path).parent == self.tm.user_dir:
            try:
                self.tm.save(self.template, self.template_path)
                QMessageBox.information(self, "成功", f"已保存:\n{self.template_path}")
                self._refresh_template_combo()
            except Exception as e:
                QMessageBox.warning(self, "失败", f"保存失败:\n{str(e)}")
        else:
            self._save_as()

    def _save_as(self):
        if self._current_item_id and self._current_item_id in self._pages:
            self._pages[self._current_item_id].save_config(self.template)

        name = self.meta_name.text() or "新模板"
        meta = self._collect_meta()
        try:
            path = self.tm.save_as(self.template, name, meta)
            self.template_path = path
            self._refresh_template_combo()
            QMessageBox.information(self, "成功", f"已保存:\n{path}")
        except Exception as e:
            QMessageBox.warning(self, "失败", f"保存失败:\n{str(e)}")

    def _export_template(self):
        if not self.template_path:
            QMessageBox.warning(self, "提示", "请先保存或加载一个模板")
            return
        file, _ = QFileDialog.getSaveFileName(
            self, "导出模板", f"{self.meta_name.text() or 'template'}.json", "JSON文件 (*.json)"
        )
        if file:
            try:
                if self._current_item_id and self._current_item_id in self._pages:
                    self._pages[self._current_item_id].save_config(self.template)
                self.tm.export_template(self.template_path, file)
                QMessageBox.information(self, "成功", f"已导出:\n{file}")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"导出失败:\n{str(e)}")

    def _import_template(self):
        file, _ = QFileDialog.getOpenFileName(self, "导入模板", "", "JSON文件 (*.json)")
        if file:
            try:
                path = self.tm.import_template(file)
                if path:
                    self.template = self.tm.load_template(path)
                    self.template_path = path
                    self._load_all_pages()
                    self._refresh_template_combo()
                    from ..templates.template_manager import read_meta
                    self._load_meta_into_ui(read_meta(path))
                    QMessageBox.information(self, "成功", f"已导入:\n{path}")
            except Exception as e:
                QMessageBox.warning(self, "失败", f"导入失败:\n{str(e)}")

    def get_template(self) -> TemplateConfig:
        if self._current_item_id and self._current_item_id in self._pages:
            self._pages[self._current_item_id].save_config(self.template)
        return self.template
