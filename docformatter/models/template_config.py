"""DocFormatter 数据模型"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict
from enum import Enum


class PrintMode(Enum):
    SINGLE = "single"
    DUPLEX = "duplex"


class SectionBreakType(Enum):
    NEXT_PAGE = "next_page"
    CONTINUOUS = "continuous"
    EVEN_PAGE = "even_page"
    ODD_PAGE = "odd_page"


class EquationFormat(Enum):
    KEEP_ORIGINAL = "keep"
    CONVERT_TO_OMML = "omml"
    RENDER_AS_IMAGE = "image"


class NumberingMode(Enum):
    GLOBAL = "global"
    PER_TYPE = "per_type"


class PageNumberFormat(str, Enum):
    ARABIC = "arabic"
    ROMAN_LOW = "roman_low"
    ROMAN_UP = "roman_up"
    LETTER_LOW = "letter_low"
    LETTER_UP = "letter_up"


class PageNumberPosition(str, Enum):
    HEADER = "header"
    FOOTER = "footer"


class PageNumberAlignment(str, Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


@dataclass
class FontConfig:
    cn_name: str = "宋体"
    en_name: str = "Times New Roman"
    size: float = 10.5
    size_name: str = "五号"
    bold: bool = False
    italic: bool = False
    color: str = "#000000"


@dataclass
class ParagraphConfig:
    line_spacing: float = 1.5
    line_spacing_type: str = "multiple"
    line_spacing_fixed: float = 20
    line_spacing_min: float = 15
    space_before: float = 0
    space_after: float = 0
    first_line_indent: float = 0
    alignment: str = "left"


# ==================== 页面设置 ====================

@dataclass
class PageConfig:
    width: float = 210.0           # 纸张宽度 (mm)
    height: float = 297.0          # 纸张高度 (mm)
    orientation: str = "portrait"  # portrait / landscape
    margin_top: float = 37.0       # 上边距 (mm)
    margin_bottom: float = 35.0    # 下边距 (mm)
    margin_left: float = 28.0      # 左边距 (mm)
    margin_right: float = 26.0     # 右边距 (mm)
    gutter: float = 0              # 装订线 (mm)
    gutter_position: str = "left"  # left / top
    page_number_start: int = 1     # 起始页码


# ==================== 封面 ====================

@dataclass
class CoverConfig:
    enabled: bool = True
    template_path: Optional[str] = None
    fields: Dict[str, str] = field(default_factory=dict)
    heading_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="黑体", size=42, size_name="初号", bold=True))
    heading_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))


# ==================== 签署页 ====================

@dataclass
class SignatureConfig:
    enabled: bool = True
    template_path: Optional[str] = None
    title: str = "签署页"
    heading_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="黑体", size=26, size_name="一号", bold=True))
    heading_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))


# ==================== 修改记录页 ====================

@dataclass
class RevisionConfig:
    enabled: bool = True
    heading_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="黑体", size=16, size_name="三号", bold=True))
    heading_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))
    table_header_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="宋体", size=10.5, size_name="五号", bold=True))
    table_header_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))
    table_body_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="宋体", size=10.5, size_name="五号"))
    table_body_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))


# ==================== 目录页 ====================

@dataclass
class TOCConfig:
    enabled: bool = True
    heading_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="黑体", size=16, size_name="三号", bold=True))
    heading_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))
    entry_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="宋体", size=12, size_name="小四"))
    entry_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(line_spacing=1.5))
    entry_indent: float = 0


# ==================== 正文 ====================

@dataclass
class BodyConfig:
    font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="宋体", size=10.5, size_name="五号"))
    paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(line_spacing=1.5, first_line_indent=21.0))


# ==================== 标题 ====================

# 标题编号风格枚举
NUMBER_STYLES = {
    "none": "无编号",
    "decimal": "1. / 2.",
    "chinese": "一、/ 二、",
    "upper_letter": "A. / B.",
    "lower_letter": "a. / b.",
    "upper_roman": "I. / II.",
    "lower_roman": "i. / ii.",
}
NUMBER_STYLE_KEYS = list(NUMBER_STYLES.keys())
NUMBER_STYLE_LABELS = list(NUMBER_STYLES.values())

MULTI_NUMBER_STYLES = {
    "none": "无编号",
    "decimal": "1.1. / 1.2.",
    "chinese": "一、(一)",
    "upper_letter": "A.1.",
    "lower_letter": "a.1.",
}
MULTI_NUMBER_LABELS = list(MULTI_NUMBER_STYLES.values())


@dataclass
class HeadingConfig:
    level: int = 1
    number_style: str = "none"         # 单级编号风格 key
    number_multi: bool = False         # 是否多级继承
    font: FontConfig = field(default_factory=FontConfig)
    paragraph: ParagraphConfig = field(default_factory=ParagraphConfig)


# ==================== 表格 ====================

@dataclass
class TableStyleConfig:
    header_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="宋体", size=10.5, size_name="五号", bold=True))
    header_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))
    header_bg_color: str = "E8E8E8"
    body_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="宋体", size=10.5, size_name="五号"))
    body_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))


@dataclass
class TableFontConfig:
    regular_table: TableStyleConfig = field(default_factory=TableStyleConfig)
    record_table: TableStyleConfig = field(default_factory=TableStyleConfig)


# ==================== 题注 ====================

@dataclass
class CaptionConfig:
    figure_prefix: str = "图"
    table_prefix: str = "表"
    equation_prefix: str = "公式"
    position_figure: str = "below"
    position_table: str = "above"
    numbering_mode: NumberingMode = NumberingMode.GLOBAL
    font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="宋体", size=10.5, size_name="五号"))
    paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))


# ==================== 代码块 ====================

@dataclass
class CodeBlockConfig:
    font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="Consolas", en_name="Consolas", size=10.5, size_name="五号"))
    paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(line_spacing=1.0, first_line_indent=0))
    bg_color: str = "#F0F0F0"


# ==================== 页眉页脚 ====================

@dataclass
class HFPageConfig:
    header_text: str = ""
    show_header: bool = True
    show_footer: bool = True
    show_page_number: bool = False
    page_number_format: PageNumberFormat = PageNumberFormat.ROMAN_UP
    page_number_position: PageNumberPosition = PageNumberPosition.FOOTER
    page_number_alignment: PageNumberAlignment = PageNumberAlignment.CENTER
    page_number_font: FontConfig = field(default_factory=lambda: FontConfig(cn_name="宋体", en_name="Times New Roman", size=10.5, size_name="五号"))
    page_number_paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig(alignment="center"))


@dataclass
class HeaderFooterConfig:
    cover_page: HFPageConfig = field(default_factory=lambda: HFPageConfig(show_footer=False, show_page_number=False))
    toc_page: HFPageConfig = field(default_factory=lambda: HFPageConfig(header_text="目录", show_page_number=True))
    body_page: HFPageConfig = field(default_factory=lambda: HFPageConfig(header_text="正文", show_page_number=True))
    different_first_page: bool = False
    different_odd_even: bool = False


# ==================== 模板配置汇总 ====================

@dataclass
class TemplateConfig:
    version: str = "1.0"
    page: PageConfig = field(default_factory=PageConfig)
    cover: CoverConfig = field(default_factory=CoverConfig)
    signature: SignatureConfig = field(default_factory=SignatureConfig)
    revision: RevisionConfig = field(default_factory=RevisionConfig)
    toc: TOCConfig = field(default_factory=TOCConfig)
    headings: List[HeadingConfig] = field(default_factory=list)
    body: BodyConfig = field(default_factory=BodyConfig)
    table_font: TableFontConfig = field(default_factory=TableFontConfig)
    caption: CaptionConfig = field(default_factory=CaptionConfig)
    code_block: CodeBlockConfig = field(default_factory=CodeBlockConfig)
    header_footer: HeaderFooterConfig = field(default_factory=HeaderFooterConfig)
    print_mode: PrintMode = PrintMode.SINGLE
    equation_format: EquationFormat = EquationFormat.KEEP_ORIGINAL


@dataclass
class FormatResult:
    success: bool = False
    input_path: str = ""
    output_path: str = ""
    figures_processed: int = 0
    tables_processed: int = 0
    equations_processed: int = 0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


@dataclass
class BatchResult:
    total: int = 0
    success_count: int = 0
    failed_count: int = 0
    results: List[FormatResult] = field(default_factory=list)


def create_default_template() -> TemplateConfig:
    level_names = ["黑体", "黑体", "楷体", "楷体", "宋体"]
    level_sizes = [26.0, 22.0, 16.0, 14.0, 10.5]
    level_bolds = [True, True, False, False, False]
    level_spacing_before = [24.0, 12.0, 12.0, 6.0, 6.0]
    level_spacing_after = [6.0, 6.0, 6.0, 6.0, 6.0]
    level_line_spacing = [2.0, 1.5, 1.5, 1.5, 1.5]
    # 默认编号风格: 中文 → 1. → (1) → ① → 无
    default_num_styles = ["chinese", "decimal", "decimal", "decimal", "none"]
    default_multi = [True, True, True, False, False]

    headings = []
    for i in range(5):
        h = HeadingConfig(
            level=i + 1,
            number_style=default_num_styles[i],
            number_multi=default_multi[i],
            font=FontConfig(cn_name=level_names[i], size=level_sizes[i], bold=level_bolds[i]),
            paragraph=ParagraphConfig(
                line_spacing=level_line_spacing[i],
                space_before=level_spacing_before[i],
                space_after=level_spacing_after[i],
            ),
        )
        headings.append(h)

    body = BodyConfig(
        font=FontConfig(cn_name="宋体", size=10.5, size_name="五号"),
        paragraph=ParagraphConfig(line_spacing=1.5, first_line_indent=21.0),
    )

    return TemplateConfig(headings=headings, body=body)
