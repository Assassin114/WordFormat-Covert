"""DocFormatter 数据模型"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict
from enum import Enum


class PrintMode(Enum):
    SINGLE = "single"
    DUPLEX = "duplex"


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
    name: str = "宋体"
    size: float = 10.5
    size_name: str = "五号"
    bold: bool = False
    italic: bool = False
    color: str = "#000000"


@dataclass
class ParagraphConfig:
    font: FontConfig = field(default_factory=FontConfig)
    line_spacing: float = 1.5
    line_spacing_type: str = "multiple"
    line_spacing_fixed: float = 20
    line_spacing_min: float = 15
    space_before: float = 0
    space_after: float = 0
    first_line_indent: float = 21
    alignment: str = "left"


@dataclass
class HeadingConfig:
    level: int = 1
    font: FontConfig = field(default_factory=FontConfig)
    paragraph: ParagraphConfig = field(default_factory=ParagraphConfig)


@dataclass
class CaptionConfig:
    figure_prefix: str = "图"
    table_prefix: str = "表"
    equation_prefix: str = "公式"
    position_figure: str = "below"
    position_table: str = "above"
    font: FontConfig = field(default_factory=lambda: FontConfig(name="宋体", size=10.5))
    paragraph: ParagraphConfig = field(default_factory=lambda: ParagraphConfig())


@dataclass
class PageNumberConfig:
    format: PageNumberFormat = PageNumberFormat.ARABIC
    position: PageNumberPosition = PageNumberPosition.FOOTER
    alignment: PageNumberAlignment = PageNumberAlignment.CENTER
    font_name: str = "Times New Roman"
    font_size: float = 10.5
    font_bold: bool = False
    font_italic: bool = False


@dataclass
class HeaderFooterConfig:
    cover_header_text: str = "文档标题"
    cover_show_header: bool = True
    cover_show_footer: bool = False
    cover_page_number: PageNumberConfig = field(default_factory=PageNumberConfig)
    toc_header_text: str = "目录"
    toc_show_header: bool = True
    toc_show_footer: bool = True
    toc_page_number: PageNumberConfig = field(default_factory=PageNumberConfig)
    body_header_text: str = "正文"
    body_show_header: bool = True
    body_show_footer: bool = True
    body_page_number: PageNumberConfig = field(default_factory=PageNumberConfig)
    different_first_page: bool = False
    different_odd_even: bool = False


@dataclass
class CoverConfig:
    enabled: bool = True
    template_path: Optional[str] = None
    fields: Dict[str, str] = field(default_factory=dict)


@dataclass
class SignatureConfig:
    enabled: bool = True
    template_path: Optional[str] = None
    fields: Dict[str, str] = field(default_factory=dict)


@dataclass
class RevisionConfig:
    enabled: bool = True
    heading_format: Optional[HeadingConfig] = None


@dataclass
class TableStyleConfig:
    header_font: FontConfig = field(default_factory=lambda: FontConfig(name="宋体", size=10.5, bold=True))
    header_bg_color: str = "E8E8E8"
    body_font: FontConfig = field(default_factory=lambda: FontConfig(name="宋体", size=10.5))
    header_align: str = "center"
    body_align_rule: str = "text_length"
    body_align_threshold: int = 10


@dataclass
class TableFontConfig:
    regular_table: TableStyleConfig = field(default_factory=TableStyleConfig)
    record_table: TableStyleConfig = field(default_factory=TableStyleConfig)


@dataclass
class TemplateConfig:
    version: str = "1.0"
    cover: CoverConfig = field(default_factory=CoverConfig)
    signature: SignatureConfig = field(default_factory=SignatureConfig)
    revision: RevisionConfig = field(default_factory=RevisionConfig)
    headings: List[HeadingConfig] = field(default_factory=list)
    body: ParagraphConfig = field(default_factory=ParagraphConfig)
    caption: CaptionConfig = field(default_factory=CaptionConfig)
    table_font: TableFontConfig = field(default_factory=TableFontConfig)
    header_footer: HeaderFooterConfig = field(default_factory=HeaderFooterConfig)
    print_mode: PrintMode = PrintMode.SINGLE


@dataclass
class FormatResult:
    success: bool = False
    input_path: str = ""
    output_path: str = ""
    figures_processed: int = 0
    tables_processed: int = 0
    equations_processed: int = 0
    errors: List[str] = field(default_factory=list)


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
    
    headings = []
    for i in range(5):
        h = HeadingConfig(
            level=i+1,
            font=FontConfig(name=level_names[i], size=level_sizes[i], bold=level_bolds[i]),
            paragraph=ParagraphConfig(line_spacing=level_line_spacing[i], space_before=level_spacing_before[i], space_after=level_spacing_after[i])
        )
        headings.append(h)
    
    body = ParagraphConfig(font=FontConfig(name="宋体", size=10.5), line_spacing=1.5, first_line_indent=21.0)
    caption = CaptionConfig(font=FontConfig(name="宋体", size=10.5), paragraph=ParagraphConfig(line_spacing=1.5))
    table_font = TableFontConfig()
    header_footer = HeaderFooterConfig()
    
    return TemplateConfig(headings=headings, body=body, caption=caption, table_font=table_font, header_footer=header_footer)
