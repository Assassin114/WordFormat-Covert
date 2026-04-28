"""
DocFormatter 数据模型
定义模板配置相关的数据类
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict
from enum import Enum


class PrintMode(Enum):
    """打印模式"""
    SINGLE = "single"      # 单面打印
    DUPLEX = "duplex"       # 双面打印


class NumberingMode(Enum):
    """编号模式"""
    GLOBAL = "global"      # 全局连续编号
    PER_TYPE = "per_type"   # 按类型独立编号


class PageNumberFormat(str, Enum):
    """页码格式"""
    ARABIC = "arabic"           # 阿拉伯数字 1, 2, 3...
    ROMAN_LOW = "roman_low"     # 小写罗马数字 i, ii, iii...
    ROMAN_UP = "roman_up"       # 大写罗马数字 I, II, III...
    LETTER_LOW = "letter_low"  # 小写字母 a, b, c...
    LETTER_UP = "letter_up"    # 大写字母 A, B, C...


class PageNumberPosition(str, Enum):
    """页码位置"""
    HEADER = "header"
    FOOTER = "footer"


class PageNumberAlignment(str, Enum):
    """页码对齐"""
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"


@dataclass
class FontConfig:
    """字体配置"""
    name: str = "宋体"           # 字体名
    size: float = 12.0          # 字号（磅）
    size_name: str = "小四"      # 中文字号名
    bold: bool = False
    italic: bool = False
    color: str = "#000000"      # RGB hex


@dataclass
class ParagraphConfig:
    """段落配置"""
    font: FontConfig = field(default_factory=FontConfig)
    line_spacing: float = 1.5    # 行距倍数
    line_spacing_type: str = "multiple"  # 固定值/最小值/单倍/多倍
    line_spacing_fixed: float = 20  # 固定值时使用（磅）
    line_spacing_min: float = 15   # 最小值时使用（磅）
    space_before: float = 0        # 段前间距（磅）
    space_after: float = 0          # 段后间距（磅）
    first_line_indent: float = 21  # 首行缩进（磅）


@dataclass
class HeadingConfig:
    """标题配置"""
    level: int = 1              # 级别 1-5
    font: FontConfig = field(default_factory=FontConfig)
    paragraph: ParagraphConfig = field(default_factory=ParagraphConfig)


@dataclass
class CaptionConfig:
    """题注配置"""
    figure_prefix: str = "图"
    table_prefix: str = "表"
    equation_prefix: str = "公式"
    position_figure: str = "below"   # below/above
    position_table: str = "above"     # below/above
    font: FontConfig = field(default_factory=FontConfig)
    paragraph: ParagraphConfig = field(default_factory=ParagraphConfig)


@dataclass
class PageNumberConfig:
    """页码配置"""
    format: PageNumberFormat = PageNumberFormat.ARABIC
    position: PageNumberPosition = PageNumberPosition.FOOTER
    alignment: PageNumberAlignment = PageNumberAlignment.CENTER
    font_name: str = "Times New Roman"
    font_size: float = 10.5
    font_bold: bool = False
    font_italic: bool = False


@dataclass
class HeaderFooterConfig:
    """页眉页脚配置"""
    # 封面/签署页/修改记录页
    cover_header_text: str = "文档标题"
    cover_show_header: bool = True
    cover_show_footer: bool = False
    cover_page_number: PageNumberConfig = field(default_factory=PageNumberConfig)
    
    # 目录页
    toc_header_text: str = "目录"
    toc_show_header: bool = True
    toc_show_footer: bool = True
    toc_page_number: PageNumberConfig = field(default_factory=PageNumberConfig)
    
    # 正文部分
    body_header_text: str = "正文"
    body_show_header: bool = True
    body_show_footer: bool = True
    body_page_number: PageNumberConfig = field(default_factory=PageNumberConfig)
    
    # 全局设置
    different_first_page: bool = False
    different_odd_even: bool = False


@dataclass
class CoverConfig:
    """封面配置"""
    enabled: bool = True
    template_path: Optional[str] = None  # 封面模板文件路径
    fields: Dict[str, str] = field(default_factory=dict)  # 字段名 -> 替换值


@dataclass
class SignatureConfig:
    """签署页配置"""
    enabled: bool = True
    template_path: Optional[str] = None
    fields: Dict[str, str] = field(default_factory=dict)


@dataclass
class RevisionConfig:
    """修改记录配置"""
    enabled: bool = True
    heading_format: Optional[HeadingConfig] = None


@dataclass
class TemplateConfig:
    """完整模板配置"""
    version: str = "1.0"
    cover: CoverConfig = field(default_factory=CoverConfig)
    signature: SignatureConfig = field(default_factory=SignatureConfig)
    revision: RevisionConfig = field(default_factory=RevisionConfig)
    headings: List[HeadingConfig] = field(default_factory=list)  # 5个
    body: ParagraphConfig = field(default_factory=ParagraphConfig)
    caption: CaptionConfig = field(default_factory=CaptionConfig)
    header_footer: HeaderFooterConfig = field(default_factory=HeaderFooterConfig)
    print_mode: PrintMode = PrintMode.SINGLE


def create_default_template() -> TemplateConfig:
    """创建默认模板配置"""
    level_names = ["黑体", "黑体", "楷体", "楷体", "宋体"]
    level_sizes = [16, 14, 12, 12, 10.5]
    level_bolds = [True, True, False, False, False]
    level_spacing_before = [24, 12, 12, 6, 6]
    level_spacing_after = [6, 6, 6, 6, 6]
    level_line_spacing = [2.0, 1.5, 1.5, 1.5, 1.5]
    
    headings = []
    for i in range(5):
        h = HeadingConfig(
            level=i+1,
            font=FontConfig(
                name=level_names[i],
                size=level_sizes[i],
                bold=level_bolds[i]
            ),
            paragraph=ParagraphConfig(
                line_spacing=level_line_spacing[i],
                space_before=level_spacing_before[i],
                space_after=level_spacing_after[i],
            )
        )
        headings.append(h)
    
    body = ParagraphConfig(
        font=FontConfig(name="宋体", size=10.5),
        line_spacing=1.5,
        first_line_indent=21
    )
    
    caption = CaptionConfig(
        font=FontConfig(name="宋体", size=10.5),
        paragraph=ParagraphConfig(
            line_spacing=1.5,
            space_before=0,
            space_after=0,
        )
    )
    
    header_footer = HeaderFooterConfig()
    
    return TemplateConfig(
        headings=headings,
        body=body,
        caption=caption,
        header_footer=header_footer,
    )


@dataclass
class FormatResult:
    """格式化结果"""
    success: bool = False
    input_path: str = ""
    output_path: str = ""
    figures_processed: int = 0
    tables_processed: int = 0
    equations_processed: int = 0
    errors: List[str] = field(default_factory=list)


@dataclass
class BatchResult:
    """批量处理结果"""
    total: int = 0
    success_count: int = 0
    failed_count: int = 0
    results: List[FormatResult] = field(default_factory=list)


@dataclass
class TableStyleConfig:
    """表格样式配置"""
    header_font: FontConfig = field(default_factory=lambda: FontConfig(name="宋体", size=10.5, bold=True))
    header_bg_color: str = "E8E8E8"   # 表头背景色
    body_font: FontConfig = field(default_factory=lambda: FontConfig(name="宋体", size=10.5))
    header_align: str = "center"     # 表头对齐
    body_align_rule: str = "text_length"  # 正文对齐规则
    body_align_threshold: int = 10    # 短文本阈值


@dataclass
class TableFontConfig:
    """表格字体配置"""
    regular_table: TableStyleConfig = field(default_factory=TableStyleConfig)
    record_table: TableStyleConfig = field(default_factory=TableStyleConfig)
