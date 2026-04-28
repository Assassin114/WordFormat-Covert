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
    DUPLEX = "duplex"      # 双面打印


class NumberingMode(Enum):
    """编号模式"""
    GLOBAL = "global"      # 全文统一编号
    BY_CHAPTER = "by_chapter"  # 按章节编号


@dataclass
class FontConfig:
    """字体配置"""
    name: str = "宋体"           # 字体名
    size: float = 12.0          # 字号（磅）
    size_name: str = "小四"      # 中文字号名（一号/小一/二号/.../六号）
    bold: bool = False
    italic: bool = False
    color: str = "#000000"      # RGB hex


@dataclass
class ParagraphConfig:
    """段落配置"""
    font: FontConfig = field(default_factory=FontConfig)
    line_spacing: float = 1.5    # 行距倍数
    space_before: float = 0      # 段前间距（磅）
    space_after: float = 0       # 段后间距（磅）
    first_line_indent: float = 0 # 首行缩进（磅）
    alignment: str = "left"      # left/center/right/justify


@dataclass
class HeadingConfig:
    """标题配置"""
    level: int                  # 1-5
    font: FontConfig = field(default_factory=FontConfig)
    paragraph: ParagraphConfig = field(default_factory=ParagraphConfig)
    outline_level: int = 0      # 大纲级别


@dataclass
class CaptionConfig:
    """题注配置"""
    prefix_figure: str = "图"    # 图序前缀
    prefix_table: str = "表"     # 表序前缀
    prefix_equation: str = "公式" # 公式序前缀
    numbering_mode: NumberingMode = NumberingMode.GLOBAL
    position_figure: str = "below"   # below/above
    position_table: str = "above"     # below/above


@dataclass
class HeaderFooterConfig:
    """页眉页脚配置"""
    show_on_first: bool = False      # 首页不同
    different_odd_even: bool = False # 奇偶页不同
    header_font: FontConfig = field(default_factory=FontConfig)
    footer_font: FontConfig = field(default_factory=FontConfig)
    page_number_format: str = "normal"  # normal/roman


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
    for i in range(1, 6):
        headings.append(HeadingConfig(
            level=i,
            font=FontConfig(
                name=level_names[i-1],
                size=level_sizes[i-1],
                bold=level_bolds[i-1]
            ),
            paragraph=ParagraphConfig(
                font=FontConfig(
                    name=level_names[i-1],
                    size=level_sizes[i-1],
                    bold=level_bolds[i-1]
                ),
                line_spacing=level_line_spacing[i-1],
                space_before=level_spacing_before[i-1],
                space_after=level_spacing_after[i-1]
            ),
            outline_level=i
        ))

    return TemplateConfig(
        body=ParagraphConfig(
            font=FontConfig(name="宋体", size=12),
            line_spacing=1.5,
            first_line_indent=21,  # 2字符
            space_before=0,
            space_after=0
        ),
        caption=CaptionConfig(),
        headings=headings,
        header_footer=HeaderFooterConfig()
    )


@dataclass
class FormatResult:
    """单个文档格式化结果"""
    success: bool
    input_path: str
    output_path: str
    figures_processed: int = 0
    tables_processed: int = 0
    equations_processed: int = 0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


@dataclass
class BatchResult:
    """批量格式化结果"""
    total: int
    success_count: int
    failed_count: int
    results: List[FormatResult] = field(default_factory=list)

# 中文标准字号映射
CHINESE_FONT_SIZES = {
    '一号': 26,
    '小一': 24,
    '二号': 22,
    '小二': 18,
    '三号': 16,
    '小三': 15,
    '四号': 14,
    '小四': 12,
    '五号': 10.5,
    '小五': 9,
    '六号': 7.5,
    '小六': 6.5,
}

# 反向映射（字号转中文）
FONT_SIZE_TO_CHINESE = {v: k for k, v in CHINESE_FONT_SIZES.items()}
