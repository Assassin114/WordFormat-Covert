# DocFormatter 软件设计说明书（SDD）

**版本：** V1.0
**日期：** 2026-04-28
**项目：** P004 - 软件开发
**技术栈：** Python 3.11 + PyQt6 + python-docx + lxml

---

## 1. 系统概述

### 1.1 系统架构

```
┌─────────────────────────────────────────────────────────────┐
│                        GUI Layer (PyQt6)                      │
│  MainWindow → TemplateConfig / BatchProcess / Word2MD        │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                     Core Engine Layer                         │
│  DocumentFormatter → StyleMapper / Numbering / TOCGenerator  │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                      Utils Layer                              │
│  OOXMLUtils / FileUtils / Logger                             │
└─────────────────────────────────────────────────────────────┘
```

### 1.2 目录结构

```
docformatter/
├── main.py                     # 程序入口
├── __init__.py
├── gui/
│   ├── __init__.py
│   ├── main_window.py          # 主窗口
│   ├── template_config.py      # 模板配置Tab
│   ├── batch_process.py        # 批量处理Tab
│   ├── word2md_tab.py          # Word转MD Tab
│   └── widgets/                # 公共组件
│       ├── __init__.py
│       ├── heading_config.py   # 标题格式配置组件
│       ├── body_config.py      # 正文格式配置组件
│       ├── caption_config.py   # 题注配置组件
│       └── file_selector.py    # 文件选择组件
├── core/
│   ├── __init__.py
│   ├── formatter.py            # 核心格式化引擎
│   ├── style_mapper.py         # 样式映射
│   ├── numbering.py             # 编号管理
│   ├── toc_generator.py        # 目录生成
│   └── word2md_converter.py    # Word转MD
├── utils/
│   ├── __init__.py
│   ├── ooxml_utils.py          # OOXML底层操作
│   ├── file_utils.py           # 文件工具
│   └── logger.py               # 日志工具
├── templates/
│   ├── __init__.py
│   ├── default_template.json   # 默认模板
│   └── template_io.py         # 模板加载/保存
├── models/
│   ├── __init__.py
│   └── template_config.py      # 模板配置数据类
├── requirements.txt
└── README.md
```

---

## 2. 数据模型

### 2.1 模板配置数据结构

```python
from dataclasses import dataclass, field
from typing import List, Optional
from enum import Enum

class PrintMode(Enum):
    SINGLE = "single"      # 单面打印
    DUPLEX = "duplex"      # 双面打印

class NumberingMode(Enum):
    BY_CHAPTER = "by_chapter"   # 按章节编号 "图 2-1"
    GLOBAL = "global"           # 全文统一编号 "图 1"

@dataclass
class FontConfig:
    name: str = "宋体"           # 字体名
    size: float = 12.0          # 字号（磅）
    bold: bool = False
    italic: bool = False
    color: str = "#000000"     # RGB hex

@dataclass
class ParagraphConfig:
    font: FontConfig
    line_spacing: float = 1.5   # 行距倍数
    space_before: float = 0     # 段前间距（磅）
    space_after: float = 0      # 段后间距（磅）
    first_line_indent: float = 0 # 首行缩进（磅）
    alignment: str = "left"     # left/center/right/justify

@dataclass
class HeadingConfig:
    level: int                  # 1-5
    font: FontConfig
    paragraph: ParagraphConfig
    outline_level: int = level  # 大纲级别

@dataclass
class CaptionConfig:
    prefix_figure: str = "图"    # 图序前缀
    prefix_table: str = "表"     # 表序前缀
    prefix_equation: str = "公式" # 公式序前缀
    numbering_mode: NumberingMode = NumberingMode.GLOBAL
    position_figure: str = "below"   # below/above
    position_table: str = "above"     # below/above
    separator: str = " "        # 如 "图 1-1" 中的分隔符

@dataclass
class HeaderFooterConfig:
    show_on_first: bool = False
    different_odd_even: bool = False
    header_font: FontConfig = field(default_factory=FontConfig)
    footer_font: FontConfig = field(default_factory=FontConfig)
    page_number_format: str = "normal"  # normal/roman

@dataclass
class CoverConfig:
    enabled: bool = True
    fields: dict = field(default_factory=dict)  # {field_name: placeholder}

@dataclass
class SignatureConfig:
    enabled: bool = True
    fields: dict = field(default_factory=dict)

@dataclass
class RevisionConfig:
    enabled: bool = True
    heading_format: HeadingConfig = None

@dataclass
class TemplateConfig:
    version: str = "1.0"
    cover: CoverConfig = field(default_factory=CoverConfig)
    signature: SignatureConfig = field(default_factory=SignatureConfig)
    revision: RevisionConfig = field(default_factory=RevisionConfig)
    headings: List[HeadingConfig] = field(default_factory=list)  # 5个
    body: ParagraphConfig = None
    caption: CaptionConfig = field(default_factory=CaptionConfig)
    header_footer: HeaderFooterConfig = field(default_factory=HeaderFooterConfig)
    print_mode: PrintMode = PrintMode.SINGLE

# 初始化默认值
def create_default_template() -> TemplateConfig:
    """创建默认模板配置"""
    headings = []
    for i in range(1, 6):
        level_names = ["黑体", "黑体", "楷体", "楷体", "宋体"]
        level_sizes = [16, 14, 12, 12, 10.5]
        headings.append(HeadingConfig(
            level=i,
            font=FontConfig(name=level_names[i-1], size=level_sizes[i-1], bold=(i<=2)),
            paragraph=ParagraphConfig(
                font=FontConfig(name=level_names[i-1], size=level_sizes[i-1], bold=(i<=2)),
                line_spacing=1.5 if i > 1 else 2.0,
                space_before=24 if i == 1 else 12,
                space_after=6 if i == 1 else 6
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
```

### 2.2 格式化结果

```python
@dataclass
class FormatResult:
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
    total: int
    success_count: int
    failed_count: int
    results: List[FormatResult] = field(default_factory=list)
```

---

## 3. 核心模块详细设计

### 3.1 formatter.py - 核心格式化引擎

```python
class DocumentFormatter:
    """
    核心格式化引擎，负责调度各子模块完成文档格式化
    """
    
    def __init__(self, template: TemplateConfig, cover_template_path: Optional[str] = None):
        self.template = template
        self.cover_template_path = cover_template_path
        self.style_mapper = StyleMapper(template)
        self.numbering_mgr = NumberingManager(template.caption)
        self.toc_gen = TOCGenerator()
        
    def format(self, input_path: str, output_path: str) -> FormatResult:
        """
        格式化单个文档
        
        处理流程：
        1. 解析文档
        2. 替换封面字段（如有）
        3. 映射样式
        4. 应用格式
        5. 处理图片（编号+同页）
        6. 处理表格（编号+同页）
        7. 处理公式（编号）
        8. 生成目录
        9. 处理页眉页脚
        10. 处理打印模式（空白页）
        11. 保存
        """
        result = FormatResult(
            success=False,
            input_path=input_path,
            output_path=output_path
        )
        
        try:
            # Step 1: 解析
            doc = Document(input_path)
            
            # Step 2: 封面替换
            if self.cover_template_path:
                self._apply_cover(doc)
            
            # Step 3: 样式映射
            style_mapping = self.style_mapper.analyze(doc)
            
            # Step 4: 应用格式
            self._apply_formats(doc, style_mapping)
            
            # Step 5-7: 处理图片/表格/公式
            result.figures_processed = self._process_images(doc)
            result.tables_processed = self._process_tables(doc)
            result.equations_processed = self._process_equations(doc)
            
            # Step 8: 生成目录
            self._generate_toc(doc)
            
            # Step 9: 页眉页脚
            self._apply_header_footer(doc)
            
            # Step 10: 打印模式
            self._handle_print_mode(doc)
            
            # Step 11: 保存
            doc.save(output_path)
            result.success = True
            
        except Exception as e:
            result.errors.append(str(e))
            
        return result
    
    def _apply_cover(self, doc: Document):
        """替换封面占位符"""
        for para in doc.paragraphs:
            for run in para.runs:
                # 匹配 {{FIELD_NAME}} 格式
                for key, value in self.template.cover.fields.items():
                    if f"{{{{{key}}}}}" in run.text:
                        run.text = run.text.replace(f"{{{{{key}}}}}", value)
    
    def _apply_formats(self, doc: Document, style_mapping: dict):
        """应用各级格式到文档元素"""
        for para in doc.paragraphs:
            # 确定段落对应的标题级别
            heading_level = self._get_heading_level(para, style_mapping)
            
            if heading_level:
                # 应用标题格式
                heading_config = self.template.headings[heading_level - 1]
                self._apply_paragraph_format(para, heading_config.paragraph)
                self._apply_font_format(para, heading_config.font)
                # 设置大纲级别
                para.style = f'Heading {heading_level}'
            else:
                # 应用正文格式
                self._apply_paragraph_format(para, self.template.body)
                self._apply_font_format(para, self.template.body.font)
    
    def _get_heading_level(self, para, style_mapping: dict) -> Optional[int]:
        """根据样式名称判断标题级别"""
        if para.style and para.style.name:
            style_name = para.style.name.lower()
            for level in range(1, 6):
                if f'heading {level}' in style_name or f'标题{level}' in style_name:
                    return level
        # 检查直接格式（基于font size粗判）
        return None
    
    def _apply_paragraph_format(self, para, config: ParagraphConfig):
        """应用段落格式（使用底层XML）"""
        pPr = para._element.get_or_add_pPr()
        # 行距
        spacing = pPr.get_or_add_spacing()
        spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', 
                    str(int(config.line_spacing * 240)))
        spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', 'auto')
        # 段前段后
        if config.space_before:
            spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before', 
                        str(int(config.space_before * 20)))
        if config.space_after:
            spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after',
                        str(int(config.space_after * 20)))
    
    def _apply_font_format(self, para, font_config: FontConfig):
        """应用字体格式"""
        for run in para.runs:
            run.font.name = font_config.name
            run.font.size = Pt(font_config.size)
            run.font.bold = font_config.bold
            run.font.italic = font_config.italic
            # 颜色
            if font_config.color != "#000000":
                run.font.color.rgb = RGBColor.from_string(font_config.color[1:])
    
    def _process_images(self, doc: Document) -> int:
        """处理所有图片：添加图序、同页控制"""
        count = 0
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                # 找到关联的段落，在下方插入图序
                count += 1
                caption_text = f"{self.template.caption.prefix_figure} {count}"
                # 添加题注段落
                # 设置孤行控制（与上图同页）
        return count
    
    def _process_tables(self, doc: Document) -> int:
        """处理所有表格：添加表序、同页控制"""
        count = 0
        for table in doc.tables:
            # 在表格上方插入表序
            count += 1
            caption_text = f"{self.template.caption.prefix_table} {count}"
            # 设置孤行控制（与下表同页）
        return count
    
    def _process_equations(self, doc: Document) -> int:
        """处理所有公式：编号、居中"""
        count = 0
        for para in doc.paragraphs:
            # 识别公式（OMML公式或图片公式）
            if self._is_equation_para(para):
                count += 1
                # 公式居中
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # 在右侧添加编号
                # (1), (2), (3) ...
        return count
    
    def _is_equation_para(self, para) -> bool:
        """判断段落是否为公式"""
        # 检查是否是OMML公式
        for run in para.runs:
            if run._element.xml.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}oMath') != -1:
                return True
        return False
    
    def _generate_toc(self, doc: Document):
        """在目录占位符处插入TOC域"""
        # 查找"目录"标题位置
        # 在其后插入TOC域
        pass
    
    def _apply_header_footer(self, doc: Document):
        """应用页眉页脚设置"""
        section = doc.sections[0]
        header = section.header
        footer = section.footer
        
        # 处理首页不同
        if not self.template.header_footer.show_on_first:
            section.different_first_page_header_footer = True
        
        # 处理奇偶页不同
        if self.template.header_footer.different_odd_even:
            section.even_odd_header_footer = True
    
    def _handle_print_mode(self, doc: Document):
        """处理单双面打印模式"""
        # 双面模式：在奇数页后添加空白页
        if self.template.print_mode == PrintMode.DUPLEX:
            # 遍历所有section，检查最后一页是否为奇数
            # 如果是，添加空白页使下一页从偶数开始
            pass
```

### 3.2 style_mapper.py - 样式映射

```python
class StyleMapper:
    """
    分析原始文档样式，建立到目标模板的映射关系
    """
    
    def __init__(self, template: TemplateConfig):
        self.template = template
        
    def analyze(self, doc: Document) -> dict:
        """
        分析文档样式，返回映射表
        
        Returns:
            dict: {
                'heading1': HeadingConfig,
                'heading2': HeadingConfig,
                ...
                'normal': ParagraphConfig,
                'caption': CaptionConfig
            }
        """
        mapping = {}
        
        # 扫描所有样式
        for style in doc.styles:
            if not style.name:
                continue
                
            style_name = style.name.lower()
            
            # 匹配标题
            for level in range(1, 6):
                if f'heading {level}' in style_name or f'标题 {level}' in style_name:
                    mapping[style.name] = self.template.headings[level - 1]
                    
            # 匹配正文
            if 'normal' in style_name or 'body' in style_name or '正文' in style_name:
                mapping[style.name] = self.template.body
                
            # 匹配题注
            if 'caption' in style_name or '题注' in style_name:
                mapping[style.name] = self.template.caption
        
        return mapping
    
    def get_style_for_paragraph(self, para, mapping: dict) -> Optional[object]:
        """
        根据段落原始样式获取目标配置
        优先使用样式名称匹配，其次使用直接格式推断
        """
        if para.style and para.style.name in mapping:
            return mapping[para.style.name]
        
        # 基于font size推断（备选方案）
        for run in para.runs:
            if run.font.size:
                size_pt = run.font.size.pt
                # 根据字号匹配可能对应的标题级别
                ...
        
        return self.template.body  # 默认返回正文格式
```

### 3.3 numbering.py - 编号管理

```python
class NumberingManager:
    """
    图序、表序、公式序的全局统一编号管理
    """
    
    def __init__(self, caption_config: CaptionConfig):
        self.caption_config = caption_config
        self.figure_count = 0
        self.table_count = 0
        self.equation_count = 0
        self.numbering_mode = caption_config.numbering_mode
        
    def reset(self):
        """重置计数器"""
        self.figure_count = 0
        self.table_count = 0
        self.equation_count = 0
        
    def get_next_figure_num(self) -> int:
        self.figure_count += 1
        return self.figure_count
    
    def get_next_table_num(self) -> int:
        self.table_count += 1
        return self.table_count
    
    def get_next_equation_num(self) -> int:
        self.equation_count += 1
        return self.equation_count
    
    def format_figure_caption(self, num: int) -> str:
        """生成图序字符串"""
        prefix = self.caption_config.prefix_figure
        if self.numbering_mode == NumberingMode.BY_CHAPTER:
            # 需要额外参数：当前章节号
            return f"{prefix} {num}"
        return f"{prefix} {num}"
    
    def format_table_caption(self, num: int) -> str:
        prefix = self.caption_config.prefix_table
        return f"{prefix} {num}"
    
    def format_equation_caption(self, num: int) -> str:
        return f"({num})"
    
    def process_document(self, doc: Document):
        """
        扫描文档，收集所有图片/表格/公式并分配序号
        这个方法在formatter.py之前调用，用于预扫描
        """
        self.reset()
        
        # 扫描图片
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                self.get_next_figure_num()
        
        # 扫描表格
        for table in doc.tables:
            self.get_next_table_num()
        
        # 扫描公式
        for para in doc.paragraphs:
            if self._is_equation_para(para):
                self.get_next_equation_num()
    
    def _is_equation_para(self, para) -> bool:
        """同formatter.py中的实现"""
        for run in para.runs:
            if run._element.xml.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}oMath') != -1:
                return True
        return False
```

### 3.4 toc_generator.py - 目录生成

```python
class TOCGenerator:
    """
    目录生成器
    """
    
    def generate(self, doc: Document, heading_configs: List[HeadingConfig]):
        """
        生成目录
        
        流程：
        1. 收集所有标题段落
        2. 按出现顺序确定章节编号
        3. 在"目录"标题后插入TOC域
        """
        # 查找目录位置
        toc_heading_idx = None
        for i, para in enumerate(doc.paragraphs):
            if self._is_toc_heading(para):
                toc_heading_idx = i
                break
        
        if toc_heading_idx is None:
            return  # 未找到目录标题，跳过
        
        # 收集标题
        headings = []
        for para in doc.paragraphs[toc_heading_idx + 1:]:
            level = self._get_heading_level(para)
            if level:
                text = para.text.strip()
                headings.append((level, text))
        
        # 构建TOC条目
        toc_entries = []
        for level, text in headings:
            indent = "  " * (level - 1)
            entry = f"{indent}{text}"
            toc_entries.append(entry)
        
        # 插入TOC域（使用SDT包裹）
        # Word的TOC通过域代码实现，这里使用简单替换方案
        toc_para = doc.paragraphs[toc_heading_idx + 1]
        
        # 清空现有内容，插入TOC标题
        # 实际Word中的TOC需要通过 OOXML 域代码实现
        # 此处简化处理，生成静态文本目录
        for entry in toc_entries:
            new_para = doc.add_paragraph(entry, style='Normal')
    
    def _is_toc_heading(self, para) -> bool:
        """判断是否为目录标题"""
        text = para.text.strip()
        return text in ['目录', 'CONTENTS', 'Table of Contents']
    
    def _get_heading_level(self, para) -> Optional[int]:
        """获取标题级别"""
        if para.style and 'Heading' in para.style.name:
            try:
                return int(para.style.name.split()[-1])
            except:
                pass
        return None
```

---

## 4. GUI模块设计

### 4.1 主窗口布局

```
┌──────────────────────────────────────────────────────────────┐
│  文件(F)  编辑(E)  帮助(H)                                   │
├──────────────────────────────────────────────────────────────┤
│  ┌────────────────┬────────────────┬────────────────┐        │
│  │   模板配置      │    批量处理    │   Word转MD     │        │
│  └────────────────┴────────────────┴────────────────┘        │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│                    [Tab内容区域]                              │
│                                                              │
│                                                              │
├──────────────────────────────────────────────────────────────┤
│  状态: 就绪                                    [进度条]      │
└──────────────────────────────────────────────────────────────┘
```

### 4.2 模板配置Tab布局

```
┌─────────────────────────────────────────────────────────────┐
│  基本信息                        │  标题格式                    │
│  ┌────────────────────────────┐ │  ┌─────────────────────┐   │
│  │ 封面模板: [选择文件]        │ │  │ 级别1  字体:[宋体▼]   │   │
│  │ 签署页:   [选择文件]        │ │  │      字号:[16▼]      │   │
│  │ 修改记录: [√]启用           │ │  │      [□加粗]         │   │
│  └────────────────────────────┘ │  │ 级别2  ...           │   │
│                                 │  │ ...                  │   │
│  正文格式                        │  └─────────────────────┘   │
│  ┌────────────────────────────┐ │                            │
│  │ 字体: [宋体▼] 字号:[12▼]   │ │  题注格式                    │
│  │ 行距: [1.5倍▼]             │ │  ┌─────────────────────┐   │
│  │ 首行缩进: [2字符]          │ │  │ 图序: [图▼] 位置:[下方▼]│   │
│  └────────────────────────────┘ │  │ 表序: [表▼] 位置:[上方▼]│   │
│                                 │  └─────────────────────┘   │
│  页眉页脚  打印模式              │                             │
│  ┌──────────┬──────────────┐   │  [保存模板] [加载模板]        │
│  │ □首页不同│ ○单面 ○双面  │   │                             │
│  └──────────┴──────────────┘   │                             │
└─────────────────────────────────────────────────────────────┘
```

### 4.3 批量处理Tab布局

```
┌─────────────────────────────────────────────────────────────┐
│  源文件                                                   │
│  ┌─────────────────────────────────────────────────────┐   │
│  │ [选择文件]  [选择文件夹]                             │   │
│  │                                                     │   │
│  │ 已选择: 3个文件                                      │   │
│  │  - doc1.docx                                        │   │
│  │  - doc2.docx                                        │   │
│  │  - doc3.docx                                        │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                             │
│  模板文件: [选择文件]                          [选择文件]   │
│                                                             │
│  输出目录: [___________________________] [浏览...]          │
│                                                             │
│  ┌─────────────────────────────────────────────────────┐   │
│  │ 日志:                                               │   │
│  │ [2026-04-28 00:47:01] 开始处理...                   │   │
│  │ [2026-04-28 00:47:02] doc1.docx 完成                │   │
│  │ [2026-04-28 00:47:03] doc2.docx 完成                │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                             │
│  [开始处理]                              进度: ████░░ 66%   │
│                                                             │
│  统计: 成功 2 / 失败 0 / 总计 3                              │
└─────────────────────────────────────────────────────────────┘
```

---

## 5. OOXML底层操作（ooxml_utils.py）

### 5.1 孤行控制（ widow/orphan control）

```python
def set_widow_control(paragraph, enable=True):
    """
    设置段落孤行控制
    widowControl: 段首控制（段落最后一行单独在页面顶部）
    orphanControl: 段尾控制（段落第一行单独在页面底部）
    """
    pPr = paragraph._element.get_or_add_pPr()
    widowControl = pPr.find(
        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}widowControl'
    )
    if widowControl is None:
        widowControl = OxmlElement('w:widowControl')
        pPr.append(widowControl)
    widowControl.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 
                      'true' if enable else 'false')
```

### 5.2 精确行距设置

```python
def set_line_spacing(paragraph, line_value: float, line_rule='auto'):
    """
    设置精确行距
    
    Args:
        line_value: 行距值
            - 倍数: 1.5, 2.0 等
            - 磅值: 20, 24 等（当 line_rule='exact'）
        line_rule: 'auto'（倍数）或 'exact'（固定值）
    """
    pPr = paragraph._element.get_or_add_pPr()
    spacing = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    
    # 240 相当于 1 倍行距（twips）
    if line_rule == 'auto':
        line_val = str(int(line_value * 240))
        spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', 'auto')
    else:
        line_val = str(int(line_value * 20))  # 磅值转twips
        spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', 'exact')
    
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', line_val)
```

### 5.3 插入分页符

```python
def insert_page_break_after(paragraph):
    """在段落后插入分页符"""
    p = paragraph._element
    # 创建 <w:br> 元素，type="page"
    br = OxmlElement('w:br')
    br.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'page')
    
    # 创建新的 run 包含分页符
    new_run = OxmlElement('w:r')
    new_run.append(br)
    p.append(new_run)
```

---

## 6. 接口定义

### 6.1 核心Formatter API

```python
from typing import List, Optional
from dataclasses import dataclass

@dataclass
class FormatOptions:
    preserve_existing_toc: bool = True
    update_fields: bool = True
    generate_toc: bool = True

class IDocumentFormatter:
    """文档格式化器接口"""
    
    def format(
        self, 
        input_path: str, 
        output_path: str,
        template_path: str,
        cover_fields: Optional[dict] = None,
        options: Optional[FormatOptions] = None
    ) -> FormatResult:
        """格式化单个文档"""
        raise NotImplementedError
    
    def batch_format(
        self,
        input_paths: List[str],
        output_dir: str,
        template_path: str,
        cover_fields: Optional[dict] = None,
        options: Optional[FormatOptions] = None,
        progress_callback: Optional[callable] = None
    ) -> BatchResult:
        """批量格式化文档"""
        raise NotImplementedError
```

### 6.2 模板配置API

```python
import json
from pathlib import Path

def load_template(path: str) -> TemplateConfig:
    """从JSON文件加载模板配置"""
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return _dict_to_template(data)

def save_template(config: TemplateConfig, path: str):
    """保存模板配置到JSON文件"""
    data = _template_to_dict(config)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def _dict_to_template(data: dict) -> TemplateConfig:
    """字典转TemplateConfig"""
    # 递归转换嵌套对象
    pass

def _template_to_dict(config: TemplateConfig) -> dict:
    """TemplateConfig转字典"""
    pass
```

---

## 7. 错误处理策略

### 7.1 错误分类

| 错误类型 | 处理方式 | 示例 |
|----------|----------|------|
| 文件不存在 | 跳过，记录日志 | 输入文件被删除 |
| 文件损坏 | 跳过，记录日志 | docx格式不完整 |
| 模板无效 | 中断，提示用户 | JSON格式错误 |
| 内存不足 | 跳过，记录日志 | 600页文档内存溢出 |
| 权限不足 | 跳过，记录日志 | 输出目录无写权限 |

### 7.2 批量处理策略

```python
def batch_format(self, input_paths, output_dir, ...):
    results = []
    for path in input_paths:
        try:
            result = self.format(path, output_path, ...)
            results.append(result)
        except Exception as e:
            results.append(FormatResult(
                success=False,
                input_path=path,
                output_path=None,
                errors=[str(e)]
            ))
    
    # 生成批量总结
    success = sum(1 for r in results if r.success)
    failed = sum(1 for r in results if not r.success)
    
    return BatchResult(
        total=len(input_paths),
        success_count=success,
        failed_count=failed,
        results=results
    )
```

---

## 8. 性能优化策略

### 8.1 600页文档优化

1. **流式读取**: 使用 `python-docx` 的增量解析，避免全量加载
2. **批量操作XML**: 减少 `save()` 调用次数，内存中完成所有修改后一次性保存
3. **多进程批处理**: 使用 `multiprocessing.Pool` 并行处理多个文件

### 8.2 内存管理

```python
import gc

def format_with_cleanup(self, input_path, output_path, ...):
    try:
        doc = Document(input_path)
        # 处理...
        doc.save(output_path)
    finally:
        # 强制垃圾回收
        del doc
        gc.collect()
```

---

## 9. 依赖关系

```
requirements.txt
─────────────────
PyQt6 >= 6.6.0
python-docx >= 1.1.0
lxml >= 5.0.0
```

---

## 10. 待确认事项

1. **封面模板字段**: 亮哥提供封面docx后确认占位符名称
2. **目录生成时机**: 已确认-格式化后生成
3. **中文字体名**: Windows环境，无需跨平台兼容
4. **输出文件名**: 已确认-原文件名_formatted.docx
5. **横向页面处理**: 已确认-整页横向，前后分节符，多图支持
6. **交叉引用处理**: 已确认-需要同步更新REF域显示文本

---

**END OF SDD**

---

## 11. 分节与横向页面处理

### 11.1 场景描述

**原始文档中的情况：**
```
[分节符 - 新节开始]
████████████ 横向页面 ██████████████
██ 图片1（宽度超出纵向页面）██
██ 图片2（可能多张拼在一起）██
████████████           ██████████████
[分节符 - 下一节开始]
...正文（恢复纵向页）...
```

**特点：**
- 整页横向
- 前后都有分节符
- 可能包含1~N张图片

### 11.2 实现方案

```python
def _handle_landscape_sections(self, doc: Document):
    """
    保留分节后的横向页面设置
    
    处理流程：
    1. 遍历所有 section
    2. 检查 pgSz/orient 属性
    3. 如果是 landscape（横向），保留该分节配置
    4. 后续处理时不在横向节内插入额外分页符
    """
    for section in doc.sections:
        pgSz = section._element.find('.//{...}pgSz')
        if pgSz is not None:
            orient = pgSz.get('{...}orient')
            w = pgSz.get('{...}w')   # 宽度（twips）
            h = pgSz.get('{...}h')   # 高度（twips）
            
            if orient == 'landscape' or (w and h and int(w) > int(h)):
                # 横向节：宽度 > 高度
                # 保留该节的页面尺寸和方向设置
                pass
```

### 11.3 处理要求

| 要求 | 说明 |
|------|------|
| 分节符保留 | 格式化时保留横向节的前后分节符 |
| 横向设置保留 | 图片所在节的 orient="landscape" 完整保留 |
| 多图支持 | 横向节内可能有1~N张图片，全部保留 |
| 后续节恢复 | 横向节后的下一节自动恢复纵向，无需额外操作 |
| 图序处理 | 横向页中的图片，图序同样位于图下方 |

### 11.4 python-docx 访问分节属性

```python
from docx.oxml.ns import qn

# 获取节的页面方向
section = doc.sections[0]
pgSz = section._element.find(qn('w:pgSz'))
orient = pgSz.get(qn('w:orient'))  # 'portrait' 或 'landscape'
w = pgSz.get(qn('w:w'))  # 宽度
h = pgSz.get(qn('w:h'))  # 高度

# 判断是否为横向
is_landscape = (orient == 'landscape') or (w and h and int(w) > int(h))
```

### 11.5 待确认更新

- **目录生成时机**（已确认：格式化后）→ SDD 4.
- **输出文件后缀**（已确认：_formatted）→ SDD 6.
- **横向页面处理**（已确认）→ 本节
- **封面占位符**：亮哥后续提供 → SDD 10.


---

## 12. 交叉引用处理

### 12.1 问题描述

原始文档中可能存在"如图1所示"、"见表2-3"等交叉引用。重新编号后引用关系必须同步更新。

**Word中的交叉引用机制：**
- 题注段落创建书签（bookmark）
- 正文引用使用 `REF` 域指向书签
- 重新编号后，REF域不会自动更新

**示例：**
```
原文：见图1所示...  → REF域显示 "1"
重编号后：题注变成"图 2"，但REF域仍显示"1" → 引用关系错误
```

### 12.2 处理流程

```
format()
  │
  ├─ 1. 解析文档
  ├─ 2. 扫描题注书签和REF引用，建立映射表
  │     ├─ 记录 fig_1, fig_2 等书签
  │     └─ 记录所有 REF fig_X 域
  │
  ├─ 3. 样式映射
  ├─ 4. 应用格式
  ├─ 5. 处理图片（重新编号）
  │     ├─ 更新题注："图 1" → "图 2"
  │     ├─ 重命名书签：fig_1 → fig_2
  │     └─ 更新所有REF域的显示文本
  │
  ├─ 6. 处理表格...
  ├─ 7. 处理公式...
  └─ 8. 保存
```

### 12.3 书签命名规则

| 类型 | 书签名前缀 | 示例 |
|------|-----------|------|
| 图片题注 | `fig_` | `fig_1`, `fig_2` |
| 表格题注 | `tbl_` | `tbl_1`, `tbl_2` |
| 公式题注 | `eq_` | `eq_1`, `eq_2` |

### 12.4 Word域代码结构

**题注书签：**
```xml
<w:bookmarkStart w:id="1" w:name="fig_1"/>
<w:r><w:t>图 1</w:t></w:r>
<w:bookmarkEnd w:id="1"/>
```

**REF引用：**
```xml
<w:fldChar w:fldCharType="begin"/>
<w:instrText> REF fig_1 \h </w:instrText>
<w:fldChar w:fldCharType="separate"/>
<w:t>1</w:t>  ← 需要更新为新编号
<w:fldChar w:fldCharType="end"/>
```

### 12.5 更新策略

1. **扫描阶段**：收集所有题注书签和REF引用，建立映射表
2. **编号阶段**：图片/表格重新编号时，同步重命名书签
3. **更新阶段**：遍历所有REF域，根据书签名查找新编号并更新显示文本

---

## 13. 纯文本引用处理

### 13.1 问题描述

除了REF域引用外，正文中可能存在直接书写的图片/表格编号，如：
- `如图1所示` — 纯文本，不是REF域
- `见下图3`
- `如表 2-1 所示`
- `式(5)`

这些纯文本引用不会自动更新，需要在重新编号后同步替换。

### 13.2 匹配模式

```python
TEXT_REF_PATTERNS = {
    'fig': [
        r'图\s*(\d+)',           # 图1, 图 1, 图   1
        r'如图\s*\d+\s*所示',    # 如图1所示
        r'见下图\s*\d+',         # 见下图1
    ],
    'tbl': [
        r'表\s*(\d+)',           # 表1, 表 1
        r'见表\s*\d+\s*所示',   # 如表1所示
        r'见下表\s*\d+',        # 见下表1
    ],
    'eq': [
        r'公式\s*(\d+)',         # 公式1
        r'见公式\s*\d+',        # 见公式1
    ]
}
```

### 13.3 处理流程

```
1. scan_document() 扫描时：
   - 遍历所有 paragraph 和 run
   - 用正则匹配纯文本引用
   - 记录 (type, match_text, num, para_element, run_element)

2. set_numbering_map() 设置编号映射：
   - old_to_new: {('fig', 1): 3, ('fig', 2): 4, ...}

3. update_all_refs_in_document() 更新：
   - 对每条纯文本引用，找到对应run
   - 执行文本替换：match_text -> 新编号
```

### 13.4 示例

```
原始文档：
  见图1所示，blabla...
  见下图3，blabla...

编号映射：
  fig: 1→3, 2→4, 3→5  ...

更新后：
  见图3所示，blablo...
  见下图5，blabla...
```

### 13.5 注意事项

- 纯文本引用可能在多个run中分裂（如"见"和"图1"在不同run）
- 当前实现假设引用文本在单个run内
- 对于跨run的引用，需要更复杂的合并逻辑
