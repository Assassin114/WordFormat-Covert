"""
统一字体格式配置工具
所有字体/段落格式化都通过此类调用，确保一致性
"""

from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from ..models.template_config import FontConfig, ParagraphConfig, TableStyleConfig, TableFontConfig


class FontFormatter:
    """
    统一字体格式配置工具类
    """
    
    @staticmethod
    def apply_font(run, font_config: FontConfig):
        """
        应用字体到 run
        
        Args:
            run: docx run对象
            font_config: FontConfig配置
        """
        if not font_config:
            return
        
        run.font.name = font_config.name
        
        # 字号转换：FontConfig.size 是磅值，run.font.size 需要半磅
        if font_config.size:
            run.font.size = Pt(font_config.size / 2)
        
        run.bold = font_config.bold
        run.italic = font_config.italic
        
        # 设置中文字体
        if font_config.name:
            try:
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_config.name)
            except:
                pass
        
        # 颜色
        if font_config.color and font_config.color != "#000000":
            try:
                rgb = RGBColor.from_string(font_config.color.lstrip('#'))
                run.font.color.rgb = rgb
            except:
                pass
    
    @staticmethod
    def apply_paragraph(para, paragraph_config: ParagraphConfig):
        """
        应用段落格式
        
        Args:
            para: docx paragraph对象
            paragraph_config: ParagraphConfig配置
        """
        if not paragraph_config:
            return
        
        pPr = para._element.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            para._element.insert(0, pPr)
        
        # 行距
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = OxmlElement('w:spacing')
            pPr.append(spacing)
        
        ls_type = paragraph_config.line_spacing_type
        if ls_type == "fixed":
            spacing.set(qn('w:line'), str(int(paragraph_config.line_spacing_fixed * 20)))
            spacing.set(qn('w:lineRule'), 'exact')
        elif ls_type == "atLeast":
            spacing.set(qn('w:line'), str(int(paragraph_config.line_spacing_min * 20)))
            spacing.set(qn('w:lineRule'), 'atLeast')
        elif ls_type == "single":
            spacing.set(qn('w:line'), '240')
            spacing.set(qn('w:lineRule'), 'auto')
        else:
            spacing.set(qn('w:line'), str(int(paragraph_config.line_spacing * 240)))
            spacing.set(qn('w:lineRule'), 'auto')
        
        # 段前段后
        if paragraph_config.space_before:
            spacing.set(qn('w:before'), str(int(paragraph_config.space_before * 20)))
        if paragraph_config.space_after:
            spacing.set(qn('w:after'), str(int(paragraph_config.space_after * 20)))
        
        # 首行缩进
        if paragraph_config.first_line_indent:
            ind = pPr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(qn('w:firstLine'), str(int(paragraph_config.first_line_indent * 20)))
        
        # 对齐方式
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        if paragraph_config.alignment in alignment_map:
            para.alignment = alignment_map[paragraph_config.alignment]
    
    @staticmethod
    def apply_style_to_run(run, font_config: FontConfig):
        """兼容性别名"""
        FontFormatter.apply_font(run, font_config)
    
    @staticmethod
    def apply_style_to_paragraph(para, paragraph_config: ParagraphConfig):
        """兼容性别名"""
        FontFormatter.apply_paragraph(para, paragraph_config)


class TableFontFormatter:
    """
    表格字体格式配置
    区分基础表（常规表格）和记录表（测试/试验记录表）
    """
    
    def __init__(self, table_config: TableFontConfig = None):
        self.table_config = table_config or self._default_table_config()
    
    def _default_table_config(self) -> TableFontConfig:
        """默认表格配置"""
        return TableFontConfig(
            regular_table=TableStyleConfig(
                header_font=FontConfig(name="宋体", size=10.5, bold=True),
                header_bg_color="E8E8E8",
                body_font=FontConfig(name="宋体", size=10.5),
                header_align="center",
                body_align_rule="text_length",
                body_align_threshold=10,
            ),
            record_table=TableStyleConfig(
                header_font=FontConfig(name="宋体", size=10.5, bold=True),
                header_bg_color="E8E8E8",
                body_font=FontConfig(name="宋体", size=10.5),
                header_align="center",
                body_align_rule="text_length",
                body_align_threshold=10,
            ),
        )
    
    def get_table_style(self, table_type: str) -> TableStyleConfig:
        """获取表格样式配置"""
        if table_type == "record_table":
            return self.table_config.record_table
        else:
            return self.table_config.regular_table
    
    def apply_table_style(self, table, table_type: str):
        """
        应用表格样式
        
        Args:
            table: docx table对象
            table_type: "regular_table" 或 "record_table"
        """
        style = self.get_table_style(table_type)
        
        for row_idx, row in enumerate(table.rows):
            is_header = (row_idx == 0)
            
            for cell in row.cells:
                for para in cell.paragraphs:
                    if is_header:
                        font_cfg = style.header_font
                        para.alignment = self._get_alignment(style.header_align)
                    else:
                        font_cfg = style.body_font
                        text = para.text.strip()
                        if len(text) <= style.body_align_threshold:
                            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        else:
                            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    for run in para.runs:
                        FontFormatter.apply_font(run, font_cfg)
    
    def _get_alignment(self, align_name: str):
        """获取对齐方式"""
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        return alignment_map.get(align_name, WD_ALIGN_PARAGRAPH.LEFT)
