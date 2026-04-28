"""
统一字体格式配置工具
所有字体/段落格式化都通过此类调用，确保一致性
"""

from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


class FontFormatter:
    """
    统一字体格式配置工具类
    """
    
    @staticmethod
    def apply_font(run, font_config):
        """应用字体到 run"""
        if not font_config:
            return
        
        run.font.name = font_config.name
        
        if font_config.size:
            run.font.size = Pt(font_config.size / 2)
        
        run.bold = font_config.bold
        run.italic = font_config.italic
        
        if font_config.name:
            try:
                run._element.rPr.rFonts.set(qn('w:eastAsia'), font_config.name)
            except:
                pass
        
        if font_config.color and font_config.color != "#000000":
            try:
                rgb = RGBColor.from_string(font_config.color.lstrip('#'))
                run.font.color.rgb = rgb
            except:
                pass
    
    @staticmethod
    def apply_paragraph(para, paragraph_config):
        """应用段落格式"""
        if not paragraph_config:
            return
        
        pPr = para._element.find(qn('w:pPr'))
        if pPr is None:
            pPr = OxmlElement('w:pPr')
            para._element.insert(0, pPr)
        
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
        
        if paragraph_config.space_before:
            spacing.set(qn('w:before'), str(int(paragraph_config.space_before * 20)))
        if paragraph_config.space_after:
            spacing.set(qn('w:after'), str(int(paragraph_config.space_after * 20)))
        
        if paragraph_config.first_line_indent:
            ind = pPr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(qn('w:firstLine'), str(int(paragraph_config.first_line_indent * 20)))
        
        alignment_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        if paragraph_config.alignment in alignment_map:
            para.alignment = alignment_map[paragraph_config.alignment]
