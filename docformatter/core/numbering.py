"""
编号管理模块
负责图序、表序、公式序的全局统一编号
"""

from typing import Optional

from ..models.template_config import CaptionConfig, NumberingMode


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
        """获取下一个图序号"""
        self.figure_count += 1
        return self.figure_count
    
    def get_next_table_num(self) -> int:
        """获取下一个表序号"""
        self.table_count += 1
        return self.table_count
    
    def get_next_equation_num(self) -> int:
        """获取下一个公式编号"""
        self.equation_count += 1
        return self.equation_count
    
    def format_figure_caption(self, num: int, chapter_num: Optional[int] = None) -> str:
        """
        生成图序字符串
        
        Args:
            num: 图序号
            chapter_num: 章节号（如果使用 BY_CHAPTER 模式）
        
        Returns:
            str: 格式化后的图序，如 "图 1" 或 "图 2-1"
        """
        prefix = self.caption_config.prefix_figure
        if self.numbering_mode == NumberingMode.BY_CHAPTER and chapter_num is not None:
            return f"{prefix} {chapter_num}-{num}"
        return f"{prefix} {num}"
    
    def format_table_caption(self, num: int, chapter_num: Optional[int] = None) -> str:
        """
        生成表序字符串
        """
        prefix = self.caption_config.prefix_table
        if self.numbering_mode == NumberingMode.BY_CHAPTER and chapter_num is not None:
            return f"{prefix} {chapter_num}-{num}"
        return f"{prefix} {num}"
    
    def format_equation_caption(self, num: int) -> str:
        """
        生成公式编号字符串
        
        Returns:
            str: 格式化后的编号，如 "(1)" 或 "(2-1)"
        """
        return f"({num})"
    
    def process_document(self, doc) -> dict:
        """
        扫描文档，预计算所有图片/表格/公式的数量
        
        Returns:
            dict: {'figures': count, 'tables': count, 'equations': count}
        """
        self.reset()
        
        # 扫描图片（通过关系 + drawing 元素双重检测）
        seen_images = set()
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                rel_id = rel.rId
                if rel_id not in seen_images:
                    seen_images.add(rel_id)
                    self.get_next_figure_num()

        # 补充：扫描 paragraph 中的 drawing/inline/anchor（遗漏的内嵌图片）
        for para in doc.paragraphs:
            xmlns_dml = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
            if (para._element.findall('.//' + WML_NS + 'drawing') or
                para._element.findall('.//' + xmlns_dml + 'inline') or
                para._element.findall('.//' + xmlns_dml + 'anchor')):
                # 检查是否已经通过 rels 计数过
                for blip in para._element.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
                    embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')
                    if embed and embed not in seen_images:
                        seen_images.add(embed)
                        self.get_next_figure_num()
        
        # 扫描表格
        for table in doc.tables:
            self.get_next_table_num()
        
        # 扫描公式
        for para in doc.paragraphs:
            if self._is_equation_para(para):
                self.get_next_equation_num()
        
        return {
            'figures': self.figure_count,
            'tables': self.table_count,
            'equations': self.equation_count,
        }
    
    def _is_equation_para(self, para) -> bool:
        """判断段落是否为公式"""
        # 检查是否是 OMML 公式
        for run in para.runs:
            try:
                if run._element.xml.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}oMath') != -1:
                    return True
            except Exception:
                pass
        return False
