"""
目录生成模块
负责在文档中生成目录
"""

from typing import List, Tuple, Optional
from ..models.template_config import HeadingConfig

from ..utils import is_landscape_section

WML_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


class TOCGenerator:
    """
    目录生成器
    @deprecated: TOC 域的实际生成逻辑已移至 md2word_converter._create_toc_page()
    """

    def __init__(self):
        self._toc_heading_idx: Optional[int] = None
    
    def find_toc_position(self, doc) -> Optional[int]:
        """
        查找目录标题位置
        
        Args:
            doc: python-docx Document 对象
        
        Returns:
            int: 目录标题段落的索引，None 表示未找到
        """
        for i, para in enumerate(doc.paragraphs):
            if self._is_toc_heading(para):
                return i
        return None
    
    def _is_toc_heading(self, para) -> bool:
        """判断是否为目录标题"""
        text = para.text.strip()
        toc_keywords = ['目录', 'CONTENTS', 'Table of Contents', 'table of contents']
        return text in toc_keywords
    
    def generate_toc(self, doc, heading_configs: List[HeadingConfig]):
        """
        @deprecated: 实际 TOC 域生成已迁移至 MD2WordConverter._create_toc_page()
        """
        return
    
    def _get_heading_level(self, para) -> Optional[int]:
        """
        获取段落对应的标题级别
        
        Args:
            para: python-docx Paragraph 对象
        
        Returns:
            int: 1-5，None 表示非标题
        """
        if para.style and para.style.name:
            style_name = para.style.name
            
            # 检查 Heading 1-5
            if 'Heading' in style_name or '标题' in style_name:
                try:
                    # 从样式名中提取级别数字
                    for level in range(1, 6):
                        if str(level) in style_name:
                            return level
                    # 如果样式名只包含 "Heading"，假设为 Heading 1
                    if style_name in ['Heading', '标题']:
                        return 1
                except Exception:
                    pass
        
        return None
    
    def is_landscape_section_before(self, doc, para_idx: int) -> bool:
        """
        检查指定段落之前的节是否为横向

        通过检查该段落所在 section 的 pgSz/orient 来判断。
        """
        for section in doc.sections:
            sect_pr = section._element.find(WML_NS + 'sectPr') if section._element.tag != WML_NS + 'sectPr' else section._element
            if sect_pr is None:
                sect_pr = section._element
            pg_sz = sect_pr.find(WML_NS + 'pgSz')
            if pg_sz is not None:
                orient = pg_sz.get(WML_NS + 'orient', '')
                if orient == 'landscape':
                    return True
        return False
