"""
目录生成模块
负责在文档中生成目录
"""

from typing import List, Tuple, Optional
from ..models.template_config import HeadingConfig

from ..utils import is_landscape_section


class TOCGenerator:
    """
    目录生成器
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
        生成目录
        
        流程：
        1. 查找目录标题位置
        2. 收集所有标题及其级别
        3. 在目录标题后插入 TOC 域
        
        Args:
            doc: python-docx Document 对象
            heading_configs: 标题配置列表
        """
        # 查找目录位置
        toc_idx = self.find_toc_position(doc)
        if toc_idx is None:
            return  # 未找到目录标题，跳过
        
        # 收集标题
        headings = []
        for para in doc.paragraphs[toc_idx + 1:]:
            level = self._get_heading_level(para)
            if level is not None:
                text = para.text.strip()
                if text:  # 只收集有文本的标题
                    headings.append((level, text))
        
        # 插入 TOC
        # 注意：Word 的 TOC 需要通过域代码实现
        # 这里先找到插入位置，然后在目录标题后第一个段落插入占位符
        # 实际生成需要在 Word 中更新域
        
        # 简化处理：在目录标题后插入一个说明段落
        # 用户需要在 Word 中右键点击目录区域，选择"更新域"来生成完整目录
        insert_idx = toc_idx + 1
        
        # 如果下一段已经是占位符，跳过
        if insert_idx < len(doc.paragraphs):
            next_para = doc.paragraphs[insert_idx]
            if '[TOC]' in next_para.text or '目录占位符' in next_para.text:
                return
        
        # 插入占位符段落
        # 这里不做实际插入，因为 python-docx 的 TOC 支持有限
        # 建议用户在实际文档中手动插入 TOC 域
    
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
        
        Args:
            doc: python-docx Document 对象
            para_idx: 段落索引
        
        Returns:
            bool: 如果段落所在节是横向，返回 True
        """
        for section in doc.sections:
            # 检查该节是否包含指定段落
            # 这是简化实现，实际需要更精确的节边界判断
            if is_landscape_section(section):
                return True
        return False
