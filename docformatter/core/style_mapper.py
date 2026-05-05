"""
样式映射模块 (已废弃)
@deprecated: 此模块已被 document_analyzer.DocumentAnalyzer 替代。
"""

from typing import Dict, Optional, List

from ..models.template_config import TemplateConfig, HeadingConfig, ParagraphConfig


class StyleMapper:
    """
    分析原始文档样式，建立到目标模板的映射关系
    """
    
    def __init__(self, template: TemplateConfig):
        self.template = template
        self._heading_mapping_cache: Dict[str, HeadingConfig] = {}
        self._body_mapping: Optional[ParagraphConfig] = None
        self._analyzed = False
    
    def analyze(self, doc) -> Dict[str, object]:
        """
        分析文档样式，返回映射表
        
        Args:
            doc: python-docx Document 对象
        
        Returns:
            dict: 映射表
        """
        self._heading_mapping_cache.clear()
        
        # 扫描文档中的所有样式
        if doc.styles:
            for style in doc.styles:
                if not style.name:
                    continue
                
                style_name = style.name.lower()
                
                # 匹配标题级别
                for level in range(1, 6):
                    if self._match_heading_style(style_name, level):
                        self._heading_mapping_cache[style.name] = self.template.headings[level - 1]
                        break
                
                # 匹配正文
                if self._match_body_style(style_name):
                    self._body_mapping = self.template.body
        
        # 如果没有找到正文样式映射，使用默认
        if self._body_mapping is None:
            self._body_mapping = self.template.body
        
        self._analyzed = True
        
        return {
            'headings': self._heading_mapping_cache.copy(),
            'body': self._body_mapping,
        }
    
    def _match_heading_style(self, style_name: str, level: int) -> bool:
        """判断样式是否匹配指定级别的标题"""
        # 中文风格名
        if f'标题 {level}' in style_name or f'标题{level}' in style_name:
            return True
        # 英文风格名
        if f'heading {level}' in style_name:
            return True
        # 简化的标题
        if style_name == f'heading{level}':
            return True
        return False
    
    def _match_body_style(self, style_name: str) -> bool:
        """判断样式是否匹配正文"""
        if 'normal' in style_name:
            return True
        if 'body' in style_name:
            return True
        if style_name in ['正文', 'article', 'document']:
            return True
        return False
    
    def get_style_for_paragraph(self, para) -> object:
        """
        根据段落获取对应的目标配置
        
        Args:
            para: python-docx Paragraph 对象
        
        Returns:
            HeadingConfig 或 ParagraphConfig
        """
        # 优先通过样式名匹配
        if para.style and para.style.name:
            style_name = para.style.name
            if style_name in self._heading_mapping_cache:
                return self._heading_mapping_cache[style_name]
        
        # 通过字体大小等特征推断（备选）
        # 这里简化处理：直接返回正文配置
        return self._body_mapping
    
    def get_heading_level(self, para) -> Optional[int]:
        """
        获取段落对应的标题级别
        
        Args:
            para: python-docx Paragraph 对象
        
        Returns:
            int: 1-5 表示标题级别，None 表示非标题
        """
        if para.style and para.style.name:
            style_name = para.style.name.lower()
            
            for level in range(1, 6):
                if self._match_heading_style(style_name, level):
                    return level
        
        # 通过直接格式判断（备选方案）
        # 如果字体大小匹配某级标题的大小，也认为是该级别标题
        for run in para.runs:
            if run.font.size:
                size_pt = run.font.size.pt
                # 匹配标题字号
                for level, heading in enumerate(self.template.headings, 1):
                    if abs(size_pt - heading.font.size) < 0.5:
                        return level
        
        return None
