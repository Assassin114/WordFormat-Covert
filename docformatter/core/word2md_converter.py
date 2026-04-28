"""
Word转Markdown转换器
支持：段落、标题、表格、图片、列表等元素的转换
"""

import re
import os
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from ..utils.logger import get_logger

logger = get_logger()


class Word2MDConverter:
    """Word文档转Markdown转换器"""
    
    def __init__(self):
        self.table_count = 0
        self.figure_count = 0
    
    def convert(self, docx_path: str, output_path: str = None) -> str:
        """
        转换Word文档为Markdown
        
        Args:
            docx_path: Word文档路径
            output_path: 输出MD文件路径（可选，不提供则返回Markdown文本）
        
        Returns:
            转换后的Markdown文本
        """
        self.table_count = 0
        self.figure_count = 0
        
        doc = Document(docx_path)
        md_lines = []
        
        # 处理每个段落
        for para in doc.paragraphs:
            md_line = self._convert_paragraph(para)
            if md_line:
                md_lines.append(md_line)
        
        # 处理表格
        for table in doc.tables:
            md_table = self._convert_table(table)
            if md_table:
                md_lines.append(md_table)
        
        md_content = '\n\n'.join(md_lines)
        
        # 如果提供了输出路径，写入文件
        if output_path:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(md_content)
            logger.info(f"转换完成: {output_path}")
        
        return md_content
    
    def _convert_paragraph(self, para) -> str:
        """转换单个段落为Markdown"""
        text = para.text.strip()
        
        if not text:
            return ''
        
        # 检测标题级别
        style_name = para.style.name if para.style else ''
        
        if style_name.startswith('Heading'):
            try:
                level = int(style_name.replace('Heading ', ''))
                # 标题使用 ATX 风格 (# ~ ######)
                return '#' * level + ' ' + text
            except:
                pass
        
        # 处理行内格式（粗体、斜体）
        md_text = self._convert_runs(para.runs)
        
        # 列表项检测
        if self._is_list_item(para):
            prefix = self._get_list_prefix(para)
            return prefix + md_text
        
        # 普通段落
        return md_text
    
    def _convert_runs(self, runs) -> str:
        """转换段落中的runs（处理粗体、斜体等）"""
        result = []
        
        for run in runs:
            text = run.text
            if not text:
                continue
            
            # 处理行内格式
            if run.bold:
                text = f'**{text}**'
            if run.font.italic:
                text = f'*{text}*'
            
            result.append(text)
        
        return ''.join(result)
    
    def _is_list_item(self, para) -> bool:
        """检测是否是列表项"""
        pPr = para._element.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        if pPr is not None:
            numPr = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr')
            if numPr is not None:
                return True
        return False
    
    def _get_list_prefix(self, para) -> str:
        """获取列表前缀（- 或 1.）"""
        # 简化处理：返回无序列表
        return '- '
    
    def _convert_table(self, table) -> str:
        """转换表格为Markdown"""
        self.table_count += 1
        table_id = self.table_count
        
        md_lines = []
        md_lines.append(f'<table id="table-{table_id}">')
        
        for row_idx, row in enumerate(table.rows):
            cells = [cell.text.strip() for cell in row.cells]
            md_lines.append('| ' + ' | '.join(cells) + ' |')
            
            # 表头分隔行
            if row_idx == 0:
                md_lines.append('| ' + ' | '.join(['---'] * len(cells)) + ' |')
        
        md_lines.append('</table>')
        
        return '\n'.join(md_lines)
    
    def convert_file(self, input_path: str, output_path: str = None) -> str:
        """
        转换Word文件为Markdown文件
        
        Args:
            input_path: 输入docx路径
            output_path: 输出md路径（可选，默认为同名.md文件）
        
        Returns:
            输出文件路径
        """
        input_path = Path(input_path)
        
        if output_path is None:
            output_path = input_path.with_suffix('.md')
        else:
            output_path = Path(output_path)
        
        self.convert(str(input_path), str(output_path))
        return str(output_path)


def convert_word_to_markdown(input_path: str, output_path: str = None) -> str:
    """便捷函数：转换Word文档为Markdown"""
    converter = Word2MDConverter()
    return converter.convert_file(input_path, output_path)


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("用法: python3 word2md_converter.py <input.docx> [output.md]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    result = convert_word_to_markdown(input_file, output_file)
    print(f"转换完成: {result}")
