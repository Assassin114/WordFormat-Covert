"""
交叉引用管理模块
负责扫描、更新Word文档中的交叉引用（REF域 + 纯文本引用）
"""

from typing import Dict, List, Tuple, Optional
import re
from lxml import etree

from ..utils.logger import get_logger

logger = get_logger()

# Word namespace
WML_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


class CrossReferenceManager:
    """
    交叉引用管理器
    
    负责：
    1. 扫描文档中的题注书签和REF引用
    2. 扫描正文中纯文本的图片/表格/公式编号引用
    3. 建立编号映射关系
    4. 重新编号时同步更新所有引用
    """
    
    # 纯文本引用模式
    # 例如：如图1所示、见下图1、如图 1 等
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
    
    def __init__(self):
        # 书签信息：name -> (type, current_number)
        self.bookmarks: Dict[str, Tuple[str, int]] = {}
        
        # REF域引用信息：bookmark_name -> element
        self.ref_fields: Dict[str, List[etree._Element]] = {}
        
        # 纯文本引用信息：(ref_type, match_text, old_num, para_element, run_element)
        self.text_refs: List[Tuple[str, str, int, etree._Element, etree._Element]] = []
        
        # 原始编号 -> 新编号 的映射（用于批量更新）
        self.num_mapping: Dict[Tuple[str, int], int] = {}  # (type, old_num) -> new_num
    
    def scan_document(self, doc) -> dict:
        """
        扫描文档，收集所有题注书签、REF引用和纯文本引用

        Returns:
            dict: 包含 bookmarks, refs, text_refs
        """
        self.bookmarks.clear()
        self.ref_fields.clear()
        self.text_refs.clear()

        # 编译正则表达式
        compiled_patterns = {}
        for ref_type, patterns in self.TEXT_REF_PATTERNS.items():
            compiled_patterns[ref_type] = [re.compile(p) for p in patterns]

        # 扫描纯文本引用（段落级别）
        for para in doc.paragraphs:
            for run in para.runs:
                run_text = run.text
                if not run_text:
                    continue
                for ref_type, patterns in compiled_patterns.items():
                    for pattern in patterns:
                        for match in pattern.finditer(run_text):
                            num_str = match.group(1) if match.groups() else match.group(0)
                            try:
                                num = int(num_str)
                                self.text_refs.append((
                                    ref_type,
                                    match.group(0),
                                    num,
                                    para._element,
                                    run._element
                                ))
                            except ValueError:
                                pass

        # 扫描书签和 REF 域（段落级别）
        for para in doc.paragraphs:
            # 查找段落中的 bookmarkStart
            for bm_start in para._element.iter(WML_NS + 'bookmarkStart'):
                name = bm_start.get(WML_NS + 'name')
                if name and self._is_caption_bookmark(name):
                    bookmark_type, num = self._parse_bookmark_name(name)
                    if bookmark_type:
                        self.bookmarks[name] = (bookmark_type, num)

            # 查找段落中的 REF 域
            instr_texts = para._element.findall('.//' + WML_NS + 'instrText')
            for instr in instr_texts:
                if instr.text and 'REF' in instr.text:
                    bookmark_name = self._extract_ref_target(instr.text)
                    if bookmark_name:
                        if bookmark_name not in self.ref_fields:
                            self.ref_fields[bookmark_name] = []
                        # 存储段落元素（而非 body 或其他祖先）
                        self.ref_fields[bookmark_name].append(para._element)

        return {
            'bookmarks': self.bookmarks.copy(),
            'refs': {k: v.copy() for k, v in self.ref_fields.items()},
            'text_refs': self.text_refs.copy()
        }
    
    def _is_caption_bookmark(self, name: str) -> bool:
        """判断书签名是否为题注书签"""
        prefixes = ['fig_', 'tbl_', 'eq_', 'figure_', 'table_', 'equation_']
        return any(name.startswith(p) for p in prefixes)
    
    def _parse_bookmark_name(self, name: str) -> Tuple[Optional[str], int]:
        """从书签名中解析类型和编号"""
        for prefix, type_name in [('fig_', 'fig'), ('figure_', 'fig'),
                                    ('tbl_', 'tbl'), ('table_', 'tbl'),
                                    ('eq_', 'eq'), ('equation_', 'eq')]:
            if name.startswith(prefix):
                try:
                    num = int(name[len(prefix):])
                    return type_name, num
                except ValueError:
                    pass
        return None, 0
    
    def _extract_ref_target(self, instr_text: str) -> Optional[str]:
        """从REF域指令中提取目标书签名称"""
        match = re.search(r'REF\s+(\w+)\s*', instr_text)
        if match:
            return match.group(1)
        return None
    
    def set_numbering_map(self, old_to_new: Dict[Tuple[str, int], int]):
        """
        设置编号映射，用于批量更新
        
        Args:
            old_to_new: {(type, old_num): new_num, ...}
                       例如: {('fig', 1): 3, ('fig', 2): 4}
        """
        self.num_mapping = old_to_new.copy()
    
    def update_all_refs_in_document(self, doc):
        """
        根据编号映射，更新文档中所有引用
        
        1. 更新REF域的显示文本
        2. 更新纯文本引用
        """
        # 建立快速查找：(type, old_num) -> new_num
        for (ref_type, old_num), new_num in self.num_mapping.items():
            self._update_refs_for_number(doc, ref_type, old_num, new_num)
    
    def _update_refs_for_number(self, doc, ref_type: str, old_num: int, new_num: int):
        """更新特定类型和编号的所有引用"""
        prefix_map = {'fig': 'fig_', 'tbl': 'tbl_', 'eq': 'eq_'}
        prefix = prefix_map.get(ref_type, '')
        bookmark_name = f"{prefix}{old_num}"
        
        # 1. 更新REF域引用
        if bookmark_name in self.ref_fields:
            for elem in self.ref_fields[bookmark_name]:
                self._update_ref_display_text(elem, str(new_num))
        
        # 2. 更新纯文本引用
        for text_ref in self.text_refs:
            r_type, match_text, num, para_elem, run_elem = text_ref
            if r_type == ref_type and num == old_num:
                # 替换run中的文本
                old_text = run_elem.text
                new_text = old_text.replace(match_text, f"{match_text[0]}{new_num}")
                run_elem.text = new_text
                logger.debug(f"更新纯文本引用: {match_text} -> {new_text[0]}{new_num}")
    
    def _update_ref_display_text(self, para_elem: etree._Element, new_num: str):
        """
        更新 REF 域的显示文本

        REF 域结构（同一段落内，兄弟级 w:r 元素）：
        <w:r><w:fldChar w:fldCharType="begin"/></w:r>
        <w:r><w:instrText> REF xxx </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>旧编号</w:t></w:r>  ← 更新这里
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
        """
        found_sep = False
        for child in para_elem:
            if child.tag != WML_NS + 'r':
                continue
            fld_char = child.find(WML_NS + 'fldChar')
            if fld_char is not None:
                fld_type = fld_char.get(WML_NS + 'fldCharType')
                if fld_type == 'separate':
                    found_sep = True
                    continue
                elif fld_type == 'end':
                    found_sep = False
                    continue
            if found_sep:
                t_elem = child.find(WML_NS + 't')
                if t_elem is not None:
                    t_elem.text = new_num
                    found_sep = False
    
    def get_captions_by_type(self, ref_type: str) -> List[int]:
        """
        获取指定类型的所有编号
        
        Args:
            ref_type: 'fig', 'tbl', 'eq'
        
        Returns:
            List[int]: 编号列表
        """
        nums = []
        for name, (r_type, num) in self.bookmarks.items():
            if r_type == ref_type:
                nums.append(num)
        return sorted(set(nums))
    
    def get_current_bookmarks(self) -> Dict[str, Tuple[str, int]]:
        """获取当前书签映射"""
        return self.bookmarks.copy()
