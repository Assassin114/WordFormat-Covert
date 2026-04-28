"""
OOXML 底层操作工具
提供 python-docx 难以直接实现的底层 XML 操作
"""

from lxml import etree
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Pt, Twips


# WordprocessingML 命名空间
WML_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def get_paragraph_pPr(paragraph):
    """获取或创建段落的 pPr 元素"""
    p = paragraph._element if hasattr(paragraph, '_element') else paragraph
    pPr = p.find(WML_NS + 'pPr')
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p.insert(0, pPr)
    return pPr


def set_line_spacing(paragraph: Paragraph, line_value: float, line_rule: str = 'auto'):
    """
    设置段落行距
    
    Args:
        paragraph: 段落对象
        line_value: 行距值
            - 倍数: 1.5, 2.0 等
            - 磅值: 20, 24 等（当 line_rule='exact'）
        line_rule: 'auto'（倍数）或 'exact'（固定值）
    """
    pPr = get_paragraph_pPr(paragraph)
    spacing = pPr.find(WML_NS + 'spacing')
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    
    if line_rule == 'auto':
        # 240 twips = 1 倍行距
        line_val = str(int(line_value * 240))
        spacing.set(WML_NS + 'lineRule', 'auto')
    else:
        # 磅值转 twips
        line_val = str(int(line_value * 20))
        spacing.set(WML_NS + 'lineRule', 'exact')
    
    spacing.set(WML_NS + 'line', line_val)


def set_space_before(paragraph: Paragraph, pts: float):
    """设置段前间距（磅）"""
    pPr = get_paragraph_pPr(paragraph)
    spacing = pPr.find(WML_NS + 'spacing')
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    # 1pt = 20 twips
    spacing.set(WML_NS + 'before', str(int(pts * 20)))


def set_space_after(paragraph: Paragraph, pts: float):
    """设置段后间距（磅）"""
    pPr = get_paragraph_pPr(paragraph)
    spacing = pPr.find(WML_NS + 'spacing')
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    # 1pt = 20 twips
    spacing.set(WML_NS + 'after', str(int(pts * 20)))


def set_first_line_indent(paragraph: Paragraph, pts: float):
    """设置首行缩进（磅）"""
    pPr = get_paragraph_pPr(paragraph)
    ind = pPr.find(WML_NS + 'ind')
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    # 1pt = 20 twips
    ind.set(WML_NS + 'firstLine', str(int(pts * 20)))


def set_alignment(paragraph: Paragraph, align: str):
    """设置段落对齐
    Args:
        align: 'left', 'center', 'right', 'both' (两端对齐)
    """
    pPr = get_paragraph_pPr(paragraph)
    jc = pPr.find(WML_NS + 'jc')
    if jc is None:
        jc = OxmlElement('w:jc')
        pPr.append(jc)
    
    align_map = {
        'left': 'left',
        'center': 'center',
        'right': 'right',
        'both': 'both',
        'justify': 'both',
    }
    jc.set(WML_NS + 'val', align_map.get(align, 'left'))


def set_widow_control(paragraph: Paragraph, enable: bool = True):
    """
    设置孤行控制
    widowControl: 控制段落最后一行单独在页面顶部
    """
    pPr = get_paragraph_pPr(paragraph)
    widowControl = pPr.find(WML_NS + 'widowControl')
    if widowControl is None:
        widowControl = OxmlElement('w:widowControl')
        pPr.append(widowControl)
    widowControl.set(WML_NS + 'val', 'true' if enable else 'false')


def set_keep_with_next(paragraph: Paragraph, enable: bool = True):
    """设置与下一段同页"""
    pPr = get_paragraph_pPr(paragraph)
    keepNext = pPr.find(WML_NS + 'keepNext')
    if keepNext is None:
        keepNext = OxmlElement('w:keepNext')
        pPr.append(keepNext)
    keepNext.set(WML_NS + 'val', 'true' if enable else 'false')


def set_page_break_before(paragraph: Paragraph):
    """设置段前分页"""
    pPr = get_paragraph_pPr(paragraph)
    pageBreak = pPr.find(WML_NS + 'pageBreakBefore')
    if pageBreak is None:
        pageBreak = OxmlElement('w:pageBreakBefore')
        pPr.append(pageBreak)
    pageBreak.set(WML_NS + 'val', 'true')


def set_paragraph_properties(paragraph: Paragraph, 
                              line_spacing: float = None,
                              space_before: float = None,
                              space_after: float = None,
                              first_line_indent: float = None,
                              alignment: str = None,
                              widow_control: bool = None,
                              keep_with_next: bool = None):
    """批量设置段落属性"""
    if line_spacing is not None:
        set_line_spacing(paragraph, line_spacing)
    if space_before is not None:
        set_space_before(paragraph, space_before)
    if space_after is not None:
        set_space_after(paragraph, space_after)
    if first_line_indent is not None:
        set_first_line_indent(paragraph, first_line_indent)
    if alignment is not None:
        set_alignment(paragraph, alignment)
    if widow_control is not None:
        set_widow_control(paragraph, widow_control)
    if keep_with_next is not None:
        set_keep_with_next(paragraph, keep_with_next)


def insert_page_break(paragraph: Paragraph):
    """在段落后插入分页符"""
    p = paragraph._element
    # 创建 <w:br> 元素，type="page"
    br = OxmlElement('w:br')
    br.set(WML_NS + 'type', 'page')
    
    # 创建新的 run 包含分页符
    new_run = OxmlElement('w:r')
    new_run.append(br)
    p.append(new_run)


def set_run_font(run, font_name: str, font_size: float = None, 
                 bold: bool = None, italic: bool = None, color: str = None):
    """设置 run 的字体属性"""
    # 字体名称
    rPr = run._element.find(WML_NS + 'rPr')
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        run._element.insert(0, rPr)
    
    # 字体名
    rFonts = rPr.find(WML_NS + 'rFonts')
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(WML_NS + 'ascii', font_name)
    rFonts.set(WML_NS + 'hAnsi', font_name)
    rFonts.set(WML_NS + 'cs', font_name)
    rFonts.set(WML_NS + 'eastAsia', font_name)
    
    # 字号
    if font_size is not None:
        sz = rPr.find(WML_NS + 'sz')
        if sz is None:
            sz = OxmlElement('w:sz')
            rPr.append(sz)
        # 字号单位是半点（half-point），所以要 * 2
        sz.set(WML_NS + 'val', str(int(font_size * 2)))
        
        szCs = rPr.find(WML_NS + 'szCs')
        if szCs is None:
            szCs = OxmlElement('w:szCs')
            rPr.append(szCs)
        szCs.set(WML_NS + 'val', str(int(font_size * 2)))
    
    # 加粗
    if bold is not None:
        b = rPr.find(WML_NS + 'b')
        if bold:
            if b is None:
                b = OxmlElement('w:b')
                rPr.append(b)
        else:
            if b is not None:
                rPr.remove(b)
    
    # 斜体
    if italic is not None:
        i = rPr.find(WML_NS + 'i')
        if italic:
            if i is None:
                i = OxmlElement('w:i')
                rPr.append(i)
        else:
            if i is not None:
                rPr.remove(i)
    
    # 颜色
    if color is not None and color != '#000000':
        color_elem = rPr.find(WML_NS + 'color')
        if color_elem is None:
            color_elem = OxmlElement('w:color')
            rPr.append(color_elem)
        # 去掉 # 前缀
        color_val = color.lstrip('#')
        color_elem.set(WML_NS + 'val', color_val)


def get_section_page_orientation(section) -> tuple:
    """
    获取节的页面方向
    
    Returns:
        tuple: (is_landscape: bool, width: int, height: int)
    """
    pPr = section._element.find('.//' + WML_NS + 'pgSz')
    if pPr is None:
        return False, 0, 0
    
    orient = pPr.get(WML_NS + 'orient', 'portrait')
    w = int(pPr.get(WML_NS + 'w', 0))
    h = int(pPr.get(WML_NS + 'h', 0))
    
    is_landscape = (orient == 'landscape') or (w > h)
    return is_landscape, w, h


def is_landscape_section(section) -> bool:
    """判断节是否为横向"""
    is_landscape, _, _ = get_section_page_orientation(section)
    return is_landscape


def add_blank_page_for_duplex(doc):
    """
    为双面打印添加空白页
    检查文档最后一页是否为奇数，如果是则添加空白页
    """
    # Word 中 section 的设置控制前一页的奇偶性
    # 简化处理：在最后一个段落后添加分页符
    pass  # TODO: 实现
