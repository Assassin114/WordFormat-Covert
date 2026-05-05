"""
文档结构分析器
分析 Word 文档的结构：分节、元素分类、标题级别推断
"""

from enum import Enum
from dataclasses import dataclass, field
from typing import List, Optional
from docx import Document
from docx.oxml.ns import qn


WML_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


class SectionType(Enum):
    COVER = "cover"
    SIGNATURE = "signature"
    REVISION = "revision"
    TOC = "toc"
    BODY = "body"


class ElementType(Enum):
    HEADING = "heading"
    PARAGRAPH = "paragraph"
    IMAGE = "image"
    TABLE = "table"
    EQUATION = "equation"


class TableCategory(Enum):
    REGULAR = "regular"
    REVISION = "revision"
    TEST_RECORD = "test_record"


class ImageKind(Enum):
    REGULAR = "regular"
    VISIO_OLE = "visio_ole"
    VISIO_EMF = "visio_emf"
    MATHTYPE = "mathtype"


class EquationKind(Enum):
    OMML = "omml"
    MATHML = "mathml"
    MATHTYPE_IMAGE = "mathtype_image"
    LATEX_TEXT = "latex_text"
    UNKNOWN = "unknown"


@dataclass
class ElementInfo:
    element_type: ElementType
    heading_level: int = 0
    text_preview: str = ""
    para_index: int = -1


@dataclass
class SectionInfo:
    section_type: SectionType
    elements: List[ElementInfo] = field(default_factory=list)


@dataclass
class DocumentStructure:
    sections: List[SectionInfo] = field(default_factory=list)


class DocumentAnalyzer:
    """文档结构分析器"""

    # 签署关键字（出现3个以上判定为签署页）
    SIGNATURE_KEYWORDS = ["拟制", "校对", "标审", "审核", "批准"]

    # 修改记录标题
    REVISION_TITLES = ["修改记录", "修订记录", "版本记录", "文档修订记录"]

    # 目录标题
    TOC_TITLES = ["目录", "目  录", "目 录", "CONTENTS", "Table of Contents"]

    # 正文起始标记
    BODY_START_MARKERS = [
        "引言", "前言", "摘要", "Abstract",
        "第一章", "第1章", "一、", "1 ", "1.",
    ]

    # 标题级字号范围（pt）
    HEADING_SIZE_RANGES = [
        (16, 99),   # 1级：≥16pt
        (14, 15.9),  # 2级
        (12, 13.9),  # 3级
        (11, 11.9),  # 4级
        (9, 10.9),   # 5级
    ]

    # 标题编号正则
    HEADING_NUM_PATTERNS = [
        (r'^第[一二三四五六七八九十]+章', 1),
        (r'^[一二三四五六七八九十]+、', 1),
        (r'^\([一二三四五六七八九十]+\)', 2),
        (r'^\d+\.', 3),
        (r'^\(\d+\)', 4),
        (r'^[①②③④⑤⑥⑦⑧⑨⑩]', 5),
    ]

    def __init__(self):
        pass

    def analyze(self, docx_path: str) -> DocumentStructure:
        """分析文档结构"""
        doc = Document(docx_path)
        structure = DocumentStructure()

        paragraphs = list(doc.paragraphs)
        if not paragraphs:
            return structure

        # 按分节符切分
        section_ranges = self._split_by_sections(paragraphs)

        # 逐节分析
        for start, end in section_ranges:
            section_paras = paragraphs[start:end]
            sect_type = self._classify_section(section_paras)
            section_info = SectionInfo(section_type=sect_type)

            for i, para in enumerate(section_paras):
                actual_idx = start + i
                elem = self._classify_element(para, actual_idx, sect_type)
                section_info.elements.append(elem)

            structure.sections.append(section_info)

        return structure

    def _split_by_sections(self, paragraphs) -> List[tuple]:
        """按分节符切分段落"""
        ranges = []
        current_start = 0

        for i, para in enumerate(paragraphs):
            pPr = para._element.find(WML_NS + 'pPr')
            if pPr is not None:
                sect_pr = pPr.find(WML_NS + 'sectPr')
                if sect_pr is not None and i > 0:
                    ranges.append((current_start, i + 1))
                    current_start = i + 1

        # 最后一个区间
        if current_start < len(paragraphs):
            ranges.append((current_start, len(paragraphs)))

        return ranges or [(0, len(paragraphs))]

    def _classify_section(self, paragraphs) -> SectionType:
        """根据节内容判断节类型"""
        texts = [p.text.strip() for p in paragraphs if p.text.strip()]

        if not texts:
            return SectionType.BODY

        first_text = texts[0] if texts else ""
        combined = " ".join(texts[:5])

        # 签名页检测
        sig_count = sum(1 for kw in self.SIGNATURE_KEYWORDS if kw in combined)
        if sig_count >= 3:
            return SectionType.SIGNATURE

        # 修改记录
        for title in self.REVISION_TITLES:
            if title in first_text or title in combined:
                return SectionType.REVISION

        # 目录
        for title in self.TOC_TITLES:
            if title in first_text:
                return SectionType.TOC
        # 目录也可能只有TOC域，无具体标题文字
        if any("TOC" in p._element.xml for p in paragraphs):
            return SectionType.TOC

        # 正文标记
        for marker in self.BODY_START_MARKERS:
            if any(marker in t for t in texts[:3]):
                return SectionType.BODY

        # 兜底：第一个节默认为封面区，其余为正文
        return SectionType.COVER

    def _classify_element(self, para, para_index: int, sect_type: SectionType) -> ElementInfo:
        """分类单个段落"""
        text = para.text.strip()

        # 1. 图片检测（含Visio）
        img_kind = self._detect_image(para)
        if img_kind is not None:
            return ElementInfo(
                element_type=ElementType.IMAGE,
                text_preview=text[:50] or f"[{img_kind.value}]",
                para_index=para_index,
            )

        # 2. 公式检测
        if self._is_equation(para):
            eq_kind = self._detect_equation_kind(para)
            return ElementInfo(
                element_type=ElementType.EQUATION,
                text_preview=text[:50] or f"[{eq_kind.value}]",
                para_index=para_index,
            )

        # 3. 标题检测（仅正文区做）
        if sect_type == SectionType.BODY:
            level = self._detect_heading_level(para)
            if level > 0:
                return ElementInfo(
                    element_type=ElementType.HEADING,
                    heading_level=level,
                    text_preview=text[:80],
                    para_index=para_index,
                )

        # 4. 普通段落
        return ElementInfo(
            element_type=ElementType.PARAGRAPH,
            text_preview=text[:80],
            para_index=para_index,
        )

    def _detect_heading_level(self, para) -> int:
        """推断标题级别，0 表示非标题"""

        # 优先级1: style name
        style_name = para.style.name if para.style else ""
        if style_name:
            style_lower = style_name.lower()
            for lv in range(1, 6):
                if f'heading {lv}' in style_lower or f'标题 {lv}' in style_lower:
                    return lv

        # 优先级2: outlineLvl
        pPr = para._element.find(WML_NS + 'pPr')
        if pPr is not None:
            outline = pPr.find(WML_NS + 'outlineLvl')
            if outline is not None:
                try:
                    lv = int(outline.get(WML_NS + 'val', '0')) + 1
                    if 1 <= lv <= 5:
                        return lv
                except (ValueError, TypeError):
                    pass

        # 优先级3: 字号+加粗
        first_run_size = None
        is_bold = False
        for run in para.runs:
            if run.font.size:
                first_run_size = run.font.size.pt
                is_bold = run.font.bold
                break

        if first_run_size and is_bold:
            for lv, (lo, hi) in enumerate(self.HEADING_SIZE_RANGES, start=1):
                if lo <= first_run_size <= hi:
                    return lv

        # 优先级4: 编号模式
        text = para.text.strip()
        if text:
            for pattern, lv in self.HEADING_NUM_PATTERNS:
                import re
                if re.match(pattern, text):
                    return lv

        return 0

    # ==================== 图片检测 ====================

    def _detect_image(self, para) -> Optional[ImageKind]:
        """检测段落中的图片类型"""

        # Visio OLE
        if self._has_visio_ole(para):
            return ImageKind.VISIO_OLE

        # Visio EMF/WMF 静态图
        if self._has_visio_emf(para):
            return ImageKind.VISIO_EMF

        # MathType 图片
        if self._has_mathtype_image(para):
            return ImageKind.MATHTYPE

        # 普通图片
        if self._has_drawing(para):
            return ImageKind.REGULAR

        return None

    def _has_drawing(self, para) -> bool:
        """检测是否有 drawing 元素"""
        if para._element.findall('.//' + WML_NS + 'drawing'):
            return True
        inline_ns = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline'
        anchor_ns = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor'
        return bool(para._element.findall('.//' + inline_ns) or para._element.findall('.//' + anchor_ns))

    def _has_visio_ole(self, para) -> bool:
        """检测 Visio OLE 嵌入对象"""
        objects = para._element.findall('.//' + WML_NS + 'object')
        for obj in objects:
            ole = obj.find('{urn:schemas-microsoft-com:office:office}OLEObject')
            if ole is not None:
                progid = ole.get('ProgID', '')
                if 'Visio' in progid:
                    return True
        return False

    def _has_visio_emf(self, para) -> bool:
        """检测 Visio 导出的 EMF/WMF 静态图片（在 AlternateContent 中）"""
        import re
        alt_contents = para._element.findall(
            './/{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
        )
        for alt in alt_contents:
            # Visio EMF 可被 Choice/Fallback 包裹
            alt_xml = alt.xml
            if 'visio' in alt_xml.lower() or 'emf' in alt_xml.lower() or 'wmf' in alt_xml.lower():
                return True
        return False

    def _has_mathtype_image(self, para) -> bool:
        """检测 MathType 图片"""
        alt_contents = para._element.findall(
            './/{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent'
        )
        return len(alt_contents) > 0

    # ==================== 公式检测 ====================

    def _is_equation(self, para) -> bool:
        """检测段落是否包含公式"""
        return self._detect_equation_kind(para) != EquationKind.UNKNOWN

    def _detect_equation_kind(self, para) -> EquationKind:
        """检测公式类型"""

        # OMML 原生公式
        omaths = para._element.findall('.//' + WML_NS + 'oMath')
        if omaths:
            return EquationKind.OMML

        # 通过 run 中的 oMath 引用
        for run in para.runs:
            try:
                xml = run._element.xml
                if WML_NS + 'oMath' in xml:
                    return EquationKind.OMML
            except Exception:
                pass

        # MathML
        mathml_ns = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
        if para._element.findall('.//' + mathml_ns + 'oMath'):
            return EquationKind.MATHML

        # MathType 图片
        if self._has_mathtype_image(para):
            return EquationKind.MATHTYPE_IMAGE

        # LaTeX 文本模式 ($$...$$ 或 $...$)
        text = para.text.strip()
        if text.startswith('$$') or (text.startswith('$') and '$' in text[1:]):
            return EquationKind.LATEX_TEXT

        return EquationKind.UNKNOWN

    # ==================== 表格分类 ====================

    def classify_table(self, table) -> TableCategory:
        """判断表格类型"""
        header_texts = []
        for cell in table.rows[0].cells:
            header_texts.append(cell.text.strip().lower())

        header_combined = " ".join(header_texts)

        # 修改记录表
        revision_markers = ["序号", "版本", "日期", "修改条款", "作者", "修改内容"]
        if sum(1 for m in revision_markers if m in header_combined) >= 3:
            return TableCategory.REVISION

        # 测试记录表
        test_markers = ["测试用例", "测试内容", "测试方法", "测试结果", "测试人员"]
        if sum(1 for m in test_markers if m in header_combined) >= 2:
            return TableCategory.TEST_RECORD

        # 试验记录表
        trial_markers = ["试验项目", "试验内容", "试验结果", "试验时间", "试验人员"]
        if sum(1 for m in trial_markers if m in header_combined) >= 2:
            return TableCategory.TEST_RECORD

        return TableCategory.REGULAR
