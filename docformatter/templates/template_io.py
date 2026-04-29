"""
模板加载和保存模块
完整序列化/反序列化 TemplateConfig
"""

import json
from pathlib import Path
from typing import Dict, Any

from ..models.template_config import (
    TemplateConfig, FontConfig, ParagraphConfig, BodyConfig,
    HeadingConfig, CaptionConfig, HeaderFooterConfig, HFPageConfig,
    CoverConfig, SignatureConfig, RevisionConfig, TOCConfig,
    TableStyleConfig, TableFontConfig, CodeBlockConfig, PageConfig,
    PrintMode, NumberingMode, EquationFormat,
    PageNumberFormat, PageNumberPosition, PageNumberAlignment,
    create_default_template,
)


# ══════════════════════════════════════════════════════
# 序列化辅助
# ══════════════════════════════════════════════════════

def _font_to_dict(f: FontConfig) -> dict:
    return {
        "cn_name": f.cn_name,
        "en_name": f.en_name,
        "size": f.size,
        "size_name": f.size_name,
        "bold": f.bold,
        "italic": f.italic,
        "color": f.color,
    }


def _font_from_dict(d: dict, default: FontConfig = None) -> FontConfig:
    if default is None:
        default = FontConfig()
    if not d:
        return default
    # 向后兼容旧模板的 'name' 字段
    if "name" in d and "cn_name" not in d:
        default.cn_name = d["name"]
    return FontConfig(
        cn_name=d.get("cn_name", default.cn_name),
        en_name=d.get("en_name", default.en_name),
        size=d.get("size", default.size),
        size_name=d.get("size_name", default.size_name),
        bold=d.get("bold", default.bold),
        italic=d.get("italic", default.italic),
        color=d.get("color", default.color),
    )


def _para_to_dict(p: ParagraphConfig) -> dict:
    return {
        "line_spacing": p.line_spacing,
        "line_spacing_type": p.line_spacing_type,
        "line_spacing_fixed": p.line_spacing_fixed,
        "line_spacing_min": p.line_spacing_min,
        "space_before": p.space_before,
        "space_after": p.space_after,
        "first_line_indent": p.first_line_indent,
        "alignment": p.alignment,
    }


def _para_from_dict(d: dict, default: ParagraphConfig = None) -> ParagraphConfig:
    if default is None:
        default = ParagraphConfig()
    if not d:
        return default
    return ParagraphConfig(
        line_spacing=d.get("line_spacing", default.line_spacing),
        line_spacing_type=d.get("line_spacing_type", default.line_spacing_type),
        line_spacing_fixed=d.get("line_spacing_fixed", default.line_spacing_fixed),
        line_spacing_min=d.get("line_spacing_min", default.line_spacing_min),
        space_before=d.get("space_before", default.space_before),
        space_after=d.get("space_after", default.space_after),
        first_line_indent=d.get("first_line_indent", default.first_line_indent),
        alignment=d.get("alignment", default.alignment),
    )


def _page_number_font_to_dict(pnf: FontConfig) -> dict:
    return _font_to_dict(pnf)


def _page_number_font_from_dict(d: dict, default: FontConfig) -> FontConfig:
    return _font_from_dict(d, default)


# ══════════════════════════════════════════════════════
# 模板 → JSON
# ══════════════════════════════════════════════════════

def template_to_dict(config: TemplateConfig) -> Dict[str, Any]:
    return {
        "version": config.version,
        "print_mode": config.print_mode.value,
        "equation_format": config.equation_format.value,

        "page": {
            "width": config.page.width,
            "height": config.page.height,
            "orientation": config.page.orientation,
            "margin_top": config.page.margin_top,
            "margin_bottom": config.page.margin_bottom,
            "margin_left": config.page.margin_left,
            "margin_right": config.page.margin_right,
            "gutter": config.page.gutter,
            "gutter_position": config.page.gutter_position,
            "page_number_start": config.page.page_number_start,
        },

        "cover": {
            "enabled": config.cover.enabled,
            "template_path": config.cover.template_path,
            "fields": config.cover.fields,
            "heading_font": _font_to_dict(config.cover.heading_font),
            "heading_paragraph": _para_to_dict(config.cover.heading_paragraph),
        },

        "signature": {
            "enabled": config.signature.enabled,
            "template_path": config.signature.template_path,
            "title": config.signature.title,
            "heading_font": _font_to_dict(config.signature.heading_font),
            "heading_paragraph": _para_to_dict(config.signature.heading_paragraph),
        },

        "revision": {
            "enabled": config.revision.enabled,
            "heading_font": _font_to_dict(config.revision.heading_font),
            "heading_paragraph": _para_to_dict(config.revision.heading_paragraph),
            "table_header_font": _font_to_dict(config.revision.table_header_font),
            "table_header_paragraph": _para_to_dict(config.revision.table_header_paragraph),
            "table_body_font": _font_to_dict(config.revision.table_body_font),
            "table_body_paragraph": _para_to_dict(config.revision.table_body_paragraph),
        },

        "toc": {
            "enabled": config.toc.enabled,
            "heading_font": _font_to_dict(config.toc.heading_font),
            "heading_paragraph": _para_to_dict(config.toc.heading_paragraph),
            "entry_font": _font_to_dict(config.toc.entry_font),
            "entry_paragraph": _para_to_dict(config.toc.entry_paragraph),
            "entry_indent": config.toc.entry_indent,
        },

        "headings": [
            {
                "level": h.level,
                "number_style": h.number_style,
                "number_multi": h.number_multi,
                "font": _font_to_dict(h.font),
                "paragraph": _para_to_dict(h.paragraph),
            }
            for h in config.headings
        ],

        "body": {
            "font": _font_to_dict(config.body.font),
            "paragraph": _para_to_dict(config.body.paragraph),
        },

        "table_font": {
            "regular_table": _table_style_to_dict(config.table_font.regular_table),
            "record_table": _table_style_to_dict(config.table_font.record_table),
        },

        "caption": {
            "figure_prefix": config.caption.figure_prefix,
            "table_prefix": config.caption.table_prefix,
            "equation_prefix": config.caption.equation_prefix,
            "position_figure": config.caption.position_figure,
            "position_table": config.caption.position_table,
            "numbering_mode": config.caption.numbering_mode.value,
            "font": _font_to_dict(config.caption.font),
            "paragraph": _para_to_dict(config.caption.paragraph),
        },

        "code_block": {
            "font": _font_to_dict(config.code_block.font),
            "paragraph": _para_to_dict(config.code_block.paragraph),
            "bg_color": config.code_block.bg_color,
        },

        "header_footer": {
            "cover_page": _hf_page_to_dict(config.header_footer.cover_page),
            "toc_page": _hf_page_to_dict(config.header_footer.toc_page),
            "body_page": _hf_page_to_dict(config.header_footer.body_page),
            "different_first_page": config.header_footer.different_first_page,
            "different_odd_even": config.header_footer.different_odd_even,
        },
    }


def _table_style_to_dict(s: TableStyleConfig) -> dict:
    return {
        "header_font": _font_to_dict(s.header_font),
        "header_paragraph": _para_to_dict(s.header_paragraph),
        "header_bg_color": s.header_bg_color,
        "body_font": _font_to_dict(s.body_font),
        "body_paragraph": _para_to_dict(s.body_paragraph),
    }


def _hf_page_to_dict(p: HFPageConfig) -> dict:
    return {
        "header_text": p.header_text,
        "show_header": p.show_header,
        "show_footer": p.show_footer,
        "show_page_number": p.show_page_number,
        "page_number_format": p.page_number_format.value,
        "page_number_position": p.page_number_position.value,
        "page_number_alignment": p.page_number_alignment.value,
        "page_number_font": _font_to_dict(p.page_number_font),
        "page_number_paragraph": _para_to_dict(p.page_number_paragraph),
    }


# ══════════════════════════════════════════════════════
# JSON → 模板
# ══════════════════════════════════════════════════════

def dict_to_template(data: Dict[str, Any]) -> TemplateConfig:
    default = create_default_template()

    default.version = data.get("version", "1.0")

    if "print_mode" in data:
        try:
            default.print_mode = PrintMode(data["print_mode"])
        except (ValueError, KeyError):
            pass

    if "equation_format" in data:
        try:
            default.equation_format = EquationFormat(data["equation_format"])
        except (ValueError, KeyError):
            pass

    # 页面设置
    if "page" in data:
        p = data["page"]
        pc = default.page
        pc.width = p.get("width", pc.width)
        pc.height = p.get("height", pc.height)
        pc.orientation = p.get("orientation", pc.orientation)
        pc.margin_top = p.get("margin_top", pc.margin_top)
        pc.margin_bottom = p.get("margin_bottom", pc.margin_bottom)
        pc.margin_left = p.get("margin_left", pc.margin_left)
        pc.margin_right = p.get("margin_right", pc.margin_right)
        pc.gutter = p.get("gutter", pc.gutter)
        pc.gutter_position = p.get("gutter_position", pc.gutter_position)
        pc.page_number_start = p.get("page_number_start", pc.page_number_start)

    # 封面
    if "cover" in data:
        c = data["cover"]
        default.cover.enabled = c.get("enabled", True)
        default.cover.template_path = c.get("template_path")
        default.cover.fields = c.get("fields", {})
        if "heading_font" in c:
            default.cover.heading_font = _font_from_dict(c["heading_font"], default.cover.heading_font)
        if "heading_paragraph" in c:
            default.cover.heading_paragraph = _para_from_dict(c["heading_paragraph"], default.cover.heading_paragraph)

    # 签署页
    if "signature" in data:
        s = data["signature"]
        default.signature.enabled = s.get("enabled", True)
        default.signature.template_path = s.get("template_path")
        default.signature.title = s.get("title", default.signature.title)
        if "heading_font" in s:
            default.signature.heading_font = _font_from_dict(s["heading_font"], default.signature.heading_font)
        if "heading_paragraph" in s:
            default.signature.heading_paragraph = _para_from_dict(s["heading_paragraph"], default.signature.heading_paragraph)

    # 修改记录
    if "revision" in data:
        r = data["revision"]
        default.revision.enabled = r.get("enabled", True)
        if "heading_font" in r:
            default.revision.heading_font = _font_from_dict(r["heading_font"], default.revision.heading_font)
        if "heading_paragraph" in r:
            default.revision.heading_paragraph = _para_from_dict(r["heading_paragraph"], default.revision.heading_paragraph)
        if "table_header_font" in r:
            default.revision.table_header_font = _font_from_dict(r["table_header_font"], default.revision.table_header_font)
        if "table_header_paragraph" in r:
            default.revision.table_header_paragraph = _para_from_dict(r["table_header_paragraph"], default.revision.table_header_paragraph)
        if "table_body_font" in r:
            default.revision.table_body_font = _font_from_dict(r["table_body_font"], default.revision.table_body_font)
        if "table_body_paragraph" in r:
            default.revision.table_body_paragraph = _para_from_dict(r["table_body_paragraph"], default.revision.table_body_paragraph)

    # 目录
    if "toc" in data:
        t = data["toc"]
        default.toc.enabled = t.get("enabled", True)
        if "heading_font" in t:
            default.toc.heading_font = _font_from_dict(t["heading_font"], default.toc.heading_font)
        if "heading_paragraph" in t:
            default.toc.heading_paragraph = _para_from_dict(t["heading_paragraph"], default.toc.heading_paragraph)
        if "entry_font" in t:
            default.toc.entry_font = _font_from_dict(t["entry_font"], default.toc.entry_font)
        if "entry_paragraph" in t:
            default.toc.entry_paragraph = _para_from_dict(t["entry_paragraph"], default.toc.entry_paragraph)
        default.toc.entry_indent = t.get("entry_indent", default.toc.entry_indent)

    # 标题
    if "headings" in data:
        for i, hd in enumerate(data["headings"]):
            if i < len(default.headings):
                h = default.headings[i]
                h.level = hd.get("level", i + 1)
                h.number_style = hd.get("number_style", h.number_style)
                h.number_multi = hd.get("number_multi", h.number_multi)
                if "font" in hd:
                    h.font = _font_from_dict(hd["font"], h.font)
                if "paragraph" in hd:
                    h.paragraph = _para_from_dict(hd["paragraph"], h.paragraph)

    # 正文
    if "body" in data:
        b = data["body"]
        if "font" in b:
            default.body.font = _font_from_dict(b["font"], default.body.font)
        if "paragraph" in b:
            default.body.paragraph = _para_from_dict(b["paragraph"], default.body.paragraph)

    # 表格
    if "table_font" in data:
        tf = data["table_font"]
        if "regular_table" in tf:
            default.table_font.regular_table = _table_style_from_dict(tf["regular_table"], default.table_font.regular_table)
        if "record_table" in tf:
            default.table_font.record_table = _table_style_from_dict(tf["record_table"], default.table_font.record_table)

    # 题注
    if "caption" in data:
        cap = data["caption"]
        default.caption.figure_prefix = cap.get("figure_prefix", default.caption.figure_prefix)
        default.caption.table_prefix = cap.get("table_prefix", default.caption.table_prefix)
        default.caption.equation_prefix = cap.get("equation_prefix", default.caption.equation_prefix)
        default.caption.position_figure = cap.get("position_figure", default.caption.position_figure)
        default.caption.position_table = cap.get("position_table", default.caption.position_table)
        if "numbering_mode" in cap:
            try:
                default.caption.numbering_mode = NumberingMode(cap["numbering_mode"])
            except (ValueError, KeyError):
                pass
        if "font" in cap:
            default.caption.font = _font_from_dict(cap["font"], default.caption.font)
        if "paragraph" in cap:
            default.caption.paragraph = _para_from_dict(cap["paragraph"], default.caption.paragraph)

    # 代码块
    if "code_block" in data:
        cb = data["code_block"]
        if "font" in cb:
            default.code_block.font = _font_from_dict(cb["font"], default.code_block.font)
        if "paragraph" in cb:
            default.code_block.paragraph = _para_from_dict(cb["paragraph"], default.code_block.paragraph)
        default.code_block.bg_color = cb.get("bg_color", default.code_block.bg_color)

    # 页眉页脚
    if "header_footer" in data:
        hf = data["header_footer"]
        if "cover_page" in hf:
            default.header_footer.cover_page = _hf_page_from_dict(hf["cover_page"], default.header_footer.cover_page)
        if "toc_page" in hf:
            default.header_footer.toc_page = _hf_page_from_dict(hf["toc_page"], default.header_footer.toc_page)
        if "body_page" in hf:
            default.header_footer.body_page = _hf_page_from_dict(hf["body_page"], default.header_footer.body_page)
        default.header_footer.different_first_page = hf.get("different_first_page", default.header_footer.different_first_page)
        default.header_footer.different_odd_even = hf.get("different_odd_even", default.header_footer.different_odd_even)

    return default


def _table_style_from_dict(d: dict, default: TableStyleConfig) -> TableStyleConfig:
    if not d:
        return default
    if "header_font" in d:
        default.header_font = _font_from_dict(d["header_font"], default.header_font)
    if "header_paragraph" in d:
        default.header_paragraph = _para_from_dict(d["header_paragraph"], default.header_paragraph)
    default.header_bg_color = d.get("header_bg_color", default.header_bg_color)
    if "body_font" in d:
        default.body_font = _font_from_dict(d["body_font"], default.body_font)
    if "body_paragraph" in d:
        default.body_paragraph = _para_from_dict(d["body_paragraph"], default.body_paragraph)
    return default


def _hf_page_from_dict(d: dict, default: HFPageConfig) -> HFPageConfig:
    if not d:
        return default
    default.header_text = d.get("header_text", default.header_text)
    default.show_header = d.get("show_header", default.show_header)
    default.show_footer = d.get("show_footer", default.show_footer)
    default.show_page_number = d.get("show_page_number", default.show_page_number)
    if "page_number_format" in d:
        try:
            default.page_number_format = PageNumberFormat(d["page_number_format"])
        except (ValueError, KeyError):
            pass
    if "page_number_position" in d:
        try:
            default.page_number_position = PageNumberPosition(d["page_number_position"])
        except (ValueError, KeyError):
            pass
    if "page_number_alignment" in d:
        try:
            default.page_number_alignment = PageNumberAlignment(d["page_number_alignment"])
        except (ValueError, KeyError):
            pass
    if "page_number_font" in d:
        default.page_number_font = _font_from_dict(d["page_number_font"], default.page_number_font)
    if "page_number_paragraph" in d:
        default.page_number_paragraph = _para_from_dict(d["page_number_paragraph"], default.page_number_paragraph)
    return default


# ══════════════════════════════════════════════════════
# 文件读写
# ══════════════════════════════════════════════════════

def save_template(config: TemplateConfig, path: str):
    """保存模板配置到JSON文件"""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(template_to_dict(config), f, ensure_ascii=False, indent=2)


def load_template(path: str) -> TemplateConfig:
    """从JSON文件加载模板配置"""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return dict_to_template(data)


def save_default_template(path: str = None):
    """保存默认模板"""
    template = create_default_template()
    if path is None:
        path = Path(__file__).parent / "default_template.json"
    save_template(template, str(path))
    return str(path)
