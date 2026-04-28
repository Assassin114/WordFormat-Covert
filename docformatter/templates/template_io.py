"""
模板加载和保存模块
"""

import json
from pathlib import Path
from typing import Dict, Any

from ..models.template_config import (
    TemplateConfig, FontConfig, ParagraphConfig, HeadingConfig,
    CaptionConfig, HeaderFooterConfig, CoverConfig, SignatureConfig,
    RevisionConfig, PrintMode, NumberingMode, create_default_template
)


def template_to_dict(config: TemplateConfig) -> Dict[str, Any]:
    """TemplateConfig 转字典"""
    return {
        'version': config.version,
        'print_mode': config.print_mode.value,
        'cover': {
            'enabled': config.cover.enabled,
            'template_path': config.cover.template_path,
            'fields': config.cover.fields,
        },
        'signature': {
            'enabled': config.signature.enabled,
            'template_path': config.signature.template_path,
            'fields': config.signature.fields,
        },
        'revision': {
            'enabled': config.revision.enabled,
        },
        'headings': [
            {
                'level': h.level,
                'font': {
                    'name': h.font.name,
                    'size': h.font.size,
                    'bold': h.font.bold,
                    'italic': h.font.italic,
                    'color': h.font.color,
                },
                'paragraph': {
                    'line_spacing': h.paragraph.line_spacing,
                    'space_before': h.paragraph.space_before,
                    'space_after': h.paragraph.space_after,
                    'first_line_indent': h.paragraph.first_line_indent,
                    'alignment': h.paragraph.alignment,
                },
                'outline_level': h.outline_level,
            }
            for h in config.headings
        ],
        'body': {
            'font': {
                'name': config.body.font.name,
                'size': config.body.font.size,
                'bold': config.body.font.bold,
                'italic': config.body.font.italic,
                'color': config.body.font.color,
            },
            'line_spacing': config.body.line_spacing,
            'space_before': config.body.space_before,
            'space_after': config.body.space_after,
            'first_line_indent': config.body.first_line_indent,
            'alignment': config.body.alignment,
        },
        'caption': {
            'prefix_figure': config.caption.prefix_figure,
            'prefix_table': config.caption.prefix_table,
            'prefix_equation': config.caption.prefix_equation,
            'numbering_mode': config.caption.numbering_mode.value,
            'position_figure': config.caption.position_figure,
            'position_table': config.caption.position_table,
        },
        'header_footer': {
            'show_on_first': config.header_footer.show_on_first,
            'different_odd_even': config.header_footer.different_odd_even,
            'page_number_format': config.header_footer.page_number_format,
        },
    }


def dict_to_template(data: Dict[str, Any]) -> TemplateConfig:
    """字典转 TemplateConfig"""
    template = create_default_template()
    template.version = data.get('version', '1.0')
    
    # Print mode
    if 'print_mode' in data:
        template.print_mode = PrintMode(data['print_mode'])
    
    # Cover
    if 'cover' in data:
        template.cover.enabled = data['cover'].get('enabled', True)
        template.cover.template_path = data['cover'].get('template_path')
        template.cover.fields = data['cover'].get('fields', {})
    
    # Signature
    if 'signature' in data:
        template.signature.enabled = data['signature'].get('enabled', True)
        template.signature.template_path = data['signature'].get('template_path')
        template.signature.fields = data['signature'].get('fields', {})
    
    # Revision
    if 'revision' in data:
        template.revision.enabled = data['revision'].get('enabled', True)
    
    # Headings
    if 'headings' in data:
        for i, h_data in enumerate(data['headings']):
            if i < len(template.headings):
                h = template.headings[i]
                h.level = h_data.get('level', i + 1)
                if 'font' in h_data:
                    h.font.name = h_data['font'].get('name', h.font.name)
                    h.font.size = h_data['font'].get('size', h.font.size)
                    h.font.bold = h_data['font'].get('bold', h.font.bold)
                    h.font.italic = h_data['font'].get('italic', h.font.italic)
                    h.font.color = h_data['font'].get('color', h.font.color)
                if 'paragraph' in h_data:
                    h.paragraph.line_spacing = h_data['paragraph'].get('line_spacing', h.paragraph.line_spacing)
                    h.paragraph.space_before = h_data['paragraph'].get('space_before', h.paragraph.space_before)
                    h.paragraph.space_after = h_data['paragraph'].get('space_after', h.paragraph.space_after)
                    h.paragraph.first_line_indent = h_data['paragraph'].get('first_line_indent', h.paragraph.first_line_indent)
                    h.paragraph.alignment = h_data['paragraph'].get('alignment', h.paragraph.alignment)
    
    # Body
    if 'body' in data:
        if 'font' in data['body']:
            template.body.font.name = data['body']['font'].get('name', template.body.font.name)
            template.body.font.size = data['body']['font'].get('size', template.body.font.size)
            template.body.font.bold = data['body']['font'].get('bold', template.body.font.bold)
            template.body.font.italic = data['body']['font'].get('italic', template.body.font.italic)
            template.body.font.color = data['body']['font'].get('color', template.body.font.color)
        template.body.line_spacing = data['body'].get('line_spacing', template.body.line_spacing)
        template.body.space_before = data['body'].get('space_before', template.body.space_before)
        template.body.space_after = data['body'].get('space_after', template.body.space_after)
        template.body.first_line_indent = data['body'].get('first_line_indent', template.body.first_line_indent)
        template.body.alignment = data['body'].get('alignment', template.body.alignment)
    
    # Caption
    if 'caption' in data:
        template.caption.prefix_figure = data['caption'].get('prefix_figure', template.caption.prefix_figure)
        template.caption.prefix_table = data['caption'].get('prefix_table', template.caption.prefix_table)
        template.caption.prefix_equation = data['caption'].get('prefix_equation', template.caption.prefix_equation)
        if 'numbering_mode' in data['caption']:
            template.caption.numbering_mode = NumberingMode(data['caption']['numbering_mode'])
        template.caption.position_figure = data['caption'].get('position_figure', template.caption.position_figure)
        template.caption.position_table = data['caption'].get('position_table', template.caption.position_table)
    
    # Header/Footer
    if 'header_footer' in data:
        template.header_footer.show_on_first = data['header_footer'].get('show_on_first', template.header_footer.show_on_first)
        template.header_footer.different_odd_even = data['header_footer'].get('different_odd_even', template.header_footer.different_odd_even)
        template.header_footer.page_number_format = data['header_footer'].get('page_number_format', template.header_footer.page_number_format)
    
    return template


def save_template(config: TemplateConfig, path: str):
    """保存模板配置到JSON文件"""
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(template_to_dict(config), f, ensure_ascii=False, indent=2)


def load_template(path: str) -> TemplateConfig:
    """从JSON文件加载模板配置"""
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return dict_to_template(data)


def save_default_template(path: str = None):
    """保存默认模板"""
    template = create_default_template()
    if path is None:
        path = Path(__file__).parent / 'default_template.json'
    save_template(template, str(path))
    return str(path)
