"""
utils - 工具模块
"""

from .ooxml_utils import (
    set_line_spacing,
    set_space_before,
    set_space_after,
    set_first_line_indent,
    set_alignment,
    set_widow_control,
    set_keep_with_next,
    set_page_break_before,
    set_paragraph_properties,
    insert_page_break,
    set_run_font,
    get_section_page_orientation,
    is_landscape_section,
    add_blank_page_for_duplex,
    get_paragraph_pPr,
)

from .file_utils import (
    scan_docx_files,
    get_output_path,
    ensure_dir,
    get_file_size,
    format_file_size,
    is_valid_docx,
)

__all__ = [
    # ooxml_utils
    'set_line_spacing',
    'set_space_before',
    'set_space_after',
    'set_first_line_indent',
    'set_alignment',
    'set_widow_control',
    'set_keep_with_next',
    'set_page_break_before',
    'set_paragraph_properties',
    'insert_page_break',
    'set_run_font',
    'get_section_page_orientation',
    'is_landscape_section',
    'add_blank_page_for_duplex',
    'get_paragraph_pPr',
    # file_utils
    'scan_docx_files',
    'get_output_path',
    'ensure_dir',
    'get_file_size',
    'format_file_size',
    'is_valid_docx',
]
