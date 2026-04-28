"""
models - 数据模型
"""

from .template_config import (
    TemplateConfig,
    FontConfig,
    ParagraphConfig,
    HeadingConfig,
    CaptionConfig,
    HeaderFooterConfig,
    CoverConfig,
    SignatureConfig,
    RevisionConfig,
    PrintMode,
    NumberingMode,
    FormatResult,
    BatchResult,
    create_default_template,
)

__all__ = [
    'TemplateConfig',
    'FontConfig',
    'ParagraphConfig',
    'HeadingConfig',
    'CaptionConfig',
    'HeaderFooterConfig',
    'CoverConfig',
    'SignatureConfig',
    'RevisionConfig',
    'PrintMode',
    'NumberingMode',
    'FormatResult',
    'BatchResult',
    'create_default_template',
]
