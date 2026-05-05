"""
core - 核心引擎
"""

from .formatter import DocumentFormatter  # @deprecated
from .style_mapper import StyleMapper  # @deprecated
from .numbering import NumberingManager
from .toc_generator import TOCGenerator
from .cross_reference import CrossReferenceManager
from .word2md_converter import Word2MDConverter
from .md2word_converter import MD2WordConverter
from .table_handler import TableHandler
from .cover_replacer import CoverReplacer
from .document_analyzer import DocumentAnalyzer
from .header_footer import HeaderFooterManager

__all__ = [
    'DocumentFormatter',
    'StyleMapper',
    'NumberingManager',
    'TOCGenerator',
    'CrossReferenceManager',
    'Word2MDConverter',
    'MD2WordConverter',
    'TableHandler',
    'CoverReplacer',
    'DocumentAnalyzer',
    'HeaderFooterManager',
]
