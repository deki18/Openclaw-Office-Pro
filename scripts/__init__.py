"""
Office Pro - Enterprise Document Automation Suite

企业级 Word 和 Excel 文档自动化工具
"""

from .word_processor import WordProcessor
from .excel_processor import ExcelProcessor, XlsxTemplateEngine

__version__ = '1.0.0'
__all__ = [
    'WordProcessor',
    'ExcelProcessor', 
    'XlsxTemplateEngine',
]
