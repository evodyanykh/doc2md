"""
Пакет formats — единая точка извлечения Markdown из документов.

Поддерживаемые форматы: .docx, .pptx, .pdf.

Использование:
    from formats import convert_file, SUPPORTED_EXTENSIONS

    result = convert_file(Path("lecture.pptx"))
    print(result.content)      # Markdown
    print(result.warnings)     # предупреждения конвертации
"""

from formats.config import SUPPORTED_EXTENSIONS
from formats.converters import CONVERTERS, convert_file, get_converter
from formats.converters.base import BaseConverter, ConversionResult

__all__ = [
    "SUPPORTED_EXTENSIONS",
    "CONVERTERS",
    "convert_file",
    "get_converter",
    "BaseConverter",
    "ConversionResult",
]
