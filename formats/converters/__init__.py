"""
Реестр конвертеров: расширение файла → класс конвертера.

Единственное место, где выбирается конвертер по формату.
И CLI (formats/main.py), и веб-интерфейс (app.py) используют этот реестр,
поэтому добавление нового формата требует правки только этого файла.
"""

from pathlib import Path

from formats.converters.base import BaseConverter, ConversionResult
from formats.converters.docx_converter import DocxConverter
from formats.converters.pdf_converter import PdfConverter
from formats.converters.pptx_converter import PptxConverter

CONVERTERS: dict[str, type[BaseConverter]] = {
    ".docx": DocxConverter,
    ".pptx": PptxConverter,
    ".pdf": PdfConverter,
}


def get_converter(extension: str) -> BaseConverter | None:
    """
    Возвращает экземпляр конвертера для расширения файла
    или None, если формат не поддерживается.

    Args:
        extension: расширение с точкой, регистр не важен (".PDF" тоже сработает).
    """
    converter_class = CONVERTERS.get(extension.lower())
    return converter_class() if converter_class else None


def convert_file(file_path: Path) -> ConversionResult:
    """
    Конвертирует файл в Markdown, сам выбирая конвертер по расширению.

    Raises:
        ValueError: формат файла не поддерживается.
        Exception:  критическая ошибка парсинга (пробрасывается из конвертера).
    """
    converter = get_converter(file_path.suffix)
    if converter is None:
        raise ValueError(f"Неподдерживаемый формат: {file_path.suffix}")
    return converter.convert(file_path)
