"""
Базовый класс конвертера.

Каждый конкретный конвертер (docx, pptx, pdf) наследует этот класс
и реализует метод convert(). Это обеспечивает единый интерфейс
для main.py вне зависимости от формата файла.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ConversionResult:
    """
    Результат конвертации одного файла.

    Attributes:
        content:   итоговый Markdown-текст.
        warnings:  список предупреждений (изображения без текста, сложные таблицы и т.д.).
        success:   True если конвертация прошла без критических ошибок.
    """
    content: str = ""
    warnings: list[str] = field(default_factory=list)
    success: bool = True

    def add_warning(self, message: str) -> None:
        self.warnings.append(message)

    def has_warnings(self) -> bool:
        return len(self.warnings) > 0


class BaseConverter(ABC):
    """
    Абстрактный базовый класс для всех конвертеров.

    Использование:
        converter = DocxConverter()
        result = converter.convert(Path("document.docx"))
        print(result.content)
    """

    @abstractmethod
    def convert(self, file_path: Path) -> ConversionResult:
        """
        Конвертирует файл в Markdown.

        Args:
            file_path: путь к исходному файлу.

        Returns:
            ConversionResult с текстом и списком предупреждений.

        Raises:
            Exception: при критической ошибке парсинга.
                       Обрабатывается в main.py через logger.log_error().
        """
        ...

    @staticmethod
    def _clean_text(text: str) -> str:
        """
        Базовая очистка текста:
        - убирает лишние пробелы внутри строк
        - убирает более двух подряд идущих пустых строк
        """
        lines = text.splitlines()
        cleaned = []
        blank_count = 0

        for line in lines:
            # Нормализуем пробелы внутри строки, но не трогаем отступы (для списков)
            stripped = line.rstrip()
            if stripped == "":
                blank_count += 1
                # Допускаем не более двух пустых строк подряд
                if blank_count <= 2:
                    cleaned.append("")
            else:
                blank_count = 0
                cleaned.append(stripped)

        return "\n".join(cleaned).strip()
