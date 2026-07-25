from dataclasses import dataclass
from typing import List, Optional, Tuple

@dataclass
class TextBlock:
    """
    Базовая единица текста на слайде.

    Это НЕ абзац и НЕ shape — это уже нормализованный блок,
    с которым дальше работает логика порядка и группировки.
    """
    slide_index: int

    text: str  # объединённый текст для совместимости с существующими методами
    fragments: List[Tuple[str, bool, bool]]  # (text, is_bold, is_list_item)

    # Геометрия (EMU → уже приведены к числам)
    x: float
    y: float
    width: float
    height: float

    # Z-order (важно при наложениях)
    z: int
 
    # Эвристические признаки
    is_title: bool = False
    level: Optional[int] = None  # уровень заголовка (h1, h2…)
    column: Optional[int] = None  # номер колонки


@dataclass
class SlideContent:
    """
    Содержимое одного слайда после парсинга.
    """
    slide_index: int
    title: Optional[str]
    blocks: List[TextBlock]


@dataclass
class Document:
    """
    Финальная модель документа
    """
    slides: List[SlideContent]
    
    
@dataclass
class Paragraph:
    """
    Логический абзац, собранный из нескольких TextBlock.
    """
    slide_index: int
    text: str
    blocks: List[TextBlock]

    # эвристики
    is_term_definition: bool = False