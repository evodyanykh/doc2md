from typing import List
from models import Document, TextBlock

class MarkdownRenderer:
    """
    Преобразует Document → Markdown

     Особенности:
    - Заголовки используются как структура (section)
    - Контент выводится под заголовками
    - Поддерживает маркеры списков
    - Добавляет разделитель слайдов
    """

    def render(self, document: Document) -> str:
        lines: List[str] = []

        for slide in document.slides:
            # Разделитель слайдов
            lines.append(f"\n---\n<!-- Slide {slide.slide_index} -->\n")
            
            if slide.title:
                lines.append(f"# {slide.title}\n")
                
            last_section = None

            for block in slide.blocks:
                # Добавляем заголовок секции один раз
                #section = getattr(block, "section", None)
                #if section and section != last_section and section != "Без заголовка":
                #    lines.append(f"## {section}\n")
                #    last_section = section

                # Добавляем содержимое блока  
                # Рендерим ТОЛЬКО контентные блоки
                lines.extend(self._render_block(block))

        return "\n\n".join(lines)

    def _render_block(self, block: TextBlock) -> list[str]:
        if getattr(block, "is_title", False):
            return []

        result: list[str] = []
        current_line = ""

        for text, is_bold in block.fragments:

            # обработка переноса строки
            if text == "\n":
                if current_line:
                    result.append(current_line)
                    current_line = ""
                continue

            if is_bold:
                current_line += f"**{text.strip()}**"
            else:
                current_line += text

        if current_line:
            result.append(current_line)

        result.append("")  # пустая строка между блоками
        return result
