from typing import List
from models import TextBlock

class SectionGrouper:
    """
    Группирует текстовые блоки по заголовкам.

    Этапы:
    1️⃣ Определяем, какой блок является заголовком (is_title = True)
    2️⃣ Присваиваем блокам section (название заголовка)
    """

    def group(self, blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Главный метод группировки.
        Принимает блоки в порядке LayoutAnalyzer.
        Возвращает те же блоки с добавленным атрибутом 'section'.
        """
        # Инициализация
        current_section = "Без заголовка"

        for block in blocks:
            # Определяем, является ли блок заголовком
            self._detect_title(block)

            if getattr(block, "is_title", False):
                # Если блок — заголовок, обновляем текущий раздел
                current_section = block.text.strip()

            # Каждому блоку присваиваем секцию
            block.section = current_section

        return blocks

    def _detect_title(self, block: TextBlock):
        """
        Эвристика заголовка:
        - короткий текст (<80 символов)
        - одна строка
        - можно расширить по высоте, шрифту, стилю
        """
        lines = block.text.splitlines()

        if len(lines) == 1 and len(block.text.strip()) < 80:
            block.is_title = True
            block.level = 1
        else:
            block.is_title = False
            block.level = 0
