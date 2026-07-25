from typing import List
from models import TextBlock

class LayoutAnalyzer:
    """
    Анализ макета слайда и порядок чтения блоков.

    Задачи:
    - определить колонки (лево/средняя/правая)
    - сортировать блоки по логике чтения:
        колонка → y → x → z
    """

    def analyze(self, blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Главный метод для анализа блоков:
        1️⃣ Определяет колонки блоков
        2️⃣ Сортирует блоки по порядку чтения
        """
        self._assign_columns(blocks)
        return self._order_blocks(blocks)

    def _assign_columns(self, blocks: List[TextBlock]):
        """
        Простейшая эвристика для определения колонок.
        - Берём минимальный и максимальный X
        - Разбиваем диапазон на 3 части
        - Определяем колонку для каждого блока (0,1,2)
        """
        if not blocks:
            return

        xs = [b.x for b in blocks]
        min_x, max_x = min(xs), max(xs)
        threshold = (max_x - min_x) / 3 if max_x > min_x else 0

        for b in blocks:
            if threshold == 0:
                b.column = 0
            elif b.x < min_x + threshold:
                b.column = 0
            elif b.x < min_x + 2 * threshold:
                b.column = 1
            else:
                b.column = 2

    def _order_blocks(self, blocks: List[TextBlock]) -> List[TextBlock]:
        """
        Сортировка блоков по порядку чтения:
        - Сначала колонка (лево → центр → право)
        - Потом y (сверху вниз)
        - Потом x (слева направо)
        - Потом z (слой)
        """
        return sorted(
            blocks,
            key=lambda b: (
                getattr(b, "column", 0),  # колонка по эвристике
                b.y,
                b.x,
                b.z
            )
        )
