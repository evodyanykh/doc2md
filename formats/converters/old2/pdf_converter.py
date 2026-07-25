"""
Конвертер PDF → Markdown.

Два режима работы для каждой страницы:
  1. Цифровой PDF: извлечение текста через PyMuPDF с анализом блоков и шрифтов.
  2. Скан (мало текста): рендеринг страницы в изображение → OCR через Tesseract.

Что сохраняется:
  - Структура заголовков (определяется по размеру шрифта относительно основного)
  - Жирный текст (определяется по флагу шрифта)
  - Колонки: определяются по горизонтальному положению блоков, каждая колонка
    выводится отдельной секцией
  - Таблицы (через pdfplumber — лучший инструмент для PDF-таблиц)
  - Символы-маркеры (□ ■ ● • ▪ и др.) нормализуются в markdown-буллеты

Библиотеки: PyMuPDF (fitz), pdfplumber, pytesseract
"""

from pathlib import Path

import fitz  # PyMuPDF
import pdfplumber
from PIL import Image

from config import IMAGE_PLACEHOLDER, OCR_FAILED_PLACEHOLDER, PDF_MIN_CHARS_FOR_TEXT, PDF_OCR_DPI
from converters.base import BaseConverter, ConversionResult
from utils.ocr import ocr_pdf_page_image, is_tesseract_available


# Флаги шрифта в PyMuPDF
FONT_FLAG_BOLD = 1 << 4   # бит 4 = bold
FONT_FLAG_ITALIC = 1 << 1  # бит 1 = italic

# Порог: если размер шрифта блока больше основного на этот множитель — это заголовок
HEADING_SIZE_RATIO_H1 = 1.6   # H1
HEADING_SIZE_RATIO_H2 = 1.35  # H2
HEADING_SIZE_RATIO_H3 = 1.15  # H3

# Unicode-символы, используемые в PDF как маркеры списков.
# Заменяются на markdown-буллет "- " при обработке текста.
BULLET_CHARS = {
    "\u25a1",  # □ пустой квадрат (checkbox)
    "\u25aa",  # ▪ маленький чёрный квадрат
    "\u25ab",  # ▫ маленький белый квадрат
    "\u25a0",  # ■ чёрный квадрат
    "\u25cf",  # ● чёрный круг
    "\u25cb",  # ○ белый круг
    "\u2022",  # • буллет
    "\u2023",  # ‣ треугольный буллет
    "\u25b8",  # ▸ малый правый треугольник
    "\u25ba",  # ► правый треугольник
    "\u2751",  # ❑ белый квадрат с тенью (часто используется в PPTX-слайдах)
    "\u274f",  # ❏ белый квадрат без тени
    "\u2610",  # ☐ квадратная скобка (ballot box)
    "\u2611",  # ☑ галочка в квадрате
    "\u2612",  # ☒ крест в квадрате
    "\u25e6",  # ◦ белый буллет
    "\u2043",  # ⁃ дефис-буллет
    # Wingdings/Symbol глифы — PyMuPDF не декодирует их в Unicode,
    # возвращает raw codepoint из Private Use Area (PUA)
    "\uf071",  # ❑ в кодировке Wingdings (встречается в PPTX→PDF)
    "\uf0b7",  # • в кодировке Wingdings
    "\uf0a8",  # ► в кодировке Wingdings
    "\uf0fc",  # ✓ в кодировке Wingdings
    # Стрелки — часто используются как маркеры в корпоративных презентациях
    "\u2794",  # ➔ тяжёлая широкая стрелка вправо (самый частый вариант)
    "\u279c",  # ➜ тяжёлая круглая стрелка вправо
    "\u27a1",  # ➡ чёрная стрелка вправо
    "\u2192",  # → стрелка вправо
    "\u21d2",  # ⇒ двойная стрелка вправо
    "\u25b6",  # ▶ чёрный правый треугольник
    "\u25b7",  # ▷ белый правый треугольник
    "\u2713",  # ✓ галочка
    "\u2714",  # ✔ жирная галочка
}

# Если более этой доли строк страницы сосредоточена в левой И правой
# половинах страницы — считаем страницу двухколоночной
COLUMN_DETECTION_THRESHOLD = 0.25

# ─── Параметры детектора невидимых таблиц (grid-table) ───────────────────────
# PPTX-таблицы без нарисованных линий pdfplumber не находит.
# Детектор ищет текстовые элементы, образующие сетку: несколько коротких значений
# выровнены по одним и тем же X-позициям на нескольких уровнях Y.
GRID_Y_TOLERANCE = 35      # строки с разницей Y-центра ≤ этого значения — одна строка таблицы
GRID_X_TOLERANCE = 100     # элементы с разницей X0 ≤ этого значения — один столбец
GRID_MIN_COLS = 3           # минимум столбцов для распознавания таблицы
GRID_MIN_ROWS = 3           # минимум строк (включая заголовок)
GRID_MAX_CELL_CHARS = 60    # максимум символов в ячейке (фильтр текстовых блоков)


class PdfConverter(BaseConverter):

    def convert(self, file_path: Path) -> ConversionResult:
        result = ConversionResult()

        doc = fitz.open(str(file_path))
        all_pages: list[str] = []

        # Вычисляем «основной» размер шрифта для всего документа
        # (самый часто встречающийся = размер основного текста)
        base_font_size = self._detect_base_font_size(doc)

        for page_num, page in enumerate(doc, start=1):
            page_text = self._convert_page(page, page_num, base_font_size, result)
            if page_text.strip():
                all_pages.append(page_text)

        doc.close()

        result.content = self._clean_text("\n\n---\n\n".join(all_pages))
        return result

    # ─── Страница ─────────────────────────────────────────────────────────────

    def _convert_page(
        self, page: fitz.Page, page_num: int, base_font_size: float, result: ConversionResult
    ) -> str:
        """
        Конвертирует одну страницу PDF.
        Определяет: цифровой текст или скан, и вызывает соответствующий метод.
        """
        # Пробуем извлечь текст напрямую
        raw_text = page.get_text("text").strip()

        # Если текста мало — это скан, используем OCR
        if len(raw_text) < PDF_MIN_CHARS_FOR_TEXT:
            return self._ocr_page(page, page_num, result)

        # Цифровой PDF: извлекаем с анализом форматирования
        return self._extract_formatted_text(page, page_num, base_font_size, result)

    # ─── Цифровой PDF: извлечение форматированного текста ────────────────────

    def _extract_formatted_text(
        self, page: fitz.Page, page_num: int, base_font_size: float, result: ConversionResult
    ) -> str:
        """
        Извлекает текст с форматированием.

        Ключевой момент: работаем на уровне отдельных СТРОК (lines), а не блоков.
        PyMuPDF при sort=True может смешивать строки из двух колонок в один блок —
        поэтому собираем все строки страницы, для каждой запоминаем координаты,
        затем сами определяем порядок чтения с учётом колонок.
        """
        lines_out: list[str] = []
        lines_out.append(f"## Страница {page_num}")
        lines_out.append("")

        # Обрабатываем таблицы один раз: получаем их bbox (для исключения из текста)
        # и готовый markdown. Фильтр: только таблицы с >= 2 колонками и >= 2 строками.
        table_bboxes, table_md = self._process_page_tables(page, result)

        # Собираем все строки страницы БЕЗ sort=True —
        # sort=True в PyMuPDF читает колонки по горизонтали, смешивая их
        page_dict = page.get_text("dict")
        image_count = 0

        # Каждый элемент: {"y": float, "x": float, "block_n": int, "spans": list}
        all_lines: list[dict] = []

        for block in page_dict.get("blocks", []):
            if block.get("type") == 1:
                image_count += 1
                continue
            if block.get("type") != 0:
                continue
            block_bbox = fitz.Rect(block["bbox"])
            if self._is_in_table(block_bbox, table_bboxes):
                continue

            block_n = block.get("number", 0)
            for line in block.get("lines", []):
                bbox = line["bbox"]
                spans = line.get("spans", [])
                if not spans:
                    continue
                all_lines.append({
                    "y":  (bbox[1] + bbox[3]) / 2,  # вертикальный центр строки
                    "x":  (bbox[0] + bbox[2]) / 2,  # горизонтальный центр строки
                    "y0": bbox[1],                   # верхний край — для слияния строк
                    "y1": bbox[3],                   # нижний край
                    "x0": bbox[0],                   # левый край — для определения отступа
                    "block_n": block_n,
                    "spans": spans,
                })

        if image_count > 0:
            result.add_warning(f"Страница {page_num}: {image_count} изображений пропущено")

        # Детектируем символы-маркеры через pdfplumber (PyMuPDF теряет ❑ из спецшрифтов)
        all_lines = self._inject_bullets_from_pdfplumber(page, all_lines, table_bboxes)

        # Сливаем строки-продолжения (перенос слова по ширине внутри одного абзаца)
        all_lines = self._merge_wrapped_lines(all_lines)

        # Ищем невидимую таблицу-сетку (PPTX-таблицы без нарисованных линий)
        grid_table_md, all_lines = self._try_extract_grid_table(all_lines, base_font_size)

        # Определяем колонки на уровне строк и рендерим
        rendered = self._render_lines_with_columns(all_lines, page.rect.width, base_font_size)
        lines_out.extend(rendered)

        # Невидимая таблица (grid-table)
        if grid_table_md:
            lines_out.append("")
            lines_out.extend(grid_table_md)

        # Реальные таблицы (отфильтрованные pdfplumber-ом) в конце страницы
        if table_md:
            lines_out.append("")
            lines_out.extend(table_md)

        return "\n".join(lines_out)

    # ─── Определение колонок и рендеринг строк ───────────────────────────────

    def _render_lines_with_columns(
        self, all_lines: list[dict], page_width: float, base_font_size: float
    ) -> list[str]:
        """
        Определяет колонки и рендерит в правильном порядке.

        Алгоритм:
          1. Разделяем строки на LEFT (x < mid) и RIGHT (x >= mid)
          2. Если обе группы достаточно заполнены — двухколоночный макет
          3. Находим Y-границу колонок: максимальный Y правой колонки.
             Строки LEFT ниже этой границы — «общий» текст под колонками.
          4. Порядок вывода: [заголовок над колонками] → [LEFT колонка] →
             [RIGHT колонка] → [общий текст под колонками]
        """
        if not all_lines:
            return []

        mid = page_width / 2
        left = [l for l in all_lines if l["x"] < mid]
        right = [l for l in all_lines if l["x"] >= mid]

        is_two_col = (
            len(left) >= 2 and len(right) >= 2
            and len(left) / len(all_lines) >= COLUMN_DETECTION_THRESHOLD
            and len(right) / len(all_lines) >= COLUMN_DETECTION_THRESHOLD
        )

        if not is_two_col:
            all_sorted = sorted(all_lines, key=lambda l: l["y"])
            return self._render_line_group(all_sorted, base_font_size)

        # ── Двухколоночный макет ──────────────────────────────────────────────

        # Y-граница колонок = нижний край правой колонки.
        # Строки LEFT, которые ниже этой границы — общий текст под колонками.
        right_max_y = max(l["y"] for l in right)

        # Строки над колонками: LEFT-строки, которые выше первой RIGHT-строки
        right_min_y = min(l["y"] for l in right)
        pre_col = [l for l in left if l["y"] < right_min_y]

        # Строки левой колонки: LEFT в зоне Y правой колонки
        left_col = [l for l in left if right_min_y <= l["y"] <= right_max_y]
        left_min_y = min((l["y"] for l in left_col), default=right_min_y)

        # Строки правой колонки, расположенные ВЫШЕ начала левой колонки —
        # по сути «заголовок» над двухколоночной зоной (например заголовок слайда
        # в правой половине, когда левая занята изображением).
        right_pre_col = [l for l in right if l["y"] < left_min_y]
        right_body    = [l for l in right if l["y"] >= left_min_y]

        # Строки под колонками: LEFT ниже правой колонки
        post_col = [l for l in left if l["y"] > right_max_y]

        output: list[str] = []

        if pre_col:
            output.extend(self._render_line_group(
                sorted(pre_col, key=lambda l: l["y"]), base_font_size
            ))
            output.append("")

        if right_pre_col:
            output.extend(self._render_line_group(
                sorted(right_pre_col, key=lambda l: l["y"]), base_font_size
            ))
            output.append("")

        if left_col:
            output.extend(self._render_line_group(
                sorted(left_col, key=lambda l: l["y"]), base_font_size
            ))
            output.append("")

        if right_body:
            output.extend(self._render_line_group(
                sorted(right_body, key=lambda l: l["y"]), base_font_size
            ))

        if post_col:
            output.append("")
            output.extend(self._render_line_group(
                sorted(post_col, key=lambda l: l["y"]), base_font_size
            ))

        return output

    def _render_line_group(
        self, lines: list[dict], base_font_size: float
    ) -> list[str]:
        """
        Рендерит группу строк в Markdown.
        Вставляет пустую строку при смене блока (разрыв абзаца).
        Если строка помечена is_bullet — добавляет префикс "- ".
        """
        output: list[str] = []
        prev_block_n: int | None = None

        for item in lines:
            text = self._process_spans(item["spans"], base_font_size)

            # Проверяем наличие реального контента.
            # split() удаляет ВСЕ Unicode-пробелы включая \u00a0.
            # Убираем markdown-маркеры и символы-буллеты (включая Wingdings \uf0xx),
            # чтобы строки вида "- -", "--", "- \uf071" считались пустыми.
            visible = "".join(text.split())
            clean = visible.replace("-", "").replace("*", "")
            for bc in BULLET_CHARS:
                clean = clean.replace(bc, "")
            if not clean:
                continue

            # Пустая строка при смене блока = разрыв между абзацами
            if prev_block_n is not None and item["block_n"] != prev_block_n:
                output.append("")

            # Если pdfplumber определил строку как элемент списка —
            # снимаем заголовочный префикс (маркер списка и H1/H2/H3 несовместимы)
            if item.get("is_bullet"):
                for _pfx in ("# ", "## ", "### ", "#### "):
                    if text.startswith(_pfx):
                        text = text[len(_pfx):]
                        break
                if not text.startswith("- "):
                    text = f"- {text}"

            output.append(text)
            prev_block_n = item["block_n"]

        return output

    def _process_spans(self, spans: list[dict], base_font_size: float) -> str:
        """
        Обрабатывает spans одной строки.
        Применяет форматирование bold/italic и определяет уровень заголовка.
        """
        if not spans:
            return ""

        # Определяем максимальный размер шрифта в строке
        max_size = max(span.get("size", 0) for span in spans)

        # Определяем уровень заголовка по соотношению размеров
        heading_prefix = self._get_heading_prefix(max_size, base_font_size)

        parts: list[str] = []
        for span in spans:
            # Пропускаем невидимый текст.
            # В PDF текст может быть скрыт двумя способами:
            # 1. rendering mode (флаг в PyMuPDF) — но он не всегда достоверен
            # 2. Белый/прозрачный цвет текста — color близко к 0xFFFFFF
            if self._is_span_hidden(span):
                continue

            text = span.get("text", "").rstrip("\n")
            if not text:
                continue

            flags = span.get("flags", 0)
            font_name = span.get("font", "")

            # Определяем bold: через флаг или по имени шрифта
            is_bold = bool(flags & FONT_FLAG_BOLD) or self._font_name_is_bold(font_name)
            is_italic = bool(flags & FONT_FLAG_ITALIC)

            # Не применяем inline-форматирование к заголовкам (избыточно).
            # Проверяем text.strip() — пустой или пробельный текст с маркерами
            # даёт артефакты вида **** или **  **, их нужно пропускать.
            if not heading_prefix and text.strip():
                if is_bold and is_italic:
                    text = f"***{text}***"
                elif is_bold:
                    text = f"**{text}**"
                elif is_italic:
                    text = f"*{text}*"

            parts.append(text)

        line_text = "".join(parts).strip()

        if heading_prefix and line_text:
            return f"{heading_prefix} {line_text}"

        # Нормализуем строки, начинающиеся с символа-маркера, в markdown-буллет
        return self._normalize_bullet(line_text)

    def _get_heading_prefix(self, font_size: float, base_size: float) -> str:
        """
        Определяет Markdown-префикс заголовка по размеру шрифта.
        Возвращает "" если это обычный текст.
        """
        if base_size <= 0:
            return ""

        ratio = font_size / base_size

        if ratio >= HEADING_SIZE_RATIO_H1:
            return "#"
        elif ratio >= HEADING_SIZE_RATIO_H2:
            return "##"
        elif ratio >= HEADING_SIZE_RATIO_H3:
            return "###"
        return ""

    def _is_span_hidden(self, span: dict) -> bool:
        """
        Определяет, является ли текстовый span невидимым.

        Два признака невидимости в PDF:
          1. Цвет текста белый или почти белый (>= 250 по каждому каналу RGB).
             Используется в слайдах для скрытия текста на белом фоне.
             ВАЖНО: на тёмном фоне белый текст — это нормально. Поэтому проверяем
             только явно белый (255, 255, 255) — точный совпадение.
          2. origin_color == color == 0xFFFFFF и размер шрифта <= 1 pt —
             классический паттерн скрытого текста в PDF-формах.
        """
        try:
            color = span.get("color", 0)
            # Белый цвет в PyMuPDF: 0xFFFFFF = 16777215
            if color == 0xFFFFFF:
                # Дополнительная проверка: шрифт очень маленький = точно скрытый
                # (белый текст нормального размера может быть на тёмном фоне)
                size = span.get("size", 12)
                if size <= 1:
                    return True
        except Exception:
            pass
        return False

    def _font_name_is_bold(self, font_name: str) -> bool:
        """
        Эвристика: шрифт считается жирным если его имя содержит 'Bold', 'Heavy', 'Black'.
        Нужно потому что не все PDF правильно выставляют флаг bold.
        """
        name_upper = font_name.upper()
        return any(kw in name_upper for kw in ("BOLD", "HEAVY", "BLACK", "DEMI"))

    def _normalize_bullet(self, text: str) -> str:
        """
        Заменяет Unicode-символы маркеров списка в начале строки на markdown "- ".

        Например: "□ Миссия, заданная 7-ФЗ" → "- Миссия, заданная 7-ФЗ"

        Обрабатывает случаи:
          - Одиночный маркер: "□текст" → "- текст"
          - Несколько маркеров подряд: "□ □текст" → "- текст"  (после слияния строк)
          - Только маркеры без текста: "□ □" → "" (пустая строка, будет отфильтрована)
          - Маркер с отступом: "  □ текст" → "- текст"
        """
        if not text:
            return text

        # Снимаем начальный пробел для проверки
        stripped = text.lstrip()
        if not stripped:
            return ""

        if stripped[0] not in BULLET_CHARS:
            return text

        # Сохраняем уровень отступа для вложенных списков
        indent = len(text) - len(stripped)
        indent_md = "  " * (indent // 4) if indent >= 4 else ""

        # Снимаем ВСЕ ведущие маркеры подряд (после слияния строк может быть несколько)
        rest = stripped[1:].lstrip()
        while rest and rest[0] in BULLET_CHARS:
            rest = rest[1:].lstrip()

        # Если после маркеров ничего нет — пустая строка (будет отфильтрована)
        if not rest.strip():
            return ""

        return f"{indent_md}- {rest}"

    # ─── Основной размер шрифта документа ────────────────────────────────────

    def _detect_base_font_size(self, doc: fitz.Document) -> float:
        """
        Определяет «основной» размер шрифта документа.
        Берёт первые 5 страниц, считает частоту каждого размера,
        возвращает самый частый (= размер основного текста).
        """
        size_counts: dict[float, int] = {}

        for page in list(doc)[:5]:
            page_dict = page.get_text("dict")
            for block in page_dict.get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        size = round(span.get("size", 0), 1)
                        if size > 0:
                            size_counts[size] = size_counts.get(size, 0) + 1

        if not size_counts:
            return 12.0  # fallback

        # Возвращаем самый частый размер
        return max(size_counts, key=size_counts.get)

    # ─── Таблицы через pdfplumber ─────────────────────────────────────────────

    def _process_page_tables(
        self, page: fitz.Page, result: ConversionResult
    ) -> tuple[list[fitz.Rect], list[str]]:
        """
        Открывает pdfplumber один раз на страницу, находит реальные таблицы и:
          1. Возвращает их bbox — чтобы исключить текст таблиц из основного извлечения
          2. Конвертирует таблицы в Markdown

        Фильтры — отсеивают ложные таблицы из PDF-слайдов:
          - >= 2 строки И >= 2 колонки                (минимальный размер)
          - > 40% пустых ячеек → пропуск              (декоративная рамка слайда)
          - любая ячейка > 200 символов → пропуск     (текстовый блок, не таблица данных)
        """
        bboxes: list[fitz.Rect] = []
        md_lines: list[str] = []

        try:
            with pdfplumber.open(page.parent.name) as pdf:
                pl_page = pdf.pages[page.number]

                page_w = page.rect.width
                page_h = page.rect.height

                for table in pl_page.find_tables():
                    bbox = table.bbox

                    # Фильтр: таблица покрывает ≥ 95% страницы по обоим измерениям →
                    # это граница/разметка слайда, а не таблица данных
                    if (page_w > 0 and page_h > 0
                            and (bbox[2] - bbox[0]) >= page_w * 0.95
                            and (bbox[3] - bbox[1]) >= page_h * 0.95):
                        continue

                    extracted = table.extract()

                    # Минимальный размер: >= 2 строки и >= 2 колонки
                    if not extracted or len(extracted) < 2:
                        continue
                    if not extracted[0] or len(extracted[0]) < 2:
                        continue

                    all_cells = [str(c or "").strip() for row in extracted for c in row]

                    # Фильтр: > 40% пустых ячеек = декоративная рамка слайда
                    empty_ratio = sum(1 for c in all_cells if not c) / len(all_cells)
                    if empty_ratio > 0.4:
                        continue

                    # Фильтр: любая ячейка длиннее 200 символов = текстовый блок, не данные
                    max_cell_len = max(len(c) for c in all_cells)
                    if max_cell_len > 200:
                        continue

                    # Прошла все фильтры — реальная таблица данных
                    bbox = table.bbox
                    bboxes.append(fitz.Rect(bbox[0], bbox[1], bbox[2], bbox[3]))

                    table_md = self._table_to_markdown(extracted, result)
                    md_lines.extend(table_md)
                    md_lines.append("")

        except Exception as e:
            result.add_warning(f"Ошибка при обработке таблиц: {e}")

        return bboxes, md_lines

    def _inject_bullets_from_pdfplumber(
        self,
        page: fitz.Page,
        all_lines: list[dict],
        table_bboxes: list[fitz.Rect],
    ) -> list[dict]:
        """
        PyMuPDF теряет символы-маркеры (❑ □ ● и т.д.) из спецшрифтов — они
        либо пустые, либо не декодируются. pdfplumber видит их как обычные символы.

        Алгоритм:
          1. Берём все символы страницы через pdfplumber (pl_page.chars)
          2. Находим Y-позиции где стоят символы-маркеры
          3. Для каждой строки all_lines: если на той же Y есть маркер —
             помечаем строку как элемент списка (is_bullet=True)
        """
        # Полный набор Unicode-символов маркеров
        all_bullet_chars = BULLET_CHARS | {
            "\u2751", "\u274f", "\u25a1", "\u25aa", "\u2022",
            "\u25cf", "\u25cb", "\u2610", "\u2611", "\u2612",
            "\uf0a8", "\uf0b7",   # Wingdings/Symbol bullet glyphs
        }
        Y_TOLERANCE = 8  # допуск совпадения Y в пунктах (больше из-за разного origin)

        try:
            with pdfplumber.open(page.parent.name) as pdf:
                pl_page = pdf.pages[page.number]
                chars = pl_page.chars  # список всех символов с координатами
        except Exception:
            return all_lines

        # Строим список (top, x0) всех символов-маркеров на странице
        bullet_positions: list[tuple[float, float]] = []
        for ch in chars:
            if ch.get("text", "") in all_bullet_chars:
                bullet_positions.append((ch["top"], ch["x0"]))

        if not bullet_positions:
            return all_lines

        # Для каждой строки проверяем совпадение по Y и X.
        # Обе проверки необходимы: при двухколоночном макете буллеты LEFT (x≈96)
        # и RIGHT (x≈976) находятся на одинаковом Y — без проверки X они
        # ложно помечают текст из противоположной колонки.
        updated = []
        for line in all_lines:
            line_y0 = line.get("y0", line["y"])
            line_x0 = line.get("x0", line["x"])

            is_bull = any(
                abs(by - line_y0) <= Y_TOLERANCE and abs(bx - line_x0) <= 300
                for by, bx in bullet_positions
            )
            if is_bull:
                line = dict(line)
                line["is_bullet"] = True
            updated.append(line)

        return updated

    def _merge_wrapped_lines(self, all_lines: list[dict]) -> list[dict]:
        """
        Сливает строки-продолжения внутри одного блока.

        Проблема: PDF хранит каждую визуальную строку отдельно. Если текст
        не поместился по ширине, он перенесён на следующую строку в том же блоке.
        В результате одна фраза разбита на несколько строк.

        Критерий слияния — обе строки в одном блоке И вертикальный зазор
        между ними меньше высоты строки (т.е. это перенос, а не новый абзац).

        НЕ сливаем: если следующая строка помечена as is_bullet (новый пункт).
        """
        if not all_lines:
            return all_lines

        merged: list[dict] = []
        i = 0

        while i < len(all_lines):
            current = dict(all_lines[i])

            # Пробуем присоединить следующие строки-продолжения
            while i + 1 < len(all_lines):
                nxt = all_lines[i + 1]

                # Следующая строка — новый пункт списка → не сливаем
                if nxt.get("is_bullet"):
                    break

                # Резервная проверка: если span-текст следующей строки начинается
                # с символа-маркера — это новый пункт списка, не продолжение абзаца.
                # Покрывает случаи, когда pdfplumber не смог определить маркер
                # из-за нестандартной кодировки шрифта.
                nxt_leading = "".join(
                    s.get("text", "") for s in nxt.get("spans", [])
                ).lstrip()
                if nxt_leading and nxt_leading[0] in BULLET_CHARS:
                    break

                # Если следующая строка расположена ВЫШЕ текущей — это другая колонка
                # (блоки пронумерованы слева направо в PPTX, а не сверху вниз).
                if nxt["y0"] < current["y0"]:
                    break

                # Нумерованный список («1. », «2. » и т.п.) — не сливаем с предыдущим
                if (len(nxt_leading) >= 2
                        and nxt_leading[0].isdigit()
                        and nxt_leading[1] in ".)"
                        and (len(nxt_leading) == 2 or not nxt_leading[2].isdigit())):
                    break

                # Строки в разных горизонтальных зонах (разные колонки) — не сливаем.
                # Порог 200pt покрывает и A4 (~595pt) и высокоDPI PDF (1920pt+).
                # Примечание: проверку block_n убрали — continuation-строки
                # в PPTX→PDF часто оказываются в отдельном PDF-блоке.
                x_dist = abs(nxt.get("x0", nxt["x"]) - current.get("x0", current["x"]))
                if x_dist > 200:
                    break

                # Зазор: используем высоту СЛЕДУЮЩЕЙ строки как эталон,
                # чтобы порог не рос бесконтрольно при накоплении строк.
                line_h = nxt["y1"] - nxt["y0"]
                gap = nxt["y0"] - current["y1"]

                # Зазор больше 0.8 высоты строки → смысловой разрыв, не сливаем.
                # (0.8 вместо 1.2 — более строгий критерий; обычный перенос внутри
                # абзаца имеет зазор ~0.05-0.2, а межабзацный отступ ~0.9-1.5)
                if line_h > 0 and gap > line_h * 0.8:
                    break

                # Сливаем: добавляем пробел между строками и объединяем spans
                space_span = {"text": " ", "size": 0, "flags": 0, "font": ""}
                current["spans"] = current["spans"] + [space_span] + nxt["spans"]
                current["y1"] = nxt["y1"]
                current["y"] = (current["y0"] + current["y1"]) / 2
                i += 1

            merged.append(current)
            i += 1

        return merged

    def _is_in_table(self, block_bbox: fitz.Rect, table_bboxes: list[fitz.Rect]) -> bool:
        """Проверяет, находится ли блок текста внутри зоны реальной таблицы."""
        for table_rect in table_bboxes:
            # Пересечение > 50% площади блока = блок внутри таблицы
            intersection = block_bbox & table_rect
            if not intersection.is_empty:
                block_area = block_bbox.get_area()
                if block_area > 0 and intersection.get_area() / block_area > 0.5:
                    return True
        return False

    def _table_to_markdown(self, table: list[list], result: ConversionResult) -> list[str]:
        """Конвертирует список списков (таблицу из pdfplumber) в Markdown."""
        if not table:
            return []

        # Очищаем ячейки: None → "", переносы строк → пробел
        cleaned: list[list[str]] = []
        for row in table:
            cleaned.append([
                str(cell).replace("\n", " ").strip() if cell is not None else ""
                for cell in row
            ])

        col_count = max(len(row) for row in cleaned)
        for row in cleaned:
            while len(row) < col_count:
                row.append("")

        md_lines: list[str] = []
        md_lines.append("| " + " | ".join(cleaned[0]) + " |")
        md_lines.append("| " + " | ".join(["---"] * col_count) + " |")
        for row in cleaned[1:]:
            md_lines.append("| " + " | ".join(row) + " |")

        if col_count > 2:
            result.add_warning(f"Таблица {len(cleaned)}×{col_count} в PDF")

        return md_lines

    # ─── Детектор невидимых таблиц (grid-table) ──────────────────────────────

    def _try_extract_grid_table(
        self, all_lines: list[dict], base_font_size: float
    ) -> tuple[list[str], list[dict]]:
        """
        Ищет невидимые таблицы — без нарисованных линий (типичны для PPTX→PDF).
        pdfplumber их не находит, поэтому применяем позиционную эвристику:
        если несколько коротких текстовых элементов образуют сетку (строки × столбцы),
        то это таблица.

        Критерии:
          - Группируем строки по Y-позиции (Y_TOLERANCE).
          - «Строки-кандидаты» — те, у которых >= GRID_MIN_COLS элементов.
          - Минимум GRID_MIN_ROWS таких строк.
          - Все ячейки коротки (< GRID_MAX_CELL_CHARS символов).
          - Позиции X согласованы между строками (GRID_X_TOLERANCE).

        Возвращает (table_md_lines, remaining_lines):
          - table_md_lines: строки Markdown-таблицы (пусто, если таблица не найдена)
          - remaining_lines: строки, не вошедшие в таблицу
        """
        if len(all_lines) < GRID_MIN_COLS * GRID_MIN_ROWS:
            return [], all_lines

        # Шаг 1: группируем по Y ───────────────────────────────────────────────
        y_groups: list[list[dict]] = []
        for line in sorted(all_lines, key=lambda l: l["y"]):
            placed = False
            for group in y_groups:
                if abs(line["y"] - group[0]["y"]) <= GRID_Y_TOLERANCE:
                    group.append(line)
                    placed = True
                    break
            if not placed:
                y_groups.append([line])

        # Шаг 2: строки-кандидаты — содержат >= GRID_MIN_COLS элементов ────────
        candidate_rows = [g for g in y_groups if len(g) >= GRID_MIN_COLS]
        if len(candidate_rows) < GRID_MIN_ROWS:
            return [], all_lines

        # Шаг 3: все ячейки должны быть короткими ─────────────────────────────
        for row in candidate_rows:
            for line in row:
                text = "".join(
                    s.get("text", "") for s in line.get("spans", [])
                ).strip()
                if len(text) > GRID_MAX_CELL_CHARS:
                    return [], all_lines

        # Шаг 4: строим множество X-позиций столбцов ──────────────────────────
        col_xs: list[float] = []
        for row in candidate_rows:
            for line in sorted(row, key=lambda l: l["x0"]):
                x = line["x0"]
                if not any(abs(x - cx) <= GRID_X_TOLERANCE for cx in col_xs):
                    col_xs.append(x)
        col_xs.sort()

        if len(col_xs) < GRID_MIN_COLS:
            return [], all_lines

        # Шаг 5: каждая строка-кандидат должна покрывать >= GRID_MIN_COLS столбцов
        for row in candidate_rows:
            row_xs = [l["x0"] for l in row]
            covered = sum(
                1 for cx in col_xs
                if any(abs(rx - cx) <= GRID_X_TOLERANCE for rx in row_xs)
            )
            if covered < GRID_MIN_COLS:
                return [], all_lines

        # Таблица найдена! ─────────────────────────────────────────────────────
        table_ids = {id(l) for row in candidate_rows for l in row}
        remaining = [l for l in all_lines if id(l) not in table_ids]

        # Шаг 6: формируем Markdown ───────────────────────────────────────────

        def cell_text(line: dict) -> str:
            return "".join(s.get("text", "") for s in line.get("spans", [])).strip()

        def row_to_cells(row: list[dict]) -> list[str]:
            cells = [""] * len(col_xs)
            for line in sorted(row, key=lambda l: l["x0"]):
                x = line["x0"]
                best = min(range(len(col_xs)), key=lambda i: abs(col_xs[i] - x))
                if abs(col_xs[best] - x) <= GRID_X_TOLERANCE:
                    t = cell_text(line)
                    cells[best] = (f"{cells[best]} {t}".strip()) if cells[best] else t
            return cells

        sorted_rows = sorted(candidate_rows, key=lambda g: g[0]["y"])
        md_rows = [row_to_cells(row) for row in sorted_rows]

        if not md_rows:
            return [], all_lines

        header = md_rows[0]
        md_lines = ["| " + " | ".join(header) + " |"]
        md_lines.append("| " + " | ".join(["---"] * len(col_xs)) + " |")
        for row_cells in md_rows[1:]:
            md_lines.append("| " + " | ".join(row_cells) + " |")

        return md_lines, remaining

    # ─── OCR для сканов ───────────────────────────────────────────────────────

    def _ocr_page(self, page: fitz.Page, page_num: int, result: ConversionResult) -> str:
        """
        Рендерит страницу PDF в изображение и применяет OCR.
        Используется когда страница является сканом.
        """
        if not is_tesseract_available():
            result.add_warning(
                f"Страница {page_num}: скан, но Tesseract недоступен — страница пропущена"
            )
            return f"## Страница {page_num}\n\n{OCR_FAILED_PLACEHOLDER}\n"

        # Рендерим страницу в изображение с нужным DPI
        # matrix масштабирует изображение: DPI/72 (72 = базовый DPI в PDF)
        zoom = PDF_OCR_DPI / 72
        matrix = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=matrix)

        # Конвертируем в PIL Image для pytesseract
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        ocr_text = ocr_pdf_page_image(img)
        result.add_warning(f"Страница {page_num} распознана через OCR (скан)")

        return f"## Страница {page_num}\n\n{ocr_text}\n"
