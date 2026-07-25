"""
Конвертер DOCX → Markdown.

Сохраняет форматирование:
  - Заголовки (Heading 1-6) → # ## ### и т.д.
  - Жирный текст → **текст**
  - Курсив → *текст*
  - Жирный + курсив → ***текст***
  - Маркированные списки → - пункт
  - Нумерованные списки → 1. пункт (с учётом уровней вложенности)
  - Таблицы → Markdown-таблица
  - Изображения → плейсхолдер или OCR-текст

Библиотека: python-docx
"""

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from config import IMAGE_PLACEHOLDER, TABLE_WARNING
from converters.base import BaseConverter, ConversionResult
from utils.ocr import ocr_image_bytes, is_tesseract_available


class DocxConverter(BaseConverter):

    def convert(self, file_path: Path) -> ConversionResult:
        result = ConversionResult()
        doc = Document(str(file_path))

        lines: list[str] = []

        # Обходим тело документа в порядке следования элементов.
        # document.element.body содержит и параграфы, и таблицы в правильном порядке.
        for child in doc.element.body:
            tag = child.tag.split("}")[-1]  # убираем XML-namespace

            if tag == "p":
                # Параграф — текст, заголовок или элемент списка
                para = Paragraph(child, doc)
                md_line = self._convert_paragraph(para, result)
                if md_line is not None:
                    lines.append(md_line)

            elif tag == "tbl":
                # Таблица
                table = Table(child, doc)
                md_table = self._convert_table(table, result)
                lines.append("")  # пустая строка перед таблицей
                lines.extend(md_table)
                lines.append("")  # пустая строка после таблицы

        result.content = self._clean_text("\n".join(lines))
        return result

    # ─── Параграфы ────────────────────────────────────────────────────────────

    def _convert_paragraph(self, para: Paragraph, result: ConversionResult) -> str | None:
        """
        Конвертирует один параграф в Markdown-строку.
        Возвращает None для полностью пустых параграфов с изображениями,
        которые уже добавлены в result через плейсхолдер.
        """
        style_name = para.style.name if para.style else ""

        # Проверяем наличие изображений в параграфе
        image_lines = self._extract_images(para, result)

        # Получаем текст параграфа с форматированием
        text = self._convert_runs(para)

        # Если параграф содержит только изображение и нет текста
        if not text.strip() and image_lines:
            return "\n".join(image_lines)

        # Объединяем текст и изображения
        if image_lines:
            text = text + "\n" + "\n".join(image_lines)

        if not text.strip():
            return ""  # пустая строка для разделения блоков

        # Определяем тип параграфа по стилю
        return self._apply_style(text, style_name, para)

    def _convert_runs(self, para: Paragraph) -> str:
        """
        Обходит все runs параграфа и применяет форматирование:
        жирный, курсив, жирный+курсив.

        Run — минимальная единица текста с одинаковым форматированием.
        """
        parts: list[str] = []

        for run in para.runs:
            text = run.text
            if not text:
                continue

            bold = run.bold
            italic = run.italic

            # Применяем форматирование только к непустому тексту —
            # иначе получаются артефакты **** или **  **
            if text.strip():
                if bold and italic:
                    text = f"***{text}***"
                elif bold:
                    text = f"**{text}**"
                elif italic:
                    text = f"*{text}*"

            parts.append(text)

        return "".join(parts)

    def _apply_style(self, text: str, style_name: str, para: Paragraph) -> str:
        """
        Применяет стиль параграфа: заголовки, списки, обычный текст.
        """
        # ─── Заголовки ────────────────────────────────────────────────────────
        # Стили: "Heading 1", "Heading 2", ..., "Заголовок 1" (русская версия Word)
        heading_map = {
            "heading 1": "#", "заголовок 1": "#",
            "heading 2": "##", "заголовок 2": "##",
            "heading 3": "###", "заголовок 3": "###",
            "heading 4": "####", "заголовок 4": "####",
            "heading 5": "#####", "заголовок 5": "#####",
            "heading 6": "######", "заголовок 6": "######",
        }
        style_lower = style_name.lower()
        for style_key, prefix in heading_map.items():
            if style_key in style_lower:
                return f"{prefix} {text}"

        # ─── Списки ───────────────────────────────────────────────────────────
        # Определяем уровень вложенности для отступа
        indent_level = self._get_list_indent(para)
        indent = "  " * indent_level  # 2 пробела на уровень

        # Маркированный список
        if any(s in style_lower for s in ["list bullet", "list paragraph", "маркированный"]):
            return f"{indent}- {text}"

        # Нумерованный список
        if any(s in style_lower for s in ["list number", "нумерованный"]):
            # Реальный номер из XML сложно получить, используем 1. —
            # при рендере Markdown номера автоматически пересчитываются
            return f"{indent}1. {text}"

        # ─── Обычный текст ────────────────────────────────────────────────────
        return text

    def _get_list_indent(self, para: Paragraph) -> int:
        """
        Определяет уровень вложенности списка из XML-структуры параграфа.
        Возвращает 0 если не список.
        """
        try:
            # numPr/ilvl — уровень вложенности в нумерованном/маркированном списке
            num_pr = para._element.find(qn("w:numPr"))
            if num_pr is not None:
                ilvl = num_pr.find(qn("w:ilvl"))
                if ilvl is not None:
                    return int(ilvl.get(qn("w:val"), 0))
        except Exception:
            pass
        return 0

    # ─── Изображения ──────────────────────────────────────────────────────────

    def _extract_images(self, para: Paragraph, result: ConversionResult) -> list[str]:
        """
        Извлекает изображения из параграфа.
        Если Tesseract доступен — пытается распознать текст.
        Иначе вставляет плейсхолдер.
        """
        image_lines: list[str] = []
        tesseract_ok = is_tesseract_available()

        # Ищем теги изображений в XML параграфа
        for drawing in para._element.findall(".//" + qn("a:blip"), para._element.nsmap):
            # Это SVG/EMF — пропускаем, берём только растровые изображения ниже
            pass

        # Ищем встроенные изображения через relationship
        for inline in para._element.findall(".//" + qn("wp:inline")):
            blip = inline.find(".//" + qn("a:blip"))
            if blip is None:
                continue

            # r:embed — ID связи с файлом изображения
            embed = blip.get(qn("r:embed"))
            if not embed:
                continue

            try:
                # Получаем байты изображения через relationships документа
                image_part = para._element.getparent()
                # Ищем relationship в родительском документе
                rel = None
                for r in para.part.rels.values():
                    if r.rId == embed:
                        rel = r
                        break

                if rel and hasattr(rel, "target_part"):
                    image_bytes = rel.target_part.blob

                    if tesseract_ok:
                        ocr_text = ocr_image_bytes(image_bytes)
                        if ocr_text and ocr_text != IMAGE_PLACEHOLDER:
                            image_lines.append(f"\n> *[Текст из изображения]:* {ocr_text}\n")
                        else:
                            image_lines.append(IMAGE_PLACEHOLDER)
                            result.add_warning("Изображение без распознанного текста")
                    else:
                        image_lines.append(IMAGE_PLACEHOLDER)
                        result.add_warning("Tesseract недоступен — изображение пропущено")

            except Exception:
                image_lines.append(IMAGE_PLACEHOLDER)
                result.add_warning("Не удалось извлечь изображение")

        return image_lines

    # ─── Таблицы ──────────────────────────────────────────────────────────────

    def _convert_table(self, table: Table, result: ConversionResult) -> list[str]:
        """
        Конвертирует таблицу Word в Markdown-таблицу.

        Markdown-таблица:
          | Заголовок 1 | Заголовок 2 |
          |-------------|-------------|
          | Ячейка 1    | Ячейка 2    |

        Ограничение: объединённые ячейки (merge) отображаются как обычные.
        """
        rows: list[list[str]] = []

        for row in table.rows:
            cells: list[str] = []
            for cell in row.cells:
                # Собираем текст всех параграфов ячейки
                cell_text = " ".join(
                    self._convert_runs(para)
                    for para in cell.paragraphs
                    if para.text.strip()
                )
                # Убираем символ переноса строки внутри ячейки (ломает таблицу)
                cells.append(cell_text.replace("\n", " ").strip())
            rows.append(cells)

        if not rows:
            return []

        # Определяем количество колонок по максимальной строке
        col_count = max(len(row) for row in rows)

        # Дополняем короткие строки пустыми ячейками
        for row in rows:
            while len(row) < col_count:
                row.append("")

        md_lines: list[str] = []

        # Первая строка — заголовок таблицы
        header = rows[0]
        md_lines.append("| " + " | ".join(header) + " |")
        md_lines.append("| " + " | ".join(["---"] * col_count) + " |")

        # Остальные строки — данные
        for row in rows[1:]:
            md_lines.append("| " + " | ".join(row) + " |")

        # Предупреждаем если таблица большая (возможны проблемы с объединёнными ячейками)
        if len(rows) > 1 and col_count > 2:
            result.add_warning(f"Таблица {len(rows)}×{col_count}: {TABLE_WARNING}")

        return md_lines
