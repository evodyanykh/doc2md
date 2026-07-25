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
import re

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

    
    # ─── Таблицы ──────────────────────────────────────────────────────────────

    def _escape_md_cell(self, text: str) -> str:
        """
        Экранирует содержимое ячейки для Markdown:
        - заменяет переносы строк на <br>
        - экранирует вертикальные черты |
        """
        text = (text or "").replace("\r", "").strip()
        text = text.replace("\n", "<br>")
        text = text.replace("|", "\\|")
        text = re.sub(r"\s*<br>\s*", "<br>", text)
        return text

    def _escape_html(self, text: str) -> str:
        """
        Минимальное экранирование для HTML-таблицы внутри Markdown.
        """
        text = (text or "").replace("\r", "").strip()
        text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        text = text.replace('"', "&quot;")
        text = text.replace("\n", "<br>")
        return text

    def _cell_text(self, cell) -> str:
        """
        Собирает текст ячейки, сохраняя базовое форматирование runs.
        """
        parts = []
        for para in cell.paragraphs:
            if para.text.strip():
                parts.append(self._convert_runs(para))
        return "\n".join([p for p in parts if p is not None]).strip()

    def _table_has_merges(self, table: Table) -> bool:
        """
        Проверяет наличие объединённых ячеек (colspan/rowspan) в таблице.
        """
        tbl = table._tbl
        return bool(tbl.xpath(".//w:tcPr/w:gridSpan | .//w:tcPr/w:vMerge"))

    def _convert_table(self, table: Table, result: ConversionResult) -> list[str]:
        """
        Конвертирует таблицу Word в Markdown.

        Если в таблице есть объединённые ячейки (merge) или очень много колонок,
        используем HTML-таблицу внутри Markdown (так корректно сохраняются colspan/rowspan).

        Иначе — обычная Markdown-таблица.
        """
        col_count_guess = max((len(r.cells) for r in table.rows), default=0)

        if self._table_has_merges(table) or col_count_guess > 12:
            return self._convert_table_to_html(table, result)

        rows: list[list[str]] = []
        for row in table.rows:
            cells: list[str] = []
            for cell in row.cells:
                cell_text = self._cell_text(cell)
                cells.append(self._escape_md_cell(cell_text))
            rows.append(cells)

        if not rows:
            return []

        col_count = max(len(row) for row in rows)
        for row in rows:
            while len(row) < col_count:
                row.append("")

        md_lines: list[str] = []
        header = rows[0]
        md_lines.append("| " + " | ".join(header) + " |")
        md_lines.append("| " + " | ".join(["---"] * col_count) + " |")
        for row in rows[1:]:
            md_lines.append("| " + " | ".join(row) + " |")

        return md_lines

    def _convert_table_to_html(self, table: Table, result: ConversionResult) -> list[str]:
        """
        Генерирует HTML-таблицу (вставляется прямо в Markdown),
        корректно обрабатывая:
          - горизонтальные объединения (w:gridSpan -> colspan)
          - вертикальные объединения (w:vMerge -> rowspan)
        """
        tbl = table._tbl

        grid_cols = tbl.xpath("./w:tblGrid/w:gridCol")
        col_count = len(grid_cols) if grid_cols else max((len(r.cells) for r in table.rows), default=0)
        row_count = len(table.rows)

        if row_count == 0 or col_count == 0:
            return []

        matrix = [[None for _ in range(col_count)] for _ in range(row_count)]

        def tc_gridspan(tc):
            gs = tc.xpath("./w:tcPr/w:gridSpan/@w:val")
            try:
                return int(gs[0]) if gs else 1
            except Exception:
                return 1

        def tc_vmerge(tc):
            vm = tc.xpath("./w:tcPr/w:vMerge")
            if not vm:
                return None
            val = vm[0].get(qn("w:val"))
            return val if val else "continue"

        tr_elems = tbl.xpath("./w:tr")

        for r, tr in enumerate(tr_elems):
            tcs = tr.xpath("./w:tc")
            c = 0

            for tc in tcs:
                while c < col_count and matrix[r][c] is not None:
                    c += 1
                if c >= col_count:
                    break

                colspan = max(1, tc_gridspan(tc))
                vmerge = tc_vmerge(tc)

                texts = tc.xpath(".//w:t/text()")
                cell_text = "".join(texts).strip()

                if vmerge == "continue":
                    if r > 0 and matrix[r-1][c] is not None:
                        parent = matrix[r-1][c]
                        parent["rowspan"] += 1
                        for cc in range(c, min(c + colspan, col_count)):
                            matrix[r][cc] = parent
                    else:
                        # continuation без родителя
                        cell_obj = {"text": cell_text, "colspan": colspan, "rowspan": 1, "r0": r, "c0": c}
                        for cc in range(c, min(c + colspan, col_count)):
                            matrix[r][cc] = cell_obj
                    c += colspan
                    continue

                cell_obj = {"text": cell_text, "colspan": colspan, "rowspan": 1, "r0": r, "c0": c}
                for cc in range(c, min(c + colspan, col_count)):
                    matrix[r][cc] = cell_obj
                c += colspan

        lines = ["<table>"]
        for r in range(row_count):
            lines.append("  <tr>")
            c = 0
            while c < col_count:
                cell = matrix[r][c]
                if cell is None:
                    lines.append("    <td></td>")
                    c += 1
                    continue

                if cell["r0"] == r and cell["c0"] == c:
                    attrs = []
                    if cell["colspan"] > 1:
                        attrs.append(f'colspan="{cell["colspan"]}"')
                    if cell["rowspan"] > 1:
                        attrs.append(f'rowspan="{cell["rowspan"]}"')

                    txt = self._escape_html(cell["text"])
                    if attrs:
                        lines.append(f'    <td {" ".join(attrs)}>{txt}</td>')
                    else:
                        lines.append(f"    <td>{txt}</td>")
                    c += cell["colspan"]
                else:
                    c += 1

            lines.append("  </tr>")
        lines.append("</table>")

        result.add_warning(
            f"Таблица {row_count}×{col_count}: использован HTML-режим из-за merged/широкой таблицы"
        )
        return lines
