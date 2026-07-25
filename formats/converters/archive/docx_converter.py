"""
Конвертер DOCX → Markdown.

Сохраняет форматирование:
  - Заголовки (Heading 1-6) → # ## ### и т.д.
  - Жирный текст → **текст**
  - Курсив → *текст*
  - Жирный + курсив → ***текст***
  - Маркированные списки → - пункт
  - Нумерованные списки → 1. пункт (с учётом уровней вложенности)
  - Гиперссылки → [текст](url)
  - Таблицы → Markdown-таблица или HTML при объединённых ячейках
  - Изображения → полностью игнорируются (не сохраняются)
"""

import logging
from pathlib import Path
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Mm
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from config import IMAGE_PLACEHOLDER, TABLE_WARNING
from converters.base import BaseConverter, ConversionResult
from utils.ocr import ocr_image_bytes, is_tesseract_available

# Настройка логирования
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    stream=sys.stderr
)
logger = logging.getLogger(__name__)


class DocxConverter(BaseConverter):

    def convert(self, file_path: Path) -> ConversionResult:
        logger.info(f"Начало конвертации файла: {file_path}")
        result = ConversionResult()
        self.document = Document(str(file_path))

        lines: list[str] = []

        for child in self.document.element.body:
            tag = child.tag.split("}")[-1]

            if tag == "p":
                para = Paragraph(child, self.document)
                md_line = self._convert_paragraph(para, result)
                if md_line is not None:
                    lines.append(md_line)

            elif tag == "tbl":
                table = Table(child, self.document)
                md_table = self._convert_table(table, result)
                lines.append("")
                lines.extend(md_table)
                lines.append("")

        result.content = self._clean_text("\n".join(lines))
        logger.info("Конвертация завершена")
        return result

    def _clean_text(self, text: str) -> str:
        lines = text.splitlines()
        cleaned = []
        prev_empty = False
        for line in lines:
            line = line.rstrip()
            if not line:
                if not prev_empty:
                    cleaned.append('')
                    prev_empty = True
            else:
                cleaned.append(line)
                prev_empty = False
        return '\n'.join(cleaned)

    # ─── Параграфы ────────────────────────────────────────────────────────────

    def _convert_paragraph(self, para: Paragraph, result: ConversionResult) -> str | None:
        style_name = para.style.name if para.style else ""
        logger.debug(f"Параграф: '{para.text[:50]}...', стиль: '{style_name}'")

        if self._is_caption_style(style_name):
            logger.debug("  → пропущен (подпись)")
            return None

        image_lines = self._extract_images(para, result)

        text = self._convert_runs(para)
        logger.debug(f"  текст после runs: '{text[:50]}...'")

        if not text.strip() and image_lines:
            return "\n".join(image_lines)

        if not text.strip():
            logger.debug("  → пустой, пропускаем")
            return ""

        styled = self._apply_style(text, style_name, para)
        logger.debug(f"  → результат: '{styled}'")
        return styled

    def _is_caption_style(self, style_name: str) -> bool:
        if not style_name:
            return False
        low = style_name.lower()
        return "caption" in low or "подпись" in low

    def _convert_runs(self, para: Paragraph) -> str:
        parts = []

        for child in para._element.iterchildren():
            tag = child.tag.split("}")[-1]

            if tag == "r":
                run = Run(child, para)
                text = run.text
                if not text:
                    continue

                bold = run.bold
                italic = run.italic

                if text.strip():
                    if bold and italic:
                        text = f"***{text}***"
                    elif bold:
                        text = f"**{text}**"
                    elif italic:
                        text = f"*{text}*"

                parts.append(text)

            elif tag == "hyperlink":
                r_id = child.get(qn("r:id"))
                url = None
                if r_id:
                    rel = para.part.rels.get(r_id)
                    if rel:
                        url = rel.target_ref

                link_text_parts = []
                for subchild in child.iterchildren():
                    if subchild.tag.split("}")[-1] == "r":
                        run = Run(subchild, para)
                        link_text_parts.append(run.text)
                link_text = "".join(link_text_parts).strip()

                if url and link_text:
                    parts.append(f"[{link_text}]({url})")
                elif url:
                    parts.append(f"[{url}]({url})")
                else:
                    parts.append(link_text)

        return "".join(parts)

    def _get_list_info(self, para: Paragraph):
        """
        Определяет, является ли параграф частью списка, и если да, то тип и уровень.
        Возвращает (is_list, is_numbered, level).
        """
        num_pr = para._element.find(qn("w:numPr"))
        if num_pr is None and para.style:
            num_pr = para.style.element.find(qn("w:pPr/w:numPr"))

        if num_pr is None:
            return False, False, 0

        ilvl_elem = num_pr.find(qn("w:ilvl"))
        level = 0
        if ilvl_elem is not None:
            try:
                level = int(ilvl_elem.get(qn("w:val"), 0))
            except:
                level = 0

        num_id_elem = num_pr.find(qn("w:numId"))
        if num_id_elem is None:
            return True, False, level

        num_id = num_id_elem.get(qn("w:val"))

        try:
            numbering_part = self.document.part.numbering_part
            if numbering_part is None:
                return True, True, level

            numbering = numbering_part.numbering_definitions
            num = None
            for n in numbering._num_lst:
                if str(n.num_id) == str(num_id):
                    num = n
                    break
            if num is None:
                return True, True, level

            abstract_num_id = num.abstract_num_id
            abstract_num = None
            for an in numbering._abstract_num_lst:
                if str(an.abstract_num_id) == str(abstract_num_id):
                    abstract_num = an
                    break

            if abstract_num is None:
                return True, True, level

            lvl = None
            for lvl_elem in abstract_num.element.xpath(f".//w:lvl[@w:ilvl='{level}']"):
                lvl = lvl_elem
                break

            if lvl is not None:
                num_fmt = lvl.find(qn("w:numFmt"))
                if num_fmt is not None:
                    fmt_val = num_fmt.get(qn("w:val"))
                    return True, (fmt_val != "bullet"), level

            return True, True, level
        except Exception:
            return True, True, level

    def _get_list_indent(self, para: Paragraph) -> int:
        try:
            num_pr = para._element.find(qn("w:numPr"))
            if num_pr is None and para.style:
                num_pr = para.style.element.find(qn("w:pPr/w:numPr"))
            if num_pr is not None:
                ilvl = num_pr.find(qn("w:ilvl"))
                if ilvl is not None:
                    return int(ilvl.get(qn("w:val"), 0))
        except:
            pass

        if para.paragraph_format and para.paragraph_format.left_indent:
            indent = para.paragraph_format.left_indent
            if isinstance(indent, Cm):
                level = int(indent.cm // 1.0)
            elif isinstance(indent, Mm):
                level = int(indent.mm // 10)
            else:
                level = 0
            return min(level, 5)
        return 0

    def _apply_style(self, text: str, style_name: str, para: Paragraph) -> str:
        # Заголовки
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

        # Получаем информацию о списке
        list_info = self._get_list_info(para)
        # Проверяем, что возвращено 3 значения, и распаковываем
        if len(list_info) == 3:
            is_list, is_numbered, xml_level = list_info
        else:
            # fallback для старых версий (на всякий случай)
            is_list, xml_level = list_info  # предполагаем, что возвращается два
            is_numbered = True  # консервативно считаем нумерованным

        indent_level = self._get_list_indent(para)
        level = max(xml_level, indent_level)
        indent = "  " * level

        if is_list:
            if is_numbered:
                return f"{indent}1. {text}"
            else:
                return f"{indent}- {text}"

        # Эвристика по началу текста
        stripped = text.lstrip()
        if re.match(r'^\d+\.', stripped):
            return f"{indent}1. {text}"
        if stripped.startswith(('-', '•', '*', '·')):
            return f"{indent}- {text}"

        return text

    # ─── Изображения (игнорируются) ───────────────────────────────────────────

    def _extract_images(self, para: Paragraph, result: ConversionResult) -> list[str]:
        return []

    # ─── Таблицы (полная версия) ─────────────────────────────────────────────

    def _escape_md_cell(self, text: str) -> str:
        text = (text or "").replace("\r", "").strip()
        text = text.replace("\n", "<br>")
        text = text.replace("|", "\\|")
        text = re.sub(r"\s*<br>\s*", "<br>", text)
        return text

    def _escape_html(self, text: str) -> str:
        text = (text or "").replace("\r", "").strip()
        text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        text = text.replace('"', "&quot;")
        text = text.replace("\n", "<br>")
        return text

    def _cell_text(self, cell) -> str:
        parts = []
        for para in cell.paragraphs:
            if para.text.strip():
                parts.append(self._convert_runs(para))
        return "\n".join([p for p in parts if p is not None]).strip()

    def _table_has_merges(self, table: Table) -> bool:
        tbl = table._tbl
        return bool(tbl.xpath(".//w:tcPr/w:gridSpan | .//w:tcPr/w:vMerge"))

    def _convert_table(self, table: Table, result: ConversionResult) -> list[str]:
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