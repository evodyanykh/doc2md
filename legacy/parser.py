import zipfile
import xml.etree.ElementTree as ET
import re
from models import SlideContent, TextBlock
from extract import extract_text

EMU_TO_PX = 1 / 914400  # перевод EMU → px


class PPTXParser:
    """
    Низкоуровневый парсер PPTX через zip и XML.

    Задачи:
    - открыть pptx как zip
    - определить порядок слайдов
    - извлечь текстовые блоки с координатами и z-order
    """

    def __init__(self, pptx_path: str):
        self.pptx_path = pptx_path
        self.zip = zipfile.ZipFile(pptx_path, "r")

    def parse(self):
        """
        Парсит все слайды в корректном порядке.
        Возвращает список объектов Slide.
        """
        slides = []
        slide_paths = self._get_slide_paths_safe()

        for idx, slide_path in enumerate(slide_paths, start=1):
            xml_bytes = self.zip.read(slide_path)
            slides.append(self._parse_slide(xml_bytes, idx))

        return slides
    
    def _get_slide_paths_safe(self):
        """
        Определяет список slideX.xml для всех слайдов.
        Если relationships rId не совпадают — используем fallback на все файлы.
        """
        try:
            pres_xml = self.zip.read("ppt/presentation.xml")
            rels_xml = self.zip.read("ppt/_rels/presentation.xml.rels")

            pres_root = ET.fromstring(pres_xml)
            rels_root = ET.fromstring(rels_xml)

            ns_pres = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
            ns_rel = {"r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

            rels = {
                rel.attrib["Id"]: rel.attrib["Target"]
                for rel in rels_root.findall(".//Relationship")
                if rel.attrib.get("Type", "").endswith("/slide")
            }

            slide_paths = []
            for sld in pres_root.findall(".//p:sldId", ns_pres):
                rId = sld.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                slide_target = rels.get(rId)
                if slide_target:
                    slide_target = slide_target.replace("../", "")
                    slide_paths.append("ppt/" + slide_target)

            if not slide_paths:
                raise Exception("rId-based slide order пустой, fallback")

            return slide_paths

        except Exception:
            slide_files = [f for f in self.zip.namelist() if re.match(r"ppt/slides/slide\d+\.xml", f)]
            slide_files.sort(key=lambda x: int(re.findall(r"slide(\d+)\.xml", x)[0]))
            return slide_files

    def _parse_slide(self, xml_bytes: bytes, index: int):
        """
        Парсит один slideX.xml.
        Возвращает SlideContent:
        - slide_index
        - title (строка или None)
        - blocks (TextBlock без заголовков)
        """
        root = ET.fromstring(xml_bytes)

        ns = {
            "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        }

        blocks = []
        slide_title = None

        for z_index, shape in enumerate(root.findall(".//p:sp", ns)):
            # 🔴 extract возвращает список (text, is_bold) на уровне run
            text_fragments = extract_text(self, shape, ns)
            if not text_fragments:
                continue

            # объединяем фрагменты в один текст для совместимости
            combined_text = "".join(frag[0] for frag in text_fragments)

            x, y, w, h = self._extract_geometry(shape, ns)
            blocks.append(
                TextBlock(
                    slide_index=index,
                    text=combined_text,          # объединённый текст для удобства
                    fragments=text_fragments,    # список фрагментов: (text, is_bold, is_list_item)
                    x=x,
                    y=y,
                    width=w,
                    height=h,
                    z=z_index,
                    is_title=False,              # оставляем, если нужно для рендера
                    # is_bold больше не нужен!
                )
            )

            slide_title = self._detect_slide_title(blocks)

        return SlideContent(slide_index=index, title=slide_title, blocks=blocks)


    def _extract_geometry(self, shape, ns):
        """
        Безопасно извлекает координаты и размеры shape.
        EMU → px.
        """
        off = shape.find(".//a:off", ns)
        ext = shape.find(".//a:ext", ns)

        x = float(off.attrib.get("x", 0)) * EMU_TO_PX if off is not None else 0
        y = float(off.attrib.get("y", 0)) * EMU_TO_PX if off is not None else 0
        w = float(ext.attrib.get("cx", 0)) * EMU_TO_PX if ext is not None else 0
        h = float(ext.attrib.get("cy", 0)) * EMU_TO_PX if ext is not None else 0

        return x, y, w, h

    def _is_slide_title(self, shape, ns) -> bool:
        """
        Определяет, является ли shape заголовком слайда
        (placeholder title или ctrTitle)
        """
        ph = shape.find(".//p:ph", ns)

        if ph is None:
            return False

        ph_type = ph.attrib.get("type", "")
        return ph_type in ("title", "ctrTitle")

    def _detect_slide_title(self, blocks: list[TextBlock]) -> str | None:
        """
        Заголовок слайда:
        - самый верхний текстовый блок
        - короткий
        """
        if not blocks:
            return None
        return blocks[0].text
