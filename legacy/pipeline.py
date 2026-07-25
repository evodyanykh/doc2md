from models import Document, SlideContent, TextBlock
from parser import PPTXParser
from ordering import LayoutAnalyzer
from renderer import MarkdownRenderer
from paragraphing import ParagraphBuilder


class PPTXToMarkdownPipeline:
    """
    Оркестратор процесса конвертации PPTX → Markdown.

    Принципы:
    - Заголовок слайда = первый текстовый блок после сортировки
    - Внутренние заголовки НЕ определяем (это сделает модель)
    - Максимально стабильный и простой пайплайн
    """

    def __init__(self, pptx_path: str):
        self.parser = PPTXParser(pptx_path)
        self.layout = LayoutAnalyzer()
        self.renderer = MarkdownRenderer()
        self.paragraph_builder = ParagraphBuilder()

    def run(self) -> str:
        """
        Запускает полный пайплайн:
        PPTX → SlideContent → Document → Markdown
        """
        raw_slides = self.parser.parse()
        slides: list[SlideContent] = []

        for slide in raw_slides:
            # 1. Определяем порядок чтения + колонки
            ordered_blocks: list[TextBlock] = self.layout.analyze(slide.blocks)

            if not ordered_blocks:
                slides.append(
                    SlideContent(
                        slide_index=slide.slide_index,
                        title=None,
                        blocks=[]
                    )
                )
                continue

            # 2. Первый блок = заголовок слайда
            title_block = ordered_blocks[0]
            title = title_block.text.strip()

            # 3. Остальные блоки = тело слайда
            content_blocks = ordered_blocks[1:]
            
            #paragraphs = self.paragraph_builder.build(content_blocks)

            slides.append(
                SlideContent(
                    slide_index=slide.slide_index,
                    title=title,
                    blocks=content_blocks
                )
            )

        document = Document(slides=slides)
        return self.renderer.render(document)
