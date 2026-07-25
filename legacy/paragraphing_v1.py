import re
from statistics import median
from enum import Enum, auto
from typing import List

from models import TextBlock, Paragraph


class ParagraphBreakReason(Enum):
    Y_GAP = auto()
    X_INDENT = auto()
    SEMANTIC_ANCHOR = auto()
    TITLE_BLOCK = auto()


TERM_PATTERN = re.compile(r"^[А-ЯЁ][А-Яа-яЁё\s]+—")


class ParagraphBuilder:
    """
    Собирает логические абзацы из упорядоченных TextBlock.
    Работает одинаково для PPTX / PDF / OCR.
    """

    def __init__(
        self,
        y_gap_multiplier: float = 1.3,
        x_indent_threshold: float = 20.0,
        force_title_break: bool = True,
        normalize_text: bool = True,
    ):
        self.y_gap_multiplier = y_gap_multiplier
        self.x_indent_threshold = x_indent_threshold
        self.force_title_break = force_title_break
        self.normalize_text = normalize_text

    def build(self, blocks: List[TextBlock]) -> List[Paragraph]:
        if not blocks:
            return []

        # базовые метрики
        heights = [b.height for b in blocks if b.height > 0]
        median_height = median(heights) if heights else 0
        y_gap_threshold = median_height * self.y_gap_multiplier

        paragraphs: List[Paragraph] = []
        current: List[TextBlock] = []

        for prev, curr in zip(blocks, blocks[1:]):
            current.append(prev)

            reason = self._detect_break_reason(
                prev=prev,
                curr=curr,
                y_gap_threshold=y_gap_threshold,
            )

            if reason is not None:
                paragraphs.append(self._flush(current))
                current = []

        current.append(blocks[-1])
        paragraphs.append(self._flush(current))

        return paragraphs

    # -------------------------
    # BREAK LOGIC
    # -------------------------

    def _detect_break_reason(
        self,
        prev: TextBlock,
        curr: TextBlock,
        y_gap_threshold: float,
    ) -> ParagraphBreakReason | None:

        # 1. Заголовок всегда начинает новый абзац
        if self.force_title_break and curr.is_title:
            return ParagraphBreakReason.TITLE_BLOCK

        # 2. Вертикальный разрыв
