from statistics import median
from typing import List
import logging

from models import TextBlock, Paragraph


log = logging.getLogger(__name__)


class ParagraphBuilder:
    """
    Минимальный ParagraphBuilder.
    Делит абзацы ТОЛЬКО по вертикальному разрыву.
    Предназначен для стабилизации и отладки.
    """

    def __init__(self, y_gap_multiplier: float = 1.8):
        self.y_gap_multiplier = y_gap_multiplier

    def build(self, blocks: List[TextBlock]) -> List[Paragraph]:
        if not blocks:
            log.debug("ParagraphBuilder: empty input")
            return []

        heights = [b.height for b in blocks if b.height > 0]
        median_height = median(heights) if heights else 0
        y_gap_threshold = median_height * self.y_gap_multiplier

        log.debug(
            "ParagraphBuilder init: median_height=%.2f y_gap_threshold=%.2f",
            median_height,
            y_gap_threshold,
        )

        paragraphs: List[Paragraph] = []
        current: List[TextBlock] = []

        for prev, curr in zip(blocks, blocks[1:]):
            current.append(prev)

            gap = curr.y - (prev.y + prev.height)

            log.debug(
                "CHECK GAP | prev(y=%.1f h=%.1f) -> curr(y=%.1f) | gap=%.1f",
                prev.y,
                prev.height,
                curr.y,
                gap,
            )

            if gap > y_gap_threshold:
                log.debug(
                    "PARAGRAPH BREAK | gap %.1f > threshold %.1f",
                    gap,
                    y_gap_threshold,
                )
                paragraphs.append(self._flush(current))
                current = []

        current.append(blocks[-1])
        paragraphs.append(self._flush(current))

        log.debug("ParagraphBuilder result: %d paragraphs", len(paragraphs))
        return paragraphs

    def _flush(self, blocks: List[TextBlock]) -> Paragraph:
        text = " ".join(b.text.strip() for b in blocks if b.text.strip())

        log.debug(
            "FLUSH PARAGRAPH | blocks=%d | text=%.60s",
            len(blocks),
            text.replace("\n", " "),
        )

        return Paragraph(
            slide_index=blocks[0].slide_index,
            text=text,
            blocks=blocks,
            is_term_definition=False,
        )
