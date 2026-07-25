"""
Версионирование выходных .md файлов.

Имя результата включает исходный формат: document_pptx_v1.md.
Это исключает конфликт, когда рядом лежат document.pptx и document.pdf —
раньше оба писали в одну цепочку document_v1.md, document_v2.md,
и по имени было не понять, из какого файла получен какой Markdown.

Правило: если document_pptx_v1.md уже существует — создаём document_pptx_v2.md и т.д.
Выходной файл всегда располагается в той же папке, что и исходный.
"""

from pathlib import Path

from formats.config import VERSION_SUFFIX


def _base_name(source: Path) -> str:
    """Базовое имя результата: 'лекция.pptx' → 'лекция_pptx'."""
    fmt = source.suffix.lstrip(".").lower()
    return f"{source.stem}_{fmt}"


def get_output_path(source: Path) -> Path:
    """
    Определяет путь для выходного .md файла с учётом версионирования.

    Алгоритм:
      1. Базовое имя = имя исходника + формат: "document_pptx"
      2. Пробуем "document_pptx_v1.md" — если не существует, возвращаем его
      3. Иначе пробуем _v2, _v3 и т.д. до первого свободного имени

    Args:
        source: путь к исходному файлу (docx / pptx / pdf).

    Returns:
        Путь к выходному .md файлу (ещё не существующему).
    """
    folder = source.parent
    base = _base_name(source)

    version = 1
    while True:
        candidate = folder / f"{base}{VERSION_SUFFIX}{version}.md"
        if not candidate.exists():
            return candidate
        version += 1


def get_latest_version(source: Path) -> Path | None:
    """
    Находит последнюю существующую версию .md файла для данного источника.
    Возвращает None, если ни одной версии ещё нет.

    Используется для проверки: нужно ли вообще конвертировать файл заново.
    """
    folder = source.parent
    base = _base_name(source)

    latest: Path | None = None
    version = 1

    while True:
        candidate = folder / f"{base}{VERSION_SUFFIX}{version}.md"
        if candidate.exists():
            latest = candidate
            version += 1
        else:
            break

    return latest
