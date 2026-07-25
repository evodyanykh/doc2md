# Legacy — прежний PPTX-конвейер и разовые утилиты

Этот код **не используется** основным приложением (`app.py`, `formats/`).
Сохранён для истории и как база для экспериментов.

## Прежний PPTX → Markdown конвейер (на «сыром» XML)

Первая версия конвертера: pptx разбирался вручную как zip-архив с XML,
без библиотеки python-pptx.

- `main.py` — точка входа (папки `from/` → `to_md/`)
- `pptx_batch_processor.py` — пакетная обработка папки
- `pipeline.py` — оркестратор: parser → ordering → renderer
- `parser.py` / `parser_v1.py` — разбор slideX.xml (текст, координаты, z-order)
- `extract.py` — извлечение фрагментов текста с признаком «жирный»
- `ordering.py` — определение колонок и порядка чтения блоков
- `grouping.py`, `paragraphing.py` / `paragraphing_v1.py` — группировка блоков в абзацы
- `models.py` — датаклассы TextBlock / SlideContent / Document
- `renderer.py` — сборка Markdown

Чем он слабее текущего `formats/converters/pptx_converter.py`:
не извлекает таблицы и изображения (нет OCR), не видит курсив,
заголовком слайда считает просто первый по порядку блок
(а не placeholder «title»), не различает уровни списков.

## Разовые утилиты

- `fixer.py` — исправление ошибок жирного форматирования в готовом Markdown
- `renamer.py` — массовое переименование файлов в папке
- `split.py` — вырезание диапазона страниц из PDF
- `extract_slides.py` — извлечение отдельных слайдов из PPTX в новый файл
