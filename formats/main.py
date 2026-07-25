"""
Точка входа — CLI для конвертации файлов в Markdown.

Использование:
  python main.py <папка>

Примеры:
  python main.py C:/Documents/materials
  python main.py .                          # текущая папка

Что делает:
  1. Рекурсивно сканирует указанную папку
  2. Находит все .docx, .pptx, .pdf файлы
  3. Конвертирует каждый файл в .md с версионированием
  4. Сохраняет .md рядом с исходным файлом
  5. Пишет отчёт report.md и errors.log в папку запуска
"""

import sys
from pathlib import Path

# Позволяет запускать и как «python formats/main.py», и как «python -m formats.main»:
# при прямом запуске добавляем корень проекта в sys.path, чтобы пакет formats находился
if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from formats.converters import CONVERTERS, get_converter
from formats.utils.logger import ConversionLogger
from formats.utils.ocr import is_tesseract_available
from formats.utils.scanner import scan_directory, print_scan_summary, group_by_folder
from formats.utils.versioning import get_output_path

# Консоль Windows по умолчанию работает в cp1251/cp866 и падает на "→" и эмодзи —
# переводим вывод в UTF-8 (символы без глифа заменяются, но скрипт не падает)
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    # ─── Аргументы командной строки ───────────────────────────────────────────
    if len(sys.argv) < 2:
        print("Использование: python main.py <папка>")
        print("Пример:        python main.py C:/Documents/materials")
        sys.exit(1)

    root = Path(sys.argv[1]).resolve()

    print(f"\n{'='*50}")
    print(f"  Конвертер документов → Markdown")
    print(f"{'='*50}")
    print(f"  Папка: {root}")

    # ─── Проверка Tesseract ────────────────────────────────────────────────────
    tesseract_ok = is_tesseract_available()
    if tesseract_ok:
        print("  OCR (Tesseract): доступен ✅")
    else:
        print("  OCR (Tesseract): недоступен ⚠️  — сканы и изображения будут пропущены")
        print("  Установите Tesseract: https://github.com/UB-Mannheim/tesseract/wiki")

    # ─── Сканирование папки ───────────────────────────────────────────────────
    print(f"\n  Сканирую папку...")
    try:
        files = scan_directory(root)
    except (FileNotFoundError, NotADirectoryError) as e:
        print(f"\n  Ошибка: {e}")
        sys.exit(1)

    print_scan_summary(files)

    if not files:
        print("\n  Нечего конвертировать. Завершение.")
        sys.exit(0)

    # ─── Инициализация логгера ────────────────────────────────────────────────
    # Логи создаются в папке запуска скрипта (не в папке документов)
    script_dir = Path(__file__).parent
    logger = ConversionLogger(script_dir)

    # ─── Обработка файлов ─────────────────────────────────────────────────────
    print(f"\n  Начинаю конвертацию...\n")

    groups = group_by_folder(files)
    total = len(files)
    processed = 0

    for folder, folder_files in groups.items():
        # Выводим название папки для удобства отслеживания прогресса
        print(f"  📁 {folder}")

        for file_path in folder_files:
            processed += 1
            ext = file_path.suffix.lower()

            print(f"     [{processed}/{total}] {file_path.name}", end=" ... ", flush=True)

            # Получаем путь для выходного файла с версионированием
            output_path = get_output_path(file_path)

            # Выбираем нужный конвертер по расширению
            converter = get_converter(ext)
            if converter is None:
                # Не должно случиться (scanner фильтрует по SUPPORTED_EXTENSIONS),
                # но на всякий случай
                logger.log_skipped(file_path, f"Неизвестное расширение: {ext}")
                print("пропущен")
                continue

            try:
                # Конвертируем файл
                result = converter.convert(file_path)

                # Записываем результат
                output_path.write_text(result.content, encoding="utf-8")

                if result.has_warnings():
                    logger.log_warning(file_path, output_path, "; ".join(result.warnings))
                    print(f"⚠️  ({len(result.warnings)} предупр.)")
                else:
                    logger.log_success(file_path, output_path)
                    print("✅")

            except Exception as e:
                logger.log_error(file_path, e)
                print(f"❌ ОШИБКА: {e}")

        print()  # пустая строка между папками

    # ─── Итоги ────────────────────────────────────────────────────────────────
    logger.print_summary()
    logger.write_report()


if __name__ == "__main__":
    main()
