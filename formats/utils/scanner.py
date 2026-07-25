"""
Рекурсивный обход папок для поиска файлов поддерживаемых форматов.

Проходит все вложенные директории и собирает список файлов,
отфильтрованных по расширению из config.SUPPORTED_EXTENSIONS.
"""

from pathlib import Path

from formats.config import SUPPORTED_EXTENSIONS


def scan_directory(root: Path) -> list[Path]:
    """
    Рекурсивно обходит root и возвращает список всех файлов
    с расширениями из SUPPORTED_EXTENSIONS.

    Файлы сортируются: сначала по папке, затем по имени — для
    предсказуемого порядка обработки и читаемого лога.

    Args:
        root: корневая папка для обхода.

    Returns:
        Список путей к файлам в порядке обхода.
    """
    if not root.exists():
        raise FileNotFoundError(f"Папка не найдена: {root}")

    if not root.is_dir():
        raise NotADirectoryError(f"Указан файл, а не папка: {root}")

    found: list[Path] = []

    # rglob("*") обходит все вложенные папки рекурсивно
    for path in sorted(root.rglob("*")):
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            found.append(path)

    return found


def group_by_folder(files: list[Path]) -> dict[Path, list[Path]]:
    """
    Группирует файлы по родительской папке.
    Удобно для вывода прогресса и отчёта.

    Args:
        files: список путей к файлам.

    Returns:
        Словарь {папка: [файл1, файл2, ...]}.
    """
    groups: dict[Path, list[Path]] = {}
    for f in files:
        folder = f.parent
        groups.setdefault(folder, []).append(f)
    return groups


def print_scan_summary(files: list[Path]) -> None:
    """
    Выводит в консоль итог сканирования:
    сколько файлов найдено и разбивку по форматам.
    """
    if not files:
        print("  Файлы для обработки не найдены.")
        return

    # Подсчёт по расширениям
    counts: dict[str, int] = {}
    for f in files:
        ext = f.suffix.lower()
        counts[ext] = counts.get(ext, 0) + 1

    print(f"  Найдено файлов: {len(files)}")
    for ext, count in sorted(counts.items()):
        print(f"    {ext}: {count}")
