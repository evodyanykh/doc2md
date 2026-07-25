"""
Логирование ошибок и формирование итогового отчёта о конвертации.

Два выхода:
  1. errors.log  — подробный лог с трассировкой ошибок (для отладки)
  2. report.md   — читаемый отчёт: что обработано, что упало и почему
"""

import logging
import traceback
from datetime import datetime
from pathlib import Path

from formats.config import LOG_FILE, REPORT_FILE


class ConversionLogger:
    """
    Собирает статистику по всем файлам за один запуск
    и записывает итоговый отчёт в report.md.
    """

    def __init__(self, base_dir: Path):
        """
        base_dir — папка, в которой будут созданы errors.log и report.md.
        Обычно это рабочая директория запуска скрипта.
        """
        self.base_dir = base_dir
        self.log_path = base_dir / LOG_FILE
        self.report_path = base_dir / REPORT_FILE

        self.success: list[dict] = []   # успешно обработанные файлы
        self.warnings: list[dict] = []  # обработаны с предупреждениями
        self.errors: list[dict] = []    # файлы с ошибками (пропущены)

        self.started_at = datetime.now()

        # Создаём именованный логгер только для нашего кода.
        # НЕ используем basicConfig — он захватывает root-логгер и тянет
        # DEBUG-сообщения от всех сторонних библиотек (PyMuPDF, pdfplumber и т.д.)
        self._logger = logging.getLogger("converter")
        self._logger.setLevel(logging.INFO)

        # Не добавляем handler повторно при повторном создании логгера
        if not self._logger.handlers:
            handler = logging.FileHandler(
                str(self.log_path), mode="a", encoding="utf-8"
            )
            handler.setFormatter(
                logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
            )
            self._logger.addHandler(handler)

        # Отключаем передачу сообщений в root-логгер (и дальше в сторонние библиотеки)
        self._logger.propagate = False

        self._logger.info(f"=== Новый запуск: {self.started_at.strftime('%Y-%m-%d %H:%M:%S')} ===")

    # ─── Публичный API ────────────────────────────────────────────────────────

    def log_success(self, src: Path, dst: Path) -> None:
        """Файл успешно конвертирован без предупреждений."""
        self.success.append({"src": src, "dst": dst})
        self._logger.info(f"OK: {src} → {dst}")

    def log_warning(self, src: Path, dst: Path, message: str) -> None:
        """
        Файл конвертирован, но с предупреждением
        (например, изображения без текста, сложные таблицы).
        """
        self.warnings.append({"src": src, "dst": dst, "message": message})
        self._logger.warning(f"WARNING: {src} → {dst} | {message}")

    def log_error(self, src: Path, error: Exception) -> None:
        """
        Файл НЕ конвертирован из-за ошибки.
        Записывает полную трассировку в errors.log.
        """
        tb = traceback.format_exc()
        self.errors.append({"src": src, "error": str(error), "traceback": tb})
        self._logger.error(f"ERROR: {src} | {error}\n{tb}")

    def log_skipped(self, src: Path, reason: str) -> None:
        """Файл пропущен намеренно (например, уже существует актуальная версия)."""
        self._logger.info(f"SKIP: {src} | {reason}")

    # ─── Отчёт ───────────────────────────────────────────────────────────────

    def write_report(self) -> None:
        """Записывает итоговый отчёт в report.md после завершения обработки."""
        finished_at = datetime.now()
        duration = finished_at - self.started_at
        total = len(self.success) + len(self.warnings) + len(self.errors)

        lines = [
            f"# Отчёт конвертации",
            f"",
            f"**Запуск:** {self.started_at.strftime('%Y-%m-%d %H:%M:%S')}  ",
            f"**Завершение:** {finished_at.strftime('%Y-%m-%d %H:%M:%S')}  ",
            f"**Время:** {str(duration).split('.')[0]}  ",
            f"",
            f"## Итог",
            f"",
            f"| Статус | Файлов |",
            f"|--------|--------|",
            f"| ✅ Успешно | {len(self.success)} |",
            f"| ⚠️ С предупреждениями | {len(self.warnings)} |",
            f"| ❌ С ошибками | {len(self.errors)} |",
            f"| **Всего** | **{total}** |",
            f"",
        ]

        if self.warnings:
            lines += [
                f"## ⚠️ Предупреждения ({len(self.warnings)})",
                f"",
            ]
            for w in self.warnings:
                lines += [
                    f"- **{w['src'].name}**",
                    f"  - Источник: `{w['src']}`",
                    f"  - Результат: `{w['dst']}`",
                    f"  - Причина: {w['message']}",
                    f"",
                ]

        if self.errors:
            lines += [
                f"## ❌ Ошибки ({len(self.errors)})",
                f"",
            ]
            for e in self.errors:
                lines += [
                    f"- **{e['src'].name}**",
                    f"  - Файл: `{e['src']}`",
                    f"  - Ошибка: `{e['error']}`",
                    f"",
                ]
            lines += [
                f"> Подробная трассировка ошибок — в файле `{LOG_FILE}`",
                f"",
            ]

        if self.success:
            lines += [
                f"## ✅ Успешно обработаны ({len(self.success)})",
                f"",
            ]
            for s in self.success:
                lines += [f"- `{s['src']}` → `{s['dst']}`"]
            lines.append("")

        self.report_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n📄 Отчёт сохранён: {self.report_path}")

    # ─── Вывод в консоль ─────────────────────────────────────────────────────

    def print_summary(self) -> None:
        """Выводит краткую сводку в консоль после завершения."""
        total = len(self.success) + len(self.warnings) + len(self.errors)
        print(f"\n{'='*50}")
        print(f"  Обработано файлов: {total}")
        print(f"  ✅ Успешно:          {len(self.success)}")
        print(f"  ⚠️  С предупреждениями: {len(self.warnings)}")
        print(f"  ❌ С ошибками:       {len(self.errors)}")
        print(f"{'='*50}")
        if self.errors:
            print(f"\n  Файлы с ошибками:")
            for e in self.errors:
                print(f"    - {e['src'].name}: {e['error']}")
        print(f"\n  Подробнее: {self.report_path}")
