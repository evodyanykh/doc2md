"""Тест конвертации PDF. Запуск: python analyze_pdf.py"""
import sys
from pathlib import Path

PDF = r"GR\презентации\Dorofeev_GR_Summit_FIN.pdf"

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from formats.converters.pdf_converter import PdfConverter

result = PdfConverter().convert(Path(PDF))
lines = result.content.splitlines()
print(f"Строк: {len(lines)}  Предупреждений: {len(result.warnings)}\n")

sep = 0
for line in lines:
    print(line)
    if line == "---":
        sep += 1
        if sep >= 3:
            print(f"\n... (всего {len(lines)} строк)")
            break
