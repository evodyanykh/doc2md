# main.py
import os
from pathlib import Path
from pptx_batch_processor import PPTXBatchProcessor

if __name__ == "__main__":
    # Создаем папки, если они не существуют
    Path("from").mkdir(exist_ok=True) #lekcii
    Path("to_md").mkdir(exist_ok=True) #lekcii_md
    
    # Проверяем, есть ли файлы для обработки
    pptx_files = list(Path("from").glob("*.pptx")) #lekcii
    
    if not pptx_files:
        print("Поместите pptx файлы в папку 'from'")
        print("Запуск обработки одиночного файла...")
        
        # Обработка одиночного файла (старый функционал)
        from pipeline import PPTXToMarkdownPipeline
        
        pptx_file = "output.pptx"
        output_file = "output.md"

        if not os.path.exists(pptx_file):
            print(f"Файл {pptx_file} не найден!")
            print("Пожалуйста, добавьте файлы в папку 'lekcii' или поместите presentation.pptx в корень")
            exit(1)

        pipeline = PPTXToMarkdownPipeline(pptx_file)
        markdown = pipeline.run()

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(markdown)

        print(f"Markdown сохранён в {output_file}, строк: {len(markdown.splitlines())}")
    else:
        # Обработка всех файлов в папке lekcii
        processor = PPTXBatchProcessor("from", "to_md") #lekcii, lekcii_md
        stats = processor.run()