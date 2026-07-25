import os
from pipeline import PPTXToMarkdownPipeline

if __name__ == "__main__":
    pptx_file = "presentation.pptx"
    output_file = "presentation.md"

    if not os.path.exists(pptx_file):
        print(f"Файл {pptx_file} не найден!")
        exit(1)

    pipeline = PPTXToMarkdownPipeline(pptx_file)
    markdown = pipeline.run()

    with open(output_file, "w", encoding="utf-8") as f:
        f.write(markdown)

    print(f"Markdown сохранён в {output_file}, строк: {len(markdown.splitlines())}")
