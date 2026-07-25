import os
from pathlib import Path
from pipeline import PPTXToMarkdownPipeline

class PPTXBatchProcessor:
    """
    Класс для пакетной обработки pptx файлов из папки lekcii 
    и сохранения их в формате md в папку lekcii_md
    """
    
    def __init__(self, input_dir, output_dir):
        """
        Инициализация процессора
        
        Args:
            input_dir (str): Папка с исходными pptx файлами
            output_dir (str): Папка для сохранения md файлов
        """
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        
        # Создаем выходную папку, если она не существует
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def get_pptx_files(self):
        """
        Получение списка pptx файлов во входной папке
        
        Returns:
            list: Список путей к pptx файлам
        """
        return list(self.input_dir.glob("*.pptx"))
    
    def process_file(self, pptx_path):
        """
        Обработка одного pptx файла
        
        Args:
            pptx_path (Path): Путь к pptx файлу
            
        Returns:
            tuple: (успех_обработки, количество_строк, сообщение)
        """
        try:
            # Создаем имя для выходного файла
            output_name = pptx_path.stem + ".md"
            output_path = self.output_dir / output_name
            
            # Запускаем конвейер обработки
            pipeline = PPTXToMarkdownPipeline(str(pptx_path))
            markdown = pipeline.run()
            
            # Сохраняем результат
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(markdown)
            
            line_count = len(markdown.splitlines())
            return True, line_count, f"Успешно: {pptx_path.name} -> {output_name}"
            
        except FileNotFoundError:
            return False, 0, f"Ошибка: Файл {pptx_path.name} не найден"
        except Exception as e:
            return False, 0, f"Ошибка при обработке {pptx_path.name}: {str(e)}"
    
    def run(self, verbose=True):
        """
        Запуск пакетной обработки всех файлов
        
        Args:
            verbose (bool): Выводить ли подробную информацию
            
        Returns:
            dict: Статистика обработки
        """
        pptx_files = self.get_pptx_files()
        
        if not pptx_files:
            print(f"Не найдено pptx файлов в папке {self.input_dir}")
            return {"total": 0, "success": 0, "failed": 0, "total_lines": 0}
        
        if verbose:
            print(f"Найдено файлов для обработки: {len(pptx_files)}")
            print("-" * 50)
        
        stats = {
            "total": len(pptx_files),
            "success": 0,
            "failed": 0,
            "total_lines": 0
        }
        
        for pptx_file in pptx_files:
            success, lines, message = self.process_file(pptx_file)
            
            if success:
                stats["success"] += 1
                stats["total_lines"] += lines
            else:
                stats["failed"] += 1
            
            if verbose:
                status = "✓" if success else "✗"
                print(f"{status} {message}")
        
        if verbose:
            print("-" * 50)
            print(f"Обработка завершена:")
            print(f"  Всего файлов: {stats['total']}")
            print(f"  Успешно: {stats['success']}")
            print(f"  С ошибками: {stats['failed']}")
            print(f"  Всего строк в MD: {stats['total_lines']}")
            print(f"  Результаты сохранены в: {self.output_dir}")
        
        return stats


if __name__ == "__main__":
    # Пример использования
    processor = PPTXBatchProcessor(
        input_dir="lekcii",
        output_dir="lekcii_md"
    )
    
    # Запуск обработки
    stats = processor.run()