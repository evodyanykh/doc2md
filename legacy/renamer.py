import os
import glob

class FileRenamer:
    def __init__(self, directory):
        """
        Инициализация класса с указанием директории для обработки
        
        Args:
            directory (str): Путь к папке с файлами
        """
        self.directory = directory
        
    def rename_files(self):
        """
        Проходит по всем файлам в указанной директории и переименовывает файлы,
        заканчивающиеся на ' copy', заменяя на '_new'
        """
        # Проверяем существование директории
        if not os.path.exists(self.directory):
            print(f"Ошибка: директория '{self.directory}' не существует")
            return
            
        if not os.path.isdir(self.directory):
            print(f"Ошибка: '{self.directory}' не является директорией")
            return
        
        # Получаем список всех файлов в директории
        # Можно использовать glob или os.listdir
        
        # Вариант 1: Используем glob для поиска файлов с ' copy' в названии
        pattern = os.path.join(self.directory, "* copy*")
        files_with_copy = glob.glob(pattern)
        
        print(f"Найдено {len(files_with_copy)} файлов с ' copy' в названии")
        
        # Переименовываем файлы
        renamed_count = 0
        for old_path in files_with_copy:
            # Получаем имя файла и директорию
            dir_name = os.path.dirname(old_path)
            filename = os.path.basename(old_path)
            
            # Проверяем, заканчивается ли имя файла на ' copy'
            # (учитываем, что после ' copy' может быть расширение файла)
            if ' copy' in filename:
                # Создаем новое имя файла
                if filename.endswith(' copy'):
                    # Если файл заканчивается точно на ' copy' (без расширения)
                    new_filename = filename.replace(' copy', '_new')
                else:
                    # Если ' copy' где-то в середине имени (например, 'file copy.txt')
                    # Находим последнее вхождение ' copy' и заменяем его
                    parts = filename.rsplit(' copy', 1)
                    if len(parts) == 2:
                        new_filename = parts[0] + '_new' + parts[1]
                    else:
                        continue
                
                # Формируем новый полный путь
                new_path = os.path.join(dir_name, new_filename)
                
                # Переименовываем файл
                try:
                    os.rename(old_path, new_path)
                    print(f"Переименован: {filename} -> {new_filename}")
                    renamed_count += 1
                except Exception as e:
                    print(f"Ошибка при переименовании {filename}: {e}")
        
        print(f"\nИтого переименовано: {renamed_count} файлов")
        
    def rename_all_files_in_directory(self):
        """
        Альтернативный метод: обрабатывает все файлы в директории,
        проверяя каждый на наличие ' copy' в имени
        """
        if not os.path.exists(self.directory):
            print(f"Ошибка: директория '{self.directory}' не существует")
            return
        
        renamed_count = 0
        # Получаем список всех файлов в директории
        for filename in os.listdir(self.directory):
            # Пропускаем директории
            if os.path.isdir(os.path.join(self.directory, filename)):
                continue
                
            # Проверяем, содержит ли имя файла ' copy'
            if ' copy' in filename:
                old_path = os.path.join(self.directory, filename)
                
                # Аналогичная логика замены
                if filename.endswith(' copy'):
                    new_filename = filename.replace(' copy', '_new')
                else:
                    parts = filename.rsplit(' copy', 1)
                    if len(parts) == 2:
                        new_filename = parts[0] + '_new' + parts[1]
                    else:
                        continue
                
                new_path = os.path.join(self.directory, new_filename)
                
                try:
                    os.rename(old_path, new_path)
                    print(f"Переименован: {filename} -> {new_filename}")
                    renamed_count += 1
                except Exception as e:
                    print(f"Ошибка при переименовании {filename}: {e}")
        
        print(f"\nИтого переименовано: {renamed_count} файлов")


# Пример использования
if __name__ == "__main__":
    # Создаем экземпляр класса с указанием пути к папке
    renamer = FileRenamer("lekcii_md")
    
    # Вызываем метод для переименования файлов
    renamer.rename_files()
    
    # Альтернативный вызов
    # renamer.rename_all_files_in_directory()