import re
from typing import List, Tuple

class MarkdownBoldFixer:
    """
    Класс для исправления ошибок форматирования жирного текста в Markdown.
    Исправляет случаи, когда после закрывающих ** отсутствует пробел перед следующим словом.
    """
    
    def __init__(self):
        # Регулярное выражение для поиска **...** конструкций
        # Используем lookbehind и lookahead для проверки границ
        self.bold_pattern = re.compile(r'(?<!\*)\*\*(?!\*)(.*?)(?<!\*)\*\*(?!\*)')
        
    def find_bold_sections(self, text: str) -> List[Tuple[int, int, str]]:
        """
        Находит все жирные секции в тексте.
        Возвращает список кортежей (start, end, content)
        """
        sections = []
        for match in self.bold_pattern.finditer(text):
            start = match.start()
            end = match.end()
            content = match.group(0)
            sections.append((start, end, content))
        return sections
    
    def needs_fix(self, text: str, end_pos: int) -> bool:
        """
        Проверяет, нуждается ли жирная секция в исправлении.
        Исправление нужно, если после закрывающих ** идет символ, который не является:
        - пробелом
        - знаком препинания, за которым обычно следует пробел
        - концом строки
        """
        if end_pos >= len(text):
            return False
        
        next_char = text[end_pos]
        
        # Если следующий символ - пробел, перевод строки или конец текста - не исправляем
        if next_char in (' ', '\n', '\t', '\r'):
            return False
        
        # Если следующий символ - знак препинания, который обычно не требует пробела после
        if next_char in ('.', ',', '!', '?', ';', ':', ')', ']', '}', '"', "'", '`'):
            return False
        
        # Если следующий символ - начало новой жирной секции или другой markdown-синтаксис
        if end_pos + 1 < len(text) and text[end_pos:end_pos+2] in ('**', '*_', '__'):
            return False
        
        return True
    
    def fix_line(self, line: str) -> str:
        """
        Исправляет одну строку текста.
        """
        if '**' not in line:
            return line
        
        result = []
        last_pos = 0
        sections = self.find_bold_sections(line)
        
        for start, end, content in sections:
            # Добавляем текст до текущей секции
            result.append(line[last_pos:start])
            
            # Проверяем, нужно ли исправление
            if self.needs_fix(line, end):
                # Добавляем жирную секцию с пробелом после
                result.append(content + ' ')
                # Обновляем позицию для поиска следующей секции
                last_pos = end
            else:
                # Оставляем как есть
                result.append(content)
                last_pos = end
        
        # Добавляем остаток строки
        result.append(line[last_pos:])
        
        return ''.join(result)
    
    def fix_text(self, text: str) -> str:
        """
        Исправляет многострочный текст.
        """
        lines = text.split('\n')
        fixed_lines = [self.fix_line(line) for line in lines]
        return '\n'.join(fixed_lines)
    
    def fix_file(self, input_file: str, output_file: str = None):
        """
        Читает файл, исправляет его и сохраняет.
        """
        with open(input_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        fixed_content = self.fix_text(content)
        
        if output_file is None:
            output_file = input_file
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(fixed_content)
        
        print(f"Файл исправлен и сохранен как: {output_file}")


# Альтернативная версия с более простым алгоритмом (регулярные выражения)
class MarkdownBoldFixerSimple:
    """
    Упрощенная версия с использованием регулярных выражений.
    """
    
    def __init__(self):
        # Паттерн для поиска **текст**за которым следует слово (без пробела)
        self.fix_pattern = re.compile(r'(\*\*[^*]+\*\*)(?=[A-Za-zА-Яа-я0-9])')
    
    def fix_text(self, text: str) -> str:
        """
        Исправляет текст, добавляя пробелы после **, за которыми сразу идет текст.
        """
        def add_space(match):
            return match.group(1) + ' '
        
        # Применяем замену
        return self.fix_pattern.sub(add_space, text)


# Пример использования
if __name__ == "__main__":
    # Пример текста с ошибками
    test_text = """
    Вот пример текста с ошибками:
    **Результат:**Сформированная картина
    **Важно:**Обратите внимание
    **Примечание:**Это текст без пробела
    Правильный пример: **Вот так:** с пробелом
    Еще пример **жирный текст**и сразу продолжение
    **Не исправляем:**!восклицательный знак
    **Не исправляем:**.точка
    **Не исправляем:**?вопрос
    """
    
    print("Исходный текст:")
    print(test_text)
    print("\n" + "="*50 + "\n")
    
    # Используем основной класс
    fixer = MarkdownBoldFixer()
    fixed_text = fixer.fix_text(test_text)
    
    print("Исправленный текст:")
    print(fixed_text)
    
    # Или используем упрощенную версию
    print("\n" + "="*50 + "\n")
    simple_fixer = MarkdownBoldFixerSimple()
    simple_fixed = simple_fixer.fix_text(test_text)
    print("Исправленный текст (простой алгоритм):")
    print(simple_fixed)
    
    # Для работы с файлами:
    fixer.fix_file('input.md', 'output.md')