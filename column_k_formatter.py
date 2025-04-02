import re

class ColumnKFormatter:
    """
    Класс для форматирования текста в столбце K
    """
    
    def process_excel_column(self, sheet, column_index=11):
        """
        Обрабатывает столбец K в Excel файле
        
        Args:
            sheet: Лист Excel для обработки
            column_index: Индекс столбца (по умолчанию 11 для столбца K)
            
        Returns:
            dict: Статистика обработки
        """
        stats = {
            "cells_processed": 0
        }
        
        # Обрабатываем каждую ячейку в столбце
        for row in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            if cell.value:
                # Форматируем текст
                formatted_text = self.format_text(str(cell.value))
                if formatted_text != cell.value:
                    cell.value = formatted_text
                    stats["cells_processed"] += 1
        
        return stats
    
    def format_text(self, text):
        """
        Форматирует текст в столбце K
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Отформатированный текст
        """
        if not text:
            return text
            
        # Удаляем лишние пробелы
        text = re.sub(r'\s+', ' ', text).strip()
        
        # Форматируем даты (например, "30. 01. 2025" -> "30.01.2025")
        text = re.sub(r'(\d{1,2})\.\s+(\d{1,2})\.\s+(\d{2,4})', r'\1.\2.\3', text)
        
        # Добавляем пробелы после знаков препинания
        text = re.sub(r'([.,;:])(?!\s)', r'\1 ', text)
        
        # Убираем пробелы перед знаками препинания
        text = re.sub(r'\s+([.,;:])', r'\1', text)
        
        # Добавляем точку в конце, если её нет
        if text and not text.endswith(('.', '!', '?')):
            text += '.'
        
        # Первая буква предложения должна быть заглавной
        if text:
            text = text[0].upper() + text[1:]
            
        # В самом конце заменяем "СПбским" на "СПб"
        text = re.sub(r'СПбским', 'СПб', text, flags=re.IGNORECASE)
            
        return text 