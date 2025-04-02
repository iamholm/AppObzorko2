import re

class FinalDateFormatter:
    """
    Класс для финальной обработки дат в столбце K после всех других преобразований
    """
    
    @staticmethod
    def process_dates_in_column_k(sheet):
        """
        Форматирует даты в столбце K, удаляя пробелы между числами в датах
        
        Args:
            sheet: Лист Excel
            
        Returns:
            int: Количество отформатированных дат
        """
        formatted_count = 0
        
        # Обрабатываем все строки
        for row in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=11)  # Столбец K
            
            if cell.value:
                original_text = str(cell.value)
                
                # Заменяем даты с пробелами на даты без пробелов
                # Например: "13. 05. 2023" -> "13.05.2023"
                modified_text = re.sub(r'(\d{1,2})\.\s+(\d{1,2})\.\s+(\d{2,4})', r'\1.\2.\3', original_text)
                
                # Повторяем замену, пока все даты не будут обработаны
                iterations = 0
                while re.search(r'(\d{1,2})\.\s+(\d{1,2})\.\s+(\d{2,4})', modified_text) and iterations < 5:
                    modified_text = re.sub(r'(\d{1,2})\.\s+(\d{1,2})\.\s+(\d{2,4})', r'\1.\2.\3', modified_text)
                    iterations += 1
                
                # Если текст изменился, обновляем ячейку
                if modified_text != original_text:
                    cell.value = modified_text
                    formatted_count += 1
        
        return formatted_count 