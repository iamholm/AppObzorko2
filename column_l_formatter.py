import re

class ColumnLFormatter:
    """
    Класс для форматирования текста в столбце L (обязанности)
    """
    
    def process_excel_column(self, sheet, column_index=12):
        """
        Обрабатывает все ячейки в указанном столбце
        
        Args:
            sheet: Лист Excel
            column_index (int): Индекс столбца (по умолчанию 12 для столбца L)
            
        Returns:
            dict: Статистика обработки
        """
        stats = {
            'cells_processed': 0
        }
        
        for row in range(1, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            if cell.value:
                original_text = str(cell.value)
                formatted_text = self.format_text(original_text)
                
                if formatted_text != original_text:
                    cell.value = formatted_text
                    stats['cells_processed'] += 1
                    
            # После всех форматирований проверяем столбец K на двойные точки
            if column_index == 12:  # Если обрабатываем столбец L
                cell_k = sheet.cell(row=row, column=11)  # Столбец K
                if cell_k.value:
                    text = str(cell_k.value)
                    # Удаляем все точки в конце
                    while text.endswith('.'):
                        text = text[:-1].strip()
                    # Добавляем одну точку
                    text = text.strip() + '.'
                    cell_k.value = text
        
        return stats
    
    def _normalize_special_terms(self, text):
        """
        Нормализует специальные термины в тексте
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Текст с нормализованными терминами
        """
        # Замена специальных терминов
        text = re.sub(r'ПМЖ', 'постоянное место жительства', text, flags=re.IGNORECASE)
        text = re.sub(r'ПЖМ', 'постоянное место жительства', text, flags=re.IGNORECASE)
        
        # Замена " р." на " раз"
        text = re.sub(r'\sр\.', ' раз', text)
        
        # Замена " м." на " мес."
        text = re.sub(r'\sм\.', ' мес.', text)
        
        # Замена "1 /мес" на "1 раз в мес"
        text = re.sub(r'(\d+)\s*/мес', r'\1 раз в мес', text)
        
        # Замена "м/ж" на "мест. жительства"
        text = re.sub(r'м/ж', 'мест. жительства', text, flags=re.IGNORECASE)
        
        # Замена " мж" на " место жительства"
        text = re.sub(r'\sмж', ' место жительства', text, flags=re.IGNORECASE)
        
        # Замена " уии" на " УИИ"
        text = re.sub(r'\sуии', ' УИИ', text, flags=re.IGNORECASE)
        
        # Замена специальных слов на полные формулировки
        replacements = {
            'Интернет': 'запрет на использование интернета',
            'Выход из дома': 'ограничение на выход из дома',
            'Общение': 'запрет на общение с определёнными лицами',
            'Связь': 'запрет на использование средств связи'
        }
        
        for old, new in replacements.items():
            text = re.sub(rf'\b{old}\b', new, text, flags=re.IGNORECASE)
        
        return text
        
    def _process_hyphens(self, text):
        """
        Обрабатывает дефисы в тексте:
        1. Если после "-" нет пробела, добавляет его
        2. Если перед "-" есть пробел, но до пробела нет "," или ";", добавляет ";"
        3. Затем убирает все "- " (дефис с пробелом)
        4. В конце убирает все оставшиеся дефисы "-"
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Текст с обработанными дефисами
        """
        # Добавляем пробел после "-" если его нет
        text = re.sub(r'-(?!\s)', '- ', text)
        
        # Добавляем ";" перед пробелом и "-" если перед пробелом нет "," или ";"
        text = re.sub(r'(?<![,;])\s+-', '; -', text)
        
        # Убираем все "- " (дефис с пробелом)
        text = text.replace('- ', '')
        
        # Убираем все оставшиеся дефисы
        text = text.replace('-', '')
        
        return text

    def format_text(self, text):
        """
        Форматирует текст в столбце L
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Отформатированный текст
        """
        if not text:
            return text
            
        # Удаляем лишние пробелы
        text = re.sub(r'\s+', ' ', text).strip()
        
        # Нормализуем специальные термины
        text = self._normalize_special_terms(text)
        
        # Обрабатываем дефисы
        text = self._process_hyphens(text)
        
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
            
        # Заменяем "; ." на "."
        text = re.sub(r';\s*\.', '.', text)
        
        # В самом конце заменяем ". ." на "."
        text = re.sub(r'\.\s*\.', '.', text)
            
        return text 