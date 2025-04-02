import re

class ColumnBFormatter:
    """
    Класс для форматирования текста в столбце B (имена)
    """
    
    def process_excel_column(self, sheet, column_index=2):
        """
        Обрабатывает все ячейки в столбце B
        
        Args:
            sheet: Лист Excel
            column_index (int): Индекс столбца (по умолчанию 2 для столбца B)
            
        Returns:
            dict: Статистика обработки
        """
        stats = {
            'cells_processed': 0,
            'names_split': 0,
            'names_moved': 0,
            'names_formatted': 0  # Для совместимости
        }
        
        # Определяем, с какой строки начать (обязательно с 1, чтобы обработать первую строку)
        start_row = 1
        
        for row in range(start_row, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=column_index)
            if cell.value:
                original_text = str(cell.value)
                
                # Проверяем наличие скобок в тексте
                if '(' in original_text and ')' in original_text:
                    # Обрабатываем имена со скобками отдельно
                    if self._process_name_with_parentheses(sheet, row, original_text):
                        stats['names_moved'] += 1
                        continue
                
                # Разделяем слитные имена (например, "ЖумановИсабекМаратбекович")
                formatted_text = self._split_joined_names(original_text)
                
                if formatted_text != original_text:
                    cell.value = formatted_text
                    stats['cells_processed'] += 1
                    stats['names_split'] += 1
                    stats['names_formatted'] += 1
                
                # Перенос имени в столбец N
                if self._move_name_to_column_n(sheet, row, cell.value):
                    stats['names_moved'] += 1
                
                # Нормализация форматирования СПб
                if cell.value:
                    cell.value = self._normalize_spb_formatting(str(cell.value))
                    
                # Удаление "г. СПб" из текста (включая первую строку)
                if cell.value:
                    cell.value = self._remove_spb_from_text(str(cell.value))
        
        return stats
    
    def _split_joined_names(self, text):
        """
        Разделяет слитные имена с заглавными буквами внутри слова до встречи с СПб
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Текст с разделенными именами
        """
        # Удаляем лишние пробелы
        text = re.sub(r'\s+', ' ', text).strip()
        
        # Пропускаем пустой текст
        if not text:
            return text
            
        # Разбиваем текст на слова
        words = text.split()
        result_words = []
        
        # Стоп-слова, после которых прекращаем обработку
        stop_words = ['СПб', 'г.', 'г.СПб', 'г. СПб', 'пр.', 'ул.', 'д.', 'р-н', 'обл.', 'респ.']
        
        # Слова, которые не нужно разделять
        preserve_words = ['СПб']
        
        # Обрабатываем слова до встречи со стоп-словом
        for i, word in enumerate(words):
            # Если встретили стоп-слово, добавляем его и все последующие слова без изменений
            if word in stop_words:
                result_words.append(word)
                # Добавляем все оставшиеся слова без обработки
                result_words.extend(words[i+1:])
                break
                
            # Сохраняем слова из списка preserve_words без изменений
            if word in preserve_words:
                result_words.append(word)
                continue
                
            # Проверяем, есть ли внутри слова заглавные буквы (не в начале)
            if len(word) > 1 and re.search(r'[А-ЯЁA-Z]', word[1:]):
                # Если слово содержит "СПб", добавляем его как есть
                if "СПб" in word:
                    result_words.append(word)
                    continue
                    
                # Находим все позиции заглавных букв
                capitals = [0] + [i for i in range(1, len(word)) if word[i].isupper()]
                
                # Если нашли заглавные буквы внутри слова
                if len(capitals) > 1:
                    # Разделяем слово по позициям заглавных букв
                    split_words = []
                    for j in range(len(capitals)):
                        start = capitals[j]
                        end = capitals[j+1] if j+1 < len(capitals) else len(word)
                        split_words.append(word[start:end])
                    
                    # Добавляем разделенные слова
                    result_words.extend(split_words)
                    continue
            
            # Если слово не нужно разделять, добавляем как есть
            result_words.append(word)
            
        return ' '.join(result_words)
    
    def _check_for_patronymic(self, sheet, row):
        """
        Проверяет, остались ли в столбце B отчества, и переносит их в столбец N
        
        Args:
            sheet: Лист Excel
            row (int): Номер строки
        """
        cell_b = sheet.cell(row=row, column=2)
        cell_n = sheet.cell(row=row, column=14)
        
        if not cell_b.value:
            return
            
        text = str(cell_b.value).strip()
        words = text.split()
        
        if not words:
            return
            
        # Проверяем первое слово на признаки отчества
        first_word = words[0]
        
        # Типичные окончания отчеств
        patronymic_endings = ['вич', 'вна', 'ич', 'ична', 'ична', 'овна', 'евна', 'ович', 'евич']
        
        is_patronymic = False
        
        # Проверяем, содержит ли первое слово типичные окончания отчеств
        for ending in patronymic_endings:
            if first_word.lower().endswith(ending) and len(first_word) > len(ending) + 1:
                is_patronymic = True
                break
                
        # Если первое слово - отчество
        if is_patronymic:
            # Проверяем, не является ли следующее слово "СПб" или "г."
            stop_words = ['СПб', 'г.', 'г.СПб', 'г. СПб', 'пр.', 'ул.', 'д.']
            
            if len(words) > 1 and any(words[1] == stop_word for stop_word in stop_words):
                # Переносим только отчество в столбец N
                patronymic = first_word
                
                # Добавляем отчество в столбец N
                if cell_n.value:
                    cell_n.value = str(cell_n.value) + ' ' + patronymic
                else:
                    cell_n.value = patronymic
                
                # Удаляем отчество из столбца B
                remaining_text = ' '.join(words[1:]).strip()
                cell_b.value = remaining_text if remaining_text else None
    
    def _move_name_to_column_n(self, sheet, row, name_text):
        """
        Переносит имя в столбец N и удаляет только имя из B, оставляя остальной текст
        
        Args:
            sheet: Лист Excel
            row (int): Номер строки
            name_text (str): Текст имени
            
        Returns:
            bool: True если имя было перенесено, False в противном случае
        """
        # Получаем ячейку в столбце N
        cell_n = sheet.cell(row=row, column=14)
        cell_b = sheet.cell(row=row, column=2)
        
        # Если текст пустой, выходим
        if not name_text:
            return False
            
        name_text = str(name_text).strip()
        
        # Удаляем запятые и другие знаки препинания из текста для корректного разделения на слова
        # Сохраняем оригинальный текст для последующего удаления из столбца B
        original_text = name_text
        processed_text = re.sub(r'([а-яА-ЯёЁa-zA-Z])[,;:.!?]+', r'\1', name_text)
        
        # Разбиваем текст на слова
        words = processed_text.split()
        
        # Если текст пустой после разбиения, выходим
        if not words:
            return False
        
        # Стоп-слова, которые указывают на конец имени
        stop_words = ['СПб', 'г.', 'г.СПб', 'г. СПб', 'пр.', 'ул.', 'д.', 'р-н', 'обл.', 'респ.', 'Гражданский', 'кв.', 'лит.']
        
        # Проверяем, начинается ли текст со стоп-слова
        if words[0] in stop_words:
            return False
        
        # Специальные окончания для иностранных имен
        foreign_name_endings = ['Оглы', 'Кизи', 'Кызы', 'оглы', 'кизи', 'кызы', 'угли', 'Угли']
        
        # Определяем максимальное количество слов для проверки
        # По умолчанию - 3 слова (фамилия, имя, отчество)
        max_words = min(len(words), 3)
        
        # Проверка на наличие специального окончания
        has_foreign_ending = False
        
        # Проверяем все слова на наличие иностранных окончаний
        for i, word in enumerate(words):
            if i >= len(words):
                break
                
            # Проверяем, является ли слово иностранным окончанием или содержит его
            for ending in foreign_name_endings:
                if word == ending or word.endswith(ending):
                    has_foreign_ending = True
                    # Если окончание находится в 4-м слове, увеличиваем max_words до 4
                    if i == 3:
                        max_words = 4
                    # Если окончание находится в 5-м слове, увеличиваем max_words до 5
                    elif i == 4:
                        max_words = 5
                    break
                    
            if has_foreign_ending:
                break
        
        # Берем слова до первого стоп-слова, но не более max_words
        name_parts = []
        for i in range(min(max_words, len(words))):
            # Если встретили стоп-слово, останавливаемся
            if words[i] in stop_words:
                break
            name_parts.append(words[i])
        
        # Если не нашли имени, выходим
        if not name_parts:
            return False
        
        # Формируем имя
        formatted_name = ' '.join(name_parts)
        
        # Очищаем имя от знаков препинания в конце
        formatted_name = re.sub(r'[.,;:!?]+$', '', formatted_name).strip()
        
        # Если в ячейке N уже есть текст, добавляем имя в начало (без точки)
        if cell_n.value:
            cell_n.value = formatted_name + ' ' + str(cell_n.value)
        else:
            cell_n.value = formatted_name
        
        # Находим начало оставшегося текста в оригинальной строке
        remaining_text = original_text
        
        # Удаляем имя из столбца B и оставляем остальной текст
        # Ищем первое стоп-слово после имени
        for i, word in enumerate(words):
            if i >= len(name_parts):
                if word in stop_words:
                    # Находим это стоп-слово в оригинальном тексте
                    pattern = r'(?:^|\s)' + re.escape(word)
                    match = re.search(pattern, remaining_text)
                    if match:
                        # Оставляем текст начиная с этого стоп-слова
                        remaining_text = remaining_text[match.start():].strip()
                        break
        
        # Если не нашли стоп-слов, удаляем имя другим способом
        if remaining_text == original_text:
            # Определяем, сколько слов мы взяли для имени
            name_word_count = len(name_parts)
            
            # Собираем оставшиеся слова
            remaining_words = words[name_word_count:]
            
            # Преобразуем обратно в текст
            if remaining_words:
                # Ищем первое оставшееся слово в оригинальном тексте
                pattern = r'(?:^|\s)' + re.escape(remaining_words[0])
                match = re.search(pattern, remaining_text)
                if match:
                    # Оставляем текст начиная с первого оставшегося слова
                    remaining_text = remaining_text[match.start():].strip()
                else:
                    # Если не нашли, просто соединяем оставшиеся слова
                    remaining_text = ' '.join(remaining_words).strip()
            else:
                remaining_text = ""
        
        # Обновляем ячейку B
        cell_b.value = remaining_text if remaining_text else None
        
        return True
    
    def _normalize_spb_formatting(self, text):
        """
        Нормализует различные варианты написания СПб
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Текст с нормализованным форматированием СПб
        """
        if not text:
            return text
            
        # Удаляем лишние пробелы
        text = re.sub(r'\s+', ' ', text).strip()
        
        # 1. "С П Б" -> "г. СПб"
        text = re.sub(r'С\s+П\s+Б', 'г. СПб', text)
        
        # 2. "г. Санкт-Петербург" -> "г. СПб"
        text = re.sub(r'г\.\s*Санкт-Петербург', 'г. СПб', text)
        text = re.sub(r'Санкт-Петербург', 'г. СПб', text)
        
        # 3. "СПб" (без префикса "г.") -> "г. СПб"
        # Проверяем, что перед "СПб" нет "г."
        if 'СПб' in text and 'г. СПб' not in text:
            # Заменяем только если "СПб" не является частью другого слова
            text = re.sub(r'(?<!\w)СПб(?!\w)', 'г. СПб', text)
        
        return text 

    def _process_name_with_parentheses(self, sheet, row, text):
        """
        Специальная обработка для имен со скобками.
        Например: "Славин (Сенин) Александр Викторович"
        
        Args:
            sheet: Лист Excel
            row (int): Номер строки
            text (str): Исходный текст
            
        Returns:
            bool: True если имя было обработано и перенесено, False в противном случае
        """
        # Получаем ячейки в столбцах N и B
        cell_n = sheet.cell(row=row, column=14)
        cell_b = sheet.cell(row=row, column=2)
        
        # Сохраняем оригинальный текст
        original_text = text
        
        # Удаляем лишние пробелы
        text = re.sub(r'\s+', ' ', text).strip()
        
        # Добавляем пробел после закрывающей скобки, если за ней сразу идет буква
        # Например: "(Науменко)Наталья" -> "(Науменко) Наталья"
        text = re.sub(r'\)([а-яА-ЯёЁa-zA-Z])', r') \1', text)
        
        # Убираем запятые и другие знаки препинания для корректного разделения на слова
        processed_text = re.sub(r'([а-яА-ЯёЁa-zA-Z])[,;:.!?]+', r'\1', text)
        
        # Стоп-слова, которые указывают на конец имени
        stop_words = ['СПб', 'г.', 'г.СПб', 'г. СПб', 'пр.', 'ул.', 'д.', 'р-н', 'обл.', 'респ.', 'Гражданский', 'кв.', 'лит.']
        
        # Разбиваем текст на слова
        words = processed_text.split()
        
        # Проверяем, начинается ли текст со стоп-слова
        if not words or words[0] in stop_words:
            return False
        
        # Специальные окончания для иностранных имен
        foreign_name_endings = ['Оглы', 'Кизи', 'Кызы', 'оглы', 'кизи', 'кызы', 'угли', 'Угли']
        
        # Проверяем наличие иностранных окончаний
        has_foreign_ending = False
        for i, word in enumerate(words):
            if i >= len(words) or i > 4:  # Проверяем только первые 5 слов
                break
                
            for ending in foreign_name_endings:
                if word == ending or word.endswith(ending):
                    has_foreign_ending = True
                    if i == 4:  # Если окончание в 5-м слове
                        words = words[:5] + words[5:]
                    break
            
            if has_foreign_ending:
                break
        
        # Находим, где заканчивается ФИО и начинается адрес
        name_end_idx = -1
        for i, word in enumerate(words):
            if word in stop_words:
                name_end_idx = i
                break
        
        # Если не нашли стоп-слов, берем максимум 4 слова для имени
        # Или 5 слов, если 4-е или 5-е слово - иностранное окончание
        if name_end_idx == -1:
            if has_foreign_ending and len(words) >= 5:
                name_end_idx = 5
            else:
                name_end_idx = min(4, len(words))
        
        # Берем слова для имени
        name_parts = words[:name_end_idx]
        remaining_words = words[name_end_idx:]
        
        # Проверяем наличие скобок в имени
        has_parentheses = any('(' in word and ')' in word for word in name_parts) or \
                          any('(' in word for word in name_parts) and any(')' in word for word in name_parts)
        
        if not has_parentheses:
            return False
        
        # Формируем имя
        formatted_name = ' '.join(name_parts)
        
        # Очищаем имя от знаков препинания в конце
        formatted_name = re.sub(r'[.,;:!?]+$', '', formatted_name).strip()
        
        # Если в ячейке N уже есть текст, добавляем имя в начало
        if cell_n.value:
            cell_n.value = formatted_name + ' ' + str(cell_n.value)
        else:
            cell_n.value = formatted_name
        
        # Находим начало оставшегося текста в оригинальной строке
        remaining_text = original_text
        
        # Если есть оставшиеся слова, находим первое в оригинальном тексте
        if remaining_words:
            # Ищем первое оставшееся слово в оригинальном тексте
            pattern = r'(?:^|\s)' + re.escape(remaining_words[0])
            match = re.search(pattern, remaining_text)
            if match:
                # Оставляем текст начиная с первого оставшегося слова
                remaining_text = remaining_text[match.start():].strip()
            else:
                # Если не нашли, просто соединяем оставшиеся слова
                remaining_text = ' '.join(remaining_words).strip()
        else:
            remaining_text = ""
        
        # Обновляем ячейку B
        cell_b.value = remaining_text if remaining_text else None
        
        return True 

    def _remove_spb_from_text(self, text):
        """
        Удаляет "г. СПб" и "г. Спб" из текста
        
        Args:
            text (str): Исходный текст
            
        Returns:
            str: Текст без упоминаний СПб
        """
        if not text:
            return text
            
        # Удаление различных вариантов "г. СПб"
        # Учитываем разные варианты регистра и пробелов
        patterns = [
            r'г\.\s*СПб',   # г. СПб
            r'г\.\s*Спб',   # г. Спб
            r'г\.\s*спб',   # г. спб
            r'г\s*СПб',     # г СПб
            r'г\s*Спб',     # г Спб
            r'г\s*спб',     # г спб
            r'СПб',         # СПб
            r'Спб',         # Спб
            r'спб'          # спб
        ]
        
        # Запоминаем, была ли обнаружена запятая после СПб
        has_comma_after_spb = re.search(r'(?:г\.\s*СПб|г\.\s*Спб|СПб|Спб),', text)
        
        for pattern in patterns:
            text = re.sub(pattern, '', text)
        
        # Удаляем лишние пробелы
        text = re.sub(r'\s+', ' ', text).strip()
        
        # Удаляем запятые в начале текста, которые могли остаться после удаления "г. СПб"
        text = re.sub(r'^,\s*', '', text).strip()
        
        # Удаляем запятые, если текст пустой или состоит только из запятых
        if text and re.match(r'^[,\s]+$', text):
            text = ''
        
        return text 