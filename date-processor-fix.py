import re
from datetime import datetime

def _extract_all_dates_from_text(text):
    """
    Извлекает все даты из текста в виде списка
    
    Например, из "14.07.25 14.08.25" извлечет ["14.07.25", "14.08.25"]
    Также обрабатывает даты с пробелами: "13. 05. 2024" -> ["13. 05. 2024"]
    """
    # Шаблоны для поиска дат в тексте
    date_patterns = [
        r'\d{1,2}\s*\.\s*\d{1,2}\s*\.\s*\d{2,4}',  # ДД.ММ.ГГ или ДД.ММ.ГГГГ (с возможными пробелами)
        r'\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{2,4}',    # ДД/ММ/ГГ или ДД/ММ/ГГГГ (с возможными пробелами)
        r'\d{1,2}\s*-\s*\d{1,2}\s*-\s*\d{2,4}',    # ДД-ММ-ГГ или ДД-ММ-ГГГГ (с возможными пробелами)
        r'\d{2}\d{2}\s*\.\s*\d{2}',                # ДДММ.ГГ (с возможными пробелами)
        r'\d{8}'                                   # ДДММГГГГ
    ]
    
    # Список для найденных дат
    found_dates = []
    
    # Ищем все даты в тексте по шаблонам
    for pattern in date_patterns:
        matches = re.findall(pattern, text)
        found_dates.extend(matches)
    
    # Удаляем возможные дубликаты
    return list(dict.fromkeys(found_dates))

def _parse_and_normalize_date(date_str):
    """
    Парсит различные форматы дат и нормализует их к формату ДД.ММ.ГГГГ
    
    Поддерживает:
    - ДД.ММ.ГГ -> ДД.ММ.ГГГГ
    - ДД.ММ.ГГГГ (оставляем как есть)
    - ДД. ММ. ГГ -> ДД.ММ.ГГГГ (с пробелами)
    - ДД. ММ. ГГГГ -> ДД.ММ.ГГГГ (с пробелами)
    - ДДММ.ГГ -> ДД.ММ.ГГГГ (пропущена точка)
    - ДД/ММ/ГГ -> ДД.ММ.ГГГГ
    - ДД-ММ-ГГ -> ДД.ММ.ГГГГ
    - ДДММГГГГ -> ДД.ММ.ГГГГ (без разделителей)
    """
    # Удаляем пробелы перед обработкой, чтобы упростить регулярные выражения
    clean_date_str = re.sub(r'\s+', '', date_str)
    
    # Проверяем различные форматы даты
    
    # Формат ДД.ММ.ГГ (двузначный год)
    match = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{2})$', clean_date_str)
    if match:
        day, month, year = match.groups()
        full_year = _expand_year(year)
        return f"{int(day):02d}.{int(month):02d}.{full_year}"
    
    # Формат ДД.ММ.ГГГГ (четырехзначный год)
    match = re.match(r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', clean_date_str)
    if match:
        day, month, year = match.groups()
        return f"{int(day):02d}.{int(month):02d}.{year}"
    
    # Формат ДДММ.ГГ (пропущена точка между днем и месяцем)
    match = re.match(r'^(\d{2})(\d{2})\.(\d{2})$', clean_date_str)
    if match:
        day, month, year = match.groups()
        full_year = _expand_year(year)
        return f"{int(day):02d}.{int(month):02d}.{full_year}"
    
    # Формат ДДММГГГГ (без разделителей)
    match = re.match(r'^(\d{2})(\d{2})(\d{4})$', clean_date_str)
    if match:
        day, month, year = match.groups()
        return f"{int(day):02d}.{int(month):02d}.{year}"
    
    # Формат ДД/ММ/ГГ
    match = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{2})$', clean_date_str)
    if match:
        day, month, year = match.groups()
        full_year = _expand_year(year)
        return f"{int(day):02d}.{int(month):02d}.{full_year}"
    
    # Формат ДД/ММ/ГГГГ
    match = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', clean_date_str)
    if match:
        day, month, year = match.groups()
        return f"{int(day):02d}.{int(month):02d}.{year}"
    
    # Формат ДД-ММ-ГГ
    match = re.match(r'^(\d{1,2})-(\d{1,2})-(\d{2})$', clean_date_str)
    if match:
        day, month, year = match.groups()
        full_year = _expand_year(year)
        return f"{int(day):02d}.{int(month):02d}.{full_year}"
    
    # Формат ДД-ММ-ГГГГ
    match = re.match(r'^(\d{1,2})-(\d{1,2})-(\d{4})$', clean_date_str)
    if match:
        day, month, year = match.groups()
        return f"{int(day):02d}.{int(month):02d}.{year}"
    
    # Если ни один из форматов не подошел, возвращаем None
    return None

def _expand_year(year_str):
    """
    Преобразует двузначный год в четырехзначный
    Правило: 00-25 -> 2000-2025, 26-99 -> 1926-1999
    """
    year = int(year_str)
    
    # Определяем текущий год для расчета порога преобразования
    current_year = datetime.now().year
    current_short_year = current_year % 100
    
    if year <= current_short_year:
        return 2000 + year
    else:
        return 1900 + year

def _normalize_dates_in_court_info(sheet, column_index=9):
    """
    Нормализует все даты в столбце с информацией о судах к формату ДД.ММ.ГГГГ
    
    Возвращает количество нормализованных дат
    """
    normalized_count = 0
    
    # Обрабатываем все ячейки в указанном столбце
    for row in range(1, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=column_index)
        value = cell.value
        
        # Пропускаем пустые ячейки
        if not value:
            continue
            
        value_str = str(value).strip()
        
        # Извлекаем все даты из текста
        dates = _extract_all_dates_from_text(value_str)
        
        # Если даты найдены
        if dates:
            modified_value = value_str
            # Для каждой найденной даты
            for date in dates:
                # Нормализуем дату
                normalized_date = _parse_and_normalize_date(date)
                if normalized_date:
                    # Заменяем исходную дату на нормализованную
                    modified_value = modified_value.replace(date, normalized_date)
                    normalized_count += 1
            
            # Обновляем значение ячейки только если были изменения
            if modified_value != value_str:
                cell.value = modified_value
    
    return normalized_count
