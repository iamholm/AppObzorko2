import pandas as pd
import re
import os
from tkinter import Tk, filedialog, Button, Label, messagebox, ttk

class ImprovedAddressProcessor:
    def __init__(self):
        # Словарь, сопоставляющий названия улиц с их типами
        # Содержит только улицы Санкт-Петербурга, которые нужно обрабатывать
        self.street_types = {
            "Академика Байкова": "ул.",
            "Академика Глушко": "аллея",
            "Академика Константинова": "ул.",
            "Академика Лебедева": "ул.",
            "Амурская": "ул.",
            "Антоновская": "ул.",
            "Арсенальная": "наб.",
            "Арсенальная": "ул.",
            "Бестужевская": "ул.",
            "Бобруйская": "ул.",
            "Богословская": "ул.",
            "Боткинская": "ул.",
            "Брюсовская": "ул.",
            "Брянцева": "ул.",
            "Бутлерова": "ул.",
            "Вавиловых": "ул.",
            "Васенко": "ул.",
            "Ватутина": "ул.",
            "Веденеева": "ул.",
            "Верности": "ул.",
            "Верхняя": "ул.",
            "Герасимовская": "ул.",
            "Гжатская": "ул.",
            "Гидротехников": "ул.",
            "Гражданский": "пр.",
            "Демьяна Бедного": "ул.",
            "Жукова": "ул.",
            "Замшина": "ул.",
            "Карпинского": "ул.",
            "Киришская": "ул.",
            "Ключевая": "ул.",
            "Комиссара Смирнова": "ул.",
            "Комсомола": "ул.",
            "Кондратьевский": "пр.",
            "Культуры": "пр.",
            "Кушелевская": "дор.",
            "Лабораторная": "ул.",
            "Лабораторный": "пр.",
            "Лесной": "пр.",
            "Литовская": "ул.",
            "Лужская": "ул.",
            "Луначарского": "пр.",
            "Маршала Блюхера": "пр.",
            "Менделеевская": "ул.",
            "Меншиковский": "пр.",
            "Металлистов": "пр.",
            "Мечникова": "пр.",
            "Минеральная": "ул.",
            "Михайлова": "ул.",
            "Нартовская": "ул.",
            "Науки": "пр.",
            "Нейшлотский": "пер.",
            "Непокорённых": "пр.",
            "Новороссийская": "ул.",
            "Обручевых": "ул.",
            "Ольги Форш": "ул.",
            "Печорская": "ул.",
            "Пискарёвский": "пр.",
            "Политехническая": "ул.",
            "Полюстровский": "пр.",
            "Просвещения": "пр.",
            "Руставели": "ул.",
            "Свердловская": "наб.",
            "Светлановский": "пр.",
            "Северный": "пр.",
            "Сибирская": "ул.",
            "Софьи Ковалевской": "ул.",
            "Старо-Муринская": "ул.",
            "Старцева": "ул.",
            "Суздальский": "пр.",
            "Тимуровская": "ул.",
            "Тихорецкий": "пр.",
            "Токсовская": "ул.",
            "Архитектора Баранова": "ул.",
            "Рериха": "ул.",
            "Усыскина": "пер.",
            "Учительская": "ул.",
            "Ушинского": "ул.",
            "Фаворского": "ул.",
            "Федосеенко": "ул.",
            "Феодосийская": "ул.",
            "Финский": "пер.",
            "Хлопина": "ул.",
            "Черкасова": "ул.",
            "Чичуринский": "пер.",
            "Чугунная": "ул."
        }
        
        # Создаем список названий улиц из ключей словаря
        self.street_names = list(self.street_types.keys())
        
        # Префиксы улиц для удаления
        self.street_prefixes = [
            "пр\. ", "пр\.", "пр ", "пр",
            "ул\. ", "ул\.", "ул ", "ул",
            "проспект ", "проспект",
            "улица ", "улица",
            "аллея ", "б-р ", "бульвар ", "дор\. ", "дорога ",
            "наб\. ", "набережная ", "пер\. ", "переулок ",
            "пл\. ", "площадь ", "проезд ", "просп\. ", "просп"
        ]
        
        # Город
        self.city_prefixes = [
            "г\.СПб", "г\. СПб", "СПб", "Санкт-Петербург", 
            "г\.Санкт-Петербург", "г\. Санкт-Петербург"
        ]
        
        # Паттерны для поиска информации о доме - от самого специфического к самому общему
        self.house_patterns = [
            # 1. Самые специфичные форматы с наибольшим количеством компонентов
            # "д. 14, корп. 1, лит. А, кв. 93" - дом, корпус, литера, квартира с запятыми
            r'\s*д\.?\s*(\d+),\s*корп\.?\s*(\d+),\s*лит\.?\s*([А-Яа-я]),\s*кв\.?\s*(\d+)',
            
            # "27-2-А-17" - дом-корпус-литера-квартира через дефисы
            r'\s*(\d+)-(\d+)-([А-Яа-я])-(\d+)',
            
            # "5-2А-2" - дом-корпуслитера-квартира через дефисы
            r'\s*(\d+)-(\d+)([А-Яа-я])-(\d+)',
            
            # "34 к.1, лит. А, кв. 64" - дом, корпус, литера, квартира с запятыми
            r'\s*(\d+)\s*к\.?\s*(\d+),\s*лит\.?\s*([А-Яа-я]),\s*кв\.?\s*(\d+)',
            
            # 2. Форматы с "пр." после названия проспекта
            # "пр., д.125, корп.3, кв.30"
            r'пр\.,\s*д\.?\s*(\d+),\s*корп\.?\s*(\d+),\s*кв\.?\s*(\d+)',
            
            # 3. Форматы с тремя компонентами через запятые
            # "д. 84, корп. 3, кв. 124"
            r'\s*д\.?\s*(\d+),\s*корп\.?\s*(\d+),\s*кв\.?\s*(\d+)',
            
            # "д.7, корп. 1, кв. 803" - без пробела после д.
            r'\s*д\.?(\d+),\s*корп\.?\s*(\d+),\s*кв\.?\s*(\d+)',
            
            # "д.6, к.1, кв.34" - с сокращением "к." вместо "корп."
            r'\s*д\.?\s*(\d+),\s*к\.?\s*(\d+),\s*кв\.?\s*(\d+)',
            
            # 4. Форматы с литерой и квартирой
            # "д. 6А кв. 31" - дом с литерой слитно и квартирой
            r'\s*д\.?\s*(\d+)([А-Яа-я])\s+кв\.?\s*(\d+)',
            
            # "д. 30/А кв. 85" - дом с литерой через слэш и квартирой
            r'\s*д\.?\s*(\d+)\/([А-Яа-я])\s+кв\.?\s*(\d+)',
            
            # "д. 24-А кв. 50" - дом-литера через дефис и квартира
            r'\s*д\.?\s*(\d+)-([А-Яа-я])\s+кв\.?\s*(\d+)',
            
            # 5. Форматы с дробным номером дома/корпуса
            # "д. 11/16, кв. 54"
            r'\s*д\.?\s*(\d+)\/(\d+),\s*кв\.?\s*(\d+)',
            
            # 6. Форматы с домом и квартирой через запятую
            # "д. 130, кв. 231"
            r'\s*д\.?\s*(\d+),\s*кв\.?\s*(\d+)',
            
            # "д.14, кв.5" - без пробелов после д. и кв.
            r'\s*д\.?(\d+),\s*кв\.?(\d+)',
            
            # 7. Остальные форматы с тремя компонентами
            # д.19 корп. 3 кв.12
            r'\s*д\.?\s+(\d+)\s+корп\.?\s+(\d+)\s+кв\.?\s+(\d+)',
            
            # д. 58 к. 1 кв. 28, д. 130 к. 1 кв. 390
            r'\s*д\.?\s+(\d+)\s+к\.?\s+(\d+)\s+кв\.?\s+(\d+)',
            
            # д.14 кор.1 кв.204
            r'\s*д\.?\s*(\d+)\s*кор\.?\s*(\d+)\s*кв\.?\s*(\d+)',
            
            # д. 48/3 кв. 87
            r'\s*д\.?\s+(\d+)\/(\d+)\s+кв\.?\s+(\d+)',
            
            # д.19 корп. 3 кв.12 - без пробелов
            r'\s*д\.?(\d+)корп\.?(\d+)кв\.?(\d+)',
            
            # д.14кор.1кв.204 - без пробелов
            r'\s*д\.?(\d+)кор\.?(\d+)кв\.?(\d+)',
            
            # д.58к.1кв.28 - без пробелов
            r'\s*д\.?(\d+)к\.?(\d+)кв\.?(\d+)',
            
            # 8. Форматы с тремя компонентами и специальными символами
            # "106—1-112" - длинное тире в формате дом-корпус-квартира
            r'\s*(\d+)—(\d+)-(\d+)',
            
            # дом-корпус-квартира: 114-4-35
            r'\s*(\d+)-(\d+)-(\d+)',
            
            # "14/3-82", "8/1-207" - дом/корпус-квартира
            r'\s*(\d+)\/(\d+)-(\d+)',
            
            # 9. Форматы с двумя компонентами
            # дом квартира: д. 14 кв. 5
            r'\s*д\.?\s+(\d+)\s+(?:кв\.?|квартира)\s*(\d+)',
            
            # дом квартира: 14 кв. 5 (без д.)
            r'\s*(\d+)\s+(?:кв\.?|квартира)\s*(\d+)',
            
            # дом/корпус квартира: 8/4 кв.34
            r'\s*(\d+)\/(\d+)\s+(?:кв\.?|квартира)\s*(\d+)',
            
            # ", 99-110" - дом-квартира с возможной запятой в начале
            r'(?:,\s*)?(\d+)-(\d+)(?!\d)(?:\s|$|,)',
            
            # 10. Дом с корпусом без квартиры
            # дом корпус: д. 19 корп. 3
            r'\s*д\.?\s+(\d+)\s+(?:корп\.?|кор\.?|к\.?)\s*(\d+)(?!\s*кв)',
            
            # дом корпус: 19 корп. 3 (без д.)
            r'\s*(\d+)\s+(?:корп\.?|кор\.?|к\.?)\s*(\d+)(?!\s*кв)',
            
            # 11. Самые общие форматы
            # просто номер дома с д.: д. 15 
            r'\s*д\.?\s+(\d+)(?!\d)(?!\s*-\d+)(?!\s*\/\d+)(?!\s*(?:корп|кор|к))(?!\s*кв)',
            
            # просто номер дома: 15
            r'\s*(\d+)(?!\d)(?!\s*-\d+)(?!\s*\/\d+)(?!\s*(?:корп|кор|к))(?!\s*кв)'
        ]

    def extract_phone(self, text):
        """
        Извлекает телефонный номер из текста
        """
        if not text or pd.isna(text):
            return None, []
            
        text = str(text)
        
        # Находим все позиции цифр в тексте
        digits_with_positions = [(i, c) for i, c in enumerate(text) if c.isdigit()]
        
        if len(digits_with_positions) < 10:
            return None, []  # Недостаточно цифр для номера телефона
        
        # Берем последние 10 цифр с их позициями
        last_10_digits_with_positions = digits_with_positions[-10:]
        
        # Проверяем, что между цифрами нет слишком больших разрывов
        max_gap = 3  # Максимально допустимый промежуток между последовательными цифрами
        
        for i in range(1, len(last_10_digits_with_positions)):
            prev_pos = last_10_digits_with_positions[i-1][0]
            curr_pos = last_10_digits_with_positions[i][0]
            
            if curr_pos - prev_pos > max_gap:
                # Слишком большой разрыв, вероятно, это не телефон
                return None, []
        
        # Извлекаем только цифры из последних 10 позиций
        last_10_digits = ''.join(digit for _, digit in last_10_digits_with_positions)
        
        # Начальная и конечная позиции 10 цифр в тексте
        start_pos = last_10_digits_with_positions[0][0]
        end_pos = last_10_digits_with_positions[-1][0] + 1
        
        # Проверяем, есть ли "8" перед номером
        raw_phone = last_10_digits
        if len(digits_with_positions) > 10:
            # Позиция и значение 11-й цифры с конца
            eleventh_pos, eleventh_digit = digits_with_positions[-11]
            
            # Проверяем, что это "8" и она расположена непосредственно перед началом основных 10 цифр
            # или отделена только допустимыми разделителями
            if eleventh_digit == '8' and (start_pos - eleventh_pos) <= max_gap:
                raw_phone = eleventh_digit + last_10_digits
                start_pos = eleventh_pos  # Включаем "8" в диапазон
        
        # Извлекаем оригинальный текст телефона
        original_phone_text = text[start_pos:end_pos]
        
        return raw_phone, [original_phone_text]

    def extract_address(self, text):
        """
        Извлекает адрес из текста согласно новым жестким правилам.
        Обрабатывает только улицы из словаря и дома/квартиры, указанные после этих улиц.
        """
        if not text or pd.isna(text):
            return None, None
            
        text = str(text).strip()
        
        # Предобработка: заменяем длинное тире на обычное
        text = text.replace('—', '-')
        
        # Создаем регулярное выражение для поиска улиц
        # Сортируем названия улиц по длине, чтобы сначала искать более длинные
        sorted_streets = sorted(self.street_names, key=len, reverse=True)
        street_pattern = '|'.join(map(re.escape, sorted_streets))
        
        # Поиск названия улицы
        street_match = re.search(fr'({street_pattern})', text, re.IGNORECASE)
        if not street_match:
            return None, None
            
        street_name = street_match.group(1)
        street_pos = street_match.start()
        
        # Проверка пробелов до и после названия улицы
        # и удаление префиксов (ул., пр., СПб и т.д.)
        
        # Проверяем, есть ли префиксы улиц перед названием
        prefix_text = text[:street_pos]
        prefix_pattern = '|'.join(self.street_prefixes)
        prefix_match = re.search(fr'({prefix_pattern})\s*$', prefix_text, re.IGNORECASE)
        
        # Ищем указание города перед префиксом улицы
        city_pattern = '|'.join(map(re.escape, self.city_prefixes))
        city_text = text[:street_pos]
        city_match = re.search(fr'({city_pattern})[,\s]*$', city_text, re.IGNORECASE)
        
        # Позиция начала адреса (для определения оригинального адреса)
        start_pos = street_pos
        if prefix_match:
            start_pos = prefix_match.start()
            if city_match and city_match.end() == prefix_match.start():
                start_pos = city_match.start()
        elif city_match:
            start_pos = city_match.start()
        
        # Ищем информацию о доме после названия улицы
        house_text = text[street_match.end():]
        
        house_info = ""
        original_house_text = ""
        
        # Отладочная информация при необходимости
        # print(f"Текст для поиска информации о доме: '{house_text}'")
        
        for i, pattern in enumerate(self.house_patterns):
            house_match = re.search(pattern, house_text, re.IGNORECASE)
            
            # Отладочная информация при необходимости
            # print(f"Паттерн #{i}: '{pattern}'")
            # print(f"Результат: {house_match}")
            
            if house_match:
                original_house_text = house_match.group(0)
                house_digits = house_match.groups()
                
                # Отладочная информация при необходимости
                # print(f"Найдено соответствие: '{original_house_text}'")
                # print(f"Группы: {house_digits}")
                
                # Форматируем информацию о доме в зависимости от шаблона
                if len(house_digits) == 4:
                    # Проверяем, является ли третья группа литерой
                    if re.match(r'[А-Яа-я]', house_digits[2]):
                        # Форматы с домом, корпусом, литерой, квартирой
                        # "д. 14, корп. 1, лит. А, кв. 93" -> "14-1А-93"
                        # "27-2-А-17" -> "27-2А-17"
                        house_info = f"{house_digits[0]}-{house_digits[1]}{house_digits[2]}-{house_digits[3]}"
                    else:
                        # Общий случай для четырех компонентов
                        house_info = f"{house_digits[0]}-{house_digits[1]}-{house_digits[2]}-{house_digits[3]}"
                elif len(house_digits) == 3:
                    # Проверяем, является ли вторая группа литерой
                    if re.match(r'[А-Яа-я]', house_digits[1]):
                        # Форматы с домом, литерой, квартирой
                        # "д. 6А кв. 31" -> "6А-31"
                        # "д. 30/А кв. 85" -> "30А-85"
                        # "д. 24-А кв. 50" -> "24А-50"
                        house_info = f"{house_digits[0]}{house_digits[1]}-{house_digits[2]}"
                    else:
                        # Общий случай для трех компонентов: дом-корпус-квартира
                        house_info = f"{house_digits[0]}-{house_digits[1]}-{house_digits[2]}"
                elif len(house_digits) == 2:
                    # дом-квартира или дом-корпус
                    house_info = f"{house_digits[0]}-{house_digits[1]}"
                elif len(house_digits) == 1:
                    # просто номер дома
                    house_info = house_digits[0]
                
                break
        
        # Определяем конечную позицию адреса
        end_pos = street_match.end()
        if original_house_text:
            house_start = text.find(original_house_text, street_match.end())
            if house_start != -1:
                end_pos = house_start + len(original_house_text)
        
        # Извлекаем полный исходный текст адреса
        original_address = text[start_pos:end_pos].strip()
        
        # Формируем стандартизированный адрес
        # Добавляем правильный тип улицы из словаря
        street_type = self.street_types.get(street_name, "ул.")
        formatted_address = f"{street_type} {street_name}"
        
        if house_info:
            formatted_address += f" {house_info}"
        
        return formatted_address.strip(), original_address

    def clean_other_info(self, text):
        """
        Очищает текст для поля "Иная информация"
        Удаляет специфичные фразы и лишние знаки пунктуации
        """
        if not text or pd.isna(text):
            return None
            
        text = str(text).strip()
        
        # Базовая очистка - заменяем множественные пробелы одиночными
        text = re.sub(r'\s+', ' ', text)
        
        # Удаляем специфичные фразы (от более специфичных к более общим)
        text = re.sub(r'Зарг\s+и\s+прож\.?', '', text, flags=re.IGNORECASE)
        text = re.sub(r'Зарег\s+и\s+прож\.?', '', text, flags=re.IGNORECASE)
        text = re.sub(r'\bГражданство\b', '', text, flags=re.IGNORECASE)
        text = re.sub(r'Р\s*Ф\b', '', text)
        text = re.sub(r'\bтел\.?\b', '', text, flags=re.IGNORECASE)
        text = re.sub(r'г\.\s*СПб\b', '', text, flags=re.IGNORECASE)
        text = re.sub(r'г\.\s+', '', text)
        
        # Очищаем от множественных знаков пунктуации
        text = re.sub(r'[,-]+', '', text)
        text = re.sub(r'^\s*[,-]\s*', '', text)
        text = re.sub(r'\s*[,-]\s*$', '', text)
        
        # Финальная очистка пробелов
        text = text.strip()
        
        if not text:
            return None
            
        return text

class Application:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка адресов")
        self.root.geometry("600x250")
        
        self.processor = ImprovedAddressProcessor()
        self.progress = ttk.Progressbar(root, orient="horizontal", length=500, mode="determinate")
        
        Label(root, text="Выберите файл Excel для обработки", font=("Arial", 12)).pack(pady=10)
        Button(root, text="Выбрать файл", command=self.process_file, font=("Arial", 12)).pack(pady=5)
        self.progress.pack(pady=10)
        
        self.status_label = Label(root, text="Ожидание выбора файла...", font=("Arial", 10))
        self.status_label.pack(pady=5)
        
        # Добавляем тестовую функцию для отладки
        Button(root, text="Тестировать обработку", command=self.test_processor, font=("Arial", 12)).pack(pady=5)

    def test_processor(self):
        # Окно для ввода тестового текста
        from tkinter import simpledialog, scrolledtext, Toplevel
        
        test_window = Toplevel(self.root)
        test_window.title("Тестирование обработки адреса")
        test_window.geometry("600x400")
        
        Label(test_window, text="Введите текст для обработки:", font=("Arial", 12)).pack(pady=5)
        
        text_input = scrolledtext.ScrolledText(test_window, width=60, height=5)
        text_input.pack(pady=5, padx=10, fill='both', expand=True)
        
        result_text = scrolledtext.ScrolledText(test_window, width=60, height=10)
        result_text.pack(pady=5, padx=10, fill='both', expand=True)
        
        def process_test_text():
            input_text = text_input.get("1.0", "end-1c")
            
            # 1. Сначала извлекаем телефон
            phones, original_phones = self.processor.extract_phone(input_text)
            
            # 2. Удаляем телефон из исходного текста
            text_without_phone = input_text
            if original_phones:
                for phone in original_phones:
                    text_without_phone = text_without_phone.replace(phone, '')
            
            # 3. Извлекаем адрес из текста без телефона
            formatted_address, original_address = self.processor.extract_address(text_without_phone)
            
            # 4. Определяем иное - всё, что осталось после удаления телефона и адреса
            other_info = text_without_phone
            if original_address:
                other_info = other_info.replace(original_address, '')
            
            # Очищаем результат с новыми правилами
            other_info = self.processor.clean_other_info(other_info)
            
            result = f"Исходный текст:\n{input_text}\n\n"
            result += f"Текст без телефона:\n{text_without_phone}\n\n"
            result += f"Телефон: {phones or 'Не найден'}\n"
            result += f"Оригинальный текст телефона: {original_phones[0] if original_phones else 'Не найден'}\n"
            result += f"Форматированный адрес: {formatted_address or 'Не найден'}\n"
            result += f"Исходный адрес: {original_address or 'Не найден'}\n"
            result += f"Прочая информация: {other_info or 'Не найдена'}\n"
            
            result_text.delete("1.0", "end")
            result_text.insert("1.0", result)
        
        Button(test_window, text="Обработать", command=process_test_text, font=("Arial", 12)).pack(pady=5)

    def process_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите файл Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
            
        try:
            self.progress["value"] = 0
            self.status_label.config(text="Начинаем обработку файла...")
            self.root.update_idletasks()
            
            # Чтение файла
            df = pd.read_excel(file_path, header=None, names=['raw_data'])
            
            # Обработка данных с прогрессом
            total_rows = len(df)
            results = []
            
            for i, row in df.iterrows():
                raw_text = row['raw_data']
                
                try:
                    # 1. Сначала извлекаем телефон
                    phones, original_phones = self.processor.extract_phone(raw_text)
                    
                    # 2. Удаляем телефон из исходного текста
                    text_without_phone = raw_text
                    if original_phones:
                        for phone in original_phones:
                            text_without_phone = text_without_phone.replace(phone, '')
                    
                    # 3. Извлекаем адрес из текста без телефона
                    formatted_address, original_address = self.processor.extract_address(text_without_phone)
                    
                    # 4. Определяем иное - всё, что осталось после удаления телефона и адреса
                    other_info = text_without_phone
                    if original_address:
                        other_info = other_info.replace(original_address, '')
                    
                    # Очищаем результат с использованием нового метода
                    other_info = self.processor.clean_other_info(other_info)
                    
                    results.append({
                        'raw_data': raw_text,
                        'address': formatted_address,
                        'phone': phones,
                        'other_info': other_info
                    })
                except Exception as e:
                    # В случае ошибки в обработке строки, добавляем ее без обработки
                    results.append({
                        'raw_data': raw_text,
                        'address': f"ОШИБКА: {str(e)}",
                        'phone': None,
                        'other_info': None
                    })
                
                # Обновление прогресса
                self.progress["value"] = (i + 1) / total_rows * 100
                self.status_label.config(text=f"Обработано {i+1} из {total_rows} строк...")
                if i % 10 == 0:  # Обновляем интерфейс каждые 10 строк
                    self.root.update_idletasks()
            
            # Создание DataFrame с результатами
            result_df = pd.DataFrame(results)
            
            # Сохранение результата
            output_dir = os.path.dirname(file_path)
            output_filename = f"processed_{os.path.basename(file_path)}"
            output_path = os.path.join(output_dir, output_filename)
            
            result_df.to_excel(output_path, index=False)
            
            messagebox.showinfo(
                "Готово!",
                f"Файл успешно обработан и сохранён как:\n{output_path}"
            )
            self.progress["value"] = 0
            self.status_label.config(text="Обработка завершена.")
            
        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Произошла ошибка при обработке файла:\n{str(e)}"
            )
            self.progress["value"] = 0
            self.status_label.config(text=f"Ошибка: {str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = Application(root)
    root.mainloop()