import sqlite3
from datetime import datetime

class DatabaseManager:
    def __init__(self, db_path="convicts.db"):
        self.db_path = db_path
        self.create_database()
    
    def create_database(self):
        """Создание базы данных и необходимых таблиц"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Создаем таблицу осужденных
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS convicts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            start_date TEXT,           -- A: дата начала учета
            birth_date TEXT,           -- C: дата рождения
            end_date TEXT,             -- F: дата окончания учета
            court_info TEXT,           -- K: информация о суде и наказании
            restrictions TEXT,         -- L: ограничения
            full_name TEXT,            -- N: ФИО
            address TEXT,              -- O: адрес
            phone TEXT,                -- P: телефон
            other_info_g TEXT,         -- G: иная информация
            other_info_h TEXT,         -- H: иная информация
            other_info_q TEXT,         -- Q: иная информация
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')
        
        # Создаем таблицу характеристик
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS characteristics (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            convict_id INTEGER,
            type TEXT CHECK(type IN ('positive', 'neutral', 'negative')),
            is_template BOOLEAN DEFAULT 0,
            text TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (convict_id) REFERENCES convicts (id)
        )
        ''')
        
        conn.commit()
        conn.close()
    
    def add_convict(self, data):
        """Добавление нового осужденного в базу данных"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Объединяем дополнительную информацию
        other_info = []
        if data.get('other_info_g'):
            other_info.append(data['other_info_g'])
        if data.get('other_info_h'):
            other_info.append(data['other_info_h'])
        if data.get('other_info_q'):
            other_info.append(data['other_info_q'])
        
        other_info_text = ' | '.join(other_info) if other_info else None
        
        cursor.execute('''
        INSERT INTO convicts (
            start_date, birth_date, end_date, court_info, 
            restrictions, full_name, address, phone, other_info_g, other_info_h, other_info_q
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data.get('start_date'),
            data.get('birth_date'),
            data.get('end_date'),
            data.get('court_info'),
            data.get('restrictions'),
            data.get('full_name'),
            data.get('address'),
            data.get('phone'),
            data.get('other_info_g'),
            data.get('other_info_h'),
            other_info_text
        ))
        
        convict_id = cursor.lastrowid
        conn.commit()
        conn.close()
        return convict_id
    
    def add_characteristic(self, convict_id, characteristic_type, text, is_template=False):
        """Добавление характеристики"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
        INSERT INTO characteristics (convict_id, type, text, is_template)
        VALUES (?, ?, ?, ?)
        ''', (convict_id, characteristic_type, text, is_template))
        
        conn.commit()
        conn.close()
    
    def get_convicts(self):
        """Получение списка всех осужденных"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT * FROM convicts')
        convicts = cursor.fetchall()
        
        conn.close()
        return convicts
    
    def get_characteristics(self, convict_id=None, characteristic_type=None):
        """Получение характеристик"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        query = 'SELECT * FROM characteristics WHERE 1=1'
        params = []
        
        if convict_id:
            query += ' AND convict_id = ?'
            params.append(convict_id)
        
        if characteristic_type:
            query += ' AND type = ?'
            params.append(characteristic_type)
        
        cursor.execute(query, params)
        characteristics = cursor.fetchall()
        
        conn.close()
        return characteristics
    
    def delete_characteristic(self, characteristic_id):
        """Удаление характеристики"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('DELETE FROM characteristics WHERE id = ?', (characteristic_id,))
        
        conn.commit()
        conn.close() 