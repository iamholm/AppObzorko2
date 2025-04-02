import tkinter as tk
from tkinter import ttk, messagebox
from database_manager import DatabaseManager

class DatabaseViewer:
    def __init__(self, parent):
        self.window = tk.Toplevel(parent)
        self.window.title("Просмотр базы данных")
        self.window.geometry("1200x800")
        
        self.db_manager = DatabaseManager()
        
        # Создаем интерфейс
        self.create_gui()
        
        # Загружаем данные
        self.load_data()
    
    def create_gui(self):
        """Создание графического интерфейса"""
        # Создаем фрейм с разделителем
        paned = ttk.PanedWindow(self.window, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Левая панель - список осужденных
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=2)
        
        # Правая панель - характеристики
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)
        
        # Настройка левой панели
        # Заголовок
        ttk.Label(left_frame, text="Список осужденных", font=("Arial", 12, "bold")).pack(pady=5)
        
        # Создаем Treeview для отображения данных
        columns = ("id", "name", "address", "phone")
        self.tree = ttk.Treeview(left_frame, columns=columns, show="headings")
        
        # Настраиваем заголовки
        self.tree.heading("id", text="ID")
        self.tree.heading("name", text="ФИО")
        self.tree.heading("address", text="Адрес")
        self.tree.heading("phone", text="Телефон")
        
        # Настраиваем ширину столбцов
        self.tree.column("id", width=50)
        self.tree.column("name", width=200)
        self.tree.column("address", width=300)
        self.tree.column("phone", width=150)
        
        # Добавляем скроллбар
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Размещаем элементы
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Привязываем событие выбора
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        
        # Настройка правой панели
        # Заголовок
        ttk.Label(right_frame, text="Характеристики", font=("Arial", 12, "bold")).pack(pady=5)
        
        # Фрейм для кнопок добавления характеристик
        buttons_frame = ttk.Frame(right_frame)
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(buttons_frame, text="Добавить положительную", 
                  command=lambda: self.add_characteristic("positive")).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="Добавить нейтральную", 
                  command=lambda: self.add_characteristic("neutral")).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="Добавить отрицательную", 
                  command=lambda: self.add_characteristic("negative")).pack(side=tk.LEFT, padx=2)
        
        # Фрейм для отображения характеристик
        characteristics_frame = ttk.LabelFrame(right_frame, text="Список характеристик")
        characteristics_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Создаем Treeview для характеристик
        char_columns = ("type", "text")
        self.char_tree = ttk.Treeview(characteristics_frame, columns=char_columns, show="headings")
        
        # Настраиваем заголовки
        self.char_tree.heading("type", text="Тип")
        self.char_tree.heading("text", text="Текст")
        
        # Настраиваем ширину столбцов
        self.char_tree.column("type", width=100)
        self.char_tree.column("text", width=300)
        
        # Добавляем скроллбар
        char_scrollbar = ttk.Scrollbar(characteristics_frame, orient=tk.VERTICAL, command=self.char_tree.yview)
        self.char_tree.configure(yscrollcommand=char_scrollbar.set)
        
        # Размещаем элементы
        self.char_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        char_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Кнопка удаления характеристики
        ttk.Button(characteristics_frame, text="Удалить выбранную", 
                  command=self.delete_characteristic).pack(pady=5)
    
    def load_data(self):
        """Загрузка данных из базы данных"""
        # Очищаем существующие данные
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Загружаем данные
        convicts = self.db_manager.get_convicts()
        for convict in convicts:
            self.tree.insert("", tk.END, values=(
                convict[0],  # id
                convict[6],  # full_name
                convict[7],  # address
                convict[8]   # phone
            ))
    
    def on_select(self, event):
        """Обработка выбора осужденного"""
        selection = self.tree.selection()
        if not selection:
            return
        
        # Получаем ID выбранного осужденного
        convict_id = self.tree.item(selection[0])["values"][0]
        
        # Загружаем характеристики
        self.load_characteristics(convict_id)
    
    def load_characteristics(self, convict_id):
        """Загрузка характеристик для выбранного осужденного"""
        # Очищаем существующие данные
        for item in self.char_tree.get_children():
            self.char_tree.delete(item)
        
        # Загружаем характеристики
        characteristics = self.db_manager.get_characteristics(convict_id=convict_id)
        for char in characteristics:
            type_text = {
                "positive": "Положительная",
                "neutral": "Нейтральная",
                "negative": "Отрицательная"
            }.get(char[2], char[2])
            
            self.char_tree.insert("", tk.END, values=(type_text, char[4]))
    
    def add_characteristic(self, char_type):
        """Добавление новой характеристики"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Сначала выберите осужденного")
            return
        
        # Создаем окно для ввода характеристики
        dialog = tk.Toplevel(self.window)
        dialog.title("Добавить характеристику")
        dialog.geometry("400x300")
        
        # Поле для ввода текста
        ttk.Label(dialog, text="Введите текст характеристики:").pack(pady=5)
        text_widget = tk.Text(dialog, wrap=tk.WORD, height=10)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        def save():
            text = text_widget.get("1.0", tk.END).strip()
            if not text:
                messagebox.showwarning("Предупреждение", "Введите текст характеристики")
                return
            
            convict_id = self.tree.item(selection[0])["values"][0]
            self.db_manager.add_characteristic(convict_id, char_type, text)
            
            # Обновляем список характеристик
            self.load_characteristics(convict_id)
            dialog.destroy()
        
        ttk.Button(dialog, text="Сохранить", command=save).pack(pady=5)
    
    def delete_characteristic(self):
        """Удаление выбранной характеристики"""
        selection = self.char_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите характеристику для удаления")
            return
        
        if messagebox.askyesno("Подтверждение", "Вы уверены, что хотите удалить эту характеристику?"):
            # TODO: Добавить метод удаления характеристики в DatabaseManager
            messagebox.showinfo("Информация", "Функция удаления характеристик будет добавлена позже") 