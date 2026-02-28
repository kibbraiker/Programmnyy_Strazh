"""
Графический интерфейс клиентского приложения
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from pathlib import Path
import sys
from datetime import datetime, timedelta
import csv
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

# Импорт календаря с проверкой доступности
try:
    from tkcalendar import DateEntry
    CALENDAR_AVAILABLE = True
except ImportError:
    CALENDAR_AVAILABLE = False

sys.path.insert(0, str(Path(__file__).parent.parent))

from common.path_helper import get_resource_path

from client.network_client import NetworkClient

class LoginWindow:
    """Окно входа в систему"""
    
    def __init__(self, parent, on_success):
        self.parent = parent
        self.on_success = on_success
        self.window = tk.Toplevel(parent)
        self.window.title("Вход в систему")
        self.window.geometry("300x180")
        self.window.resizable(False, False)
        
        # Центрируем окно
        self.window.transient(parent)
        self.window.grab_set()
        
        # Поля ввода
        ttk.Label(self.window, text="Логин:").pack(pady=5)
        self.login_entry = ttk.Entry(self.window, width=25)
        self.login_entry.pack(pady=5)
        
        ttk.Label(self.window, text="Пароль:").pack(pady=5)
        self.password_entry = ttk.Entry(self.window, width=25, show="*")
        self.password_entry.pack(pady=5)
        
        # Кнопка входа
        login_button = ttk.Button(self.window, text="Войти", command=self.login)
        login_button.pack(pady=10, padx=20, fill=tk.X, ipadx=10)
        
        # Фокус на поле логина
        self.login_entry.focus()
        self.window.bind('<Return>', lambda e: self.login())
    
    def login(self):
        """Выполняет вход в систему"""
        login = self.login_entry.get()
        password = self.password_entry.get()
        
        if not login or not password:
            messagebox.showerror("Ошибка", "Введите логин и пароль")
            return
        
        # Загружаем конфигурацию
        config_path = get_resource_path("config/client_config.json")
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить конфигурацию: {e}")
            return
        
        try:
            client = NetworkClient(config['server_host'], config['server_port'])
            response = client.authenticate(login, password)
            
            if response and response.get('status') == 'OK':
                self.window.destroy()
                self.on_success(response.get('user'), client)
            else:
                error_msg = response.get('message', 'Неверный логин или пароль') if response else 'Не удалось подключиться к серверу'
                messagebox.showerror("Ошибка", error_msg)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка подключения: {e}\n\nУбедитесь, что сервер запущен!")


class SettingsWindow:
    """Окно настроек приложения"""
    
    def __init__(self, parent, network_client=None):
        self.parent = parent
        self.network_client = network_client
        self.window = tk.Toplevel(parent)
        self.window.title("Настройки")
        self.window.geometry("700x600")
        self.window.resizable(True, True)
        
        # Центрируем окно
        self.window.transient(parent)
        self.window.grab_set()
        
        # Загружаем конфигурацию
        self.config_path = get_resource_path("config/client_config.json")
        self.config = self.load_config()
        
        # Применяем тему
        self.apply_theme_to_window()
        
        # Создаем интерфейс
        self.create_interface()
        
        # Применяем тему к текстовым полям
        self.apply_theme_to_settings_widgets(self.window)
        
        # Фокус на окне
        self.window.focus()
    
    def apply_theme_to_window(self):
        """Применяет тему к окну настроек"""
        theme = self.config.get('theme', 'light')
        style = ttk.Style()
        
        if theme == 'dark':
            self.window.configure(bg='#2b2b2b')
            style.theme_use('clam')
            style.configure('TFrame', background='#2b2b2b')
            style.configure('TLabel', background='#2b2b2b', foreground='#ffffff')
            style.configure('TLabelFrame', background='#2b2b2b', foreground='#ffffff')
            style.configure('TLabelFrame.Label', background='#2b2b2b', foreground='#ffffff')
            style.configure('TButton', background='#3c3c3c', foreground='#ffffff')
            style.map('TButton', background=[('active', '#4c4c4c')])
            style.configure('TEntry', fieldbackground='#3c3c3c', foreground='#ffffff')
            style.configure('TCombobox', fieldbackground='#3c3c3c', foreground='#ffffff')
            style.configure('TNotebook', background='#2b2b2b')
            style.configure('TNotebook.Tab', background='#3c3c3c', foreground='#ffffff')
            style.map('TNotebook.Tab', background=[('selected', '#4c4c4c')])
            style.configure('Treeview', background='#3c3c3c', foreground='#ffffff', fieldbackground='#3c3c3c')
            style.configure('Treeview.Heading', background='#4c4c4c', foreground='#ffffff')
            style.map('Treeview', background=[('selected', '#555555')])
            style.configure('TCheckbutton', background='#2b2b2b', foreground='#ffffff')
            style.configure('TRadiobutton', background='#2b2b2b', foreground='#ffffff')
        else:
            self.window.configure(bg='SystemButtonFace')
            style.theme_use('clam')
            style.configure('TFrame', background='SystemButtonFace')
            style.configure('TLabel', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TLabelFrame', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TLabelFrame.Label', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TButton', background='SystemButtonFace', foreground='SystemWindowText')
            style.map('TButton', background=[('active', 'SystemHighlight')])
            style.configure('TEntry', fieldbackground='SystemWindow', foreground='SystemWindowText')
            style.configure('TCombobox', fieldbackground='SystemWindow', foreground='SystemWindowText')
            style.configure('TNotebook', background='SystemButtonFace')
            style.configure('TNotebook.Tab', background='SystemButtonFace', foreground='SystemWindowText')
            style.map('TNotebook.Tab', background=[('selected', 'SystemHighlight')])
            style.configure('Treeview', background='SystemWindow', foreground='SystemWindowText', fieldbackground='SystemWindow')
            style.configure('Treeview.Heading', background='SystemButtonFace', foreground='SystemWindowText')
            style.map('Treeview', background=[('selected', 'SystemHighlight')])
            style.configure('TCheckbutton', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TRadiobutton', background='SystemButtonFace', foreground='SystemWindowText')
        
        # Применяем тему к текстовым полям после создания интерфейса
        # (будет вызвано после create_interface)
    
    def apply_theme_to_settings_widgets(self, widget):
        """Рекурсивно применяет тему ко всем виджетам в окне настроек"""
        theme = self.config.get('theme', 'light')
        
        if isinstance(widget, tk.Text):
            if theme == 'dark':
                widget.configure(bg='#3c3c3c', fg='#ffffff', insertbackground='#ffffff')
            else:
                widget.configure(bg='SystemWindow', fg='SystemWindowText', insertbackground='SystemWindowText')
        elif isinstance(widget, tk.Frame):
            if theme == 'dark':
                widget.configure(bg='#2b2b2b')
            else:
                widget.configure(bg='SystemButtonFace')
        
        # Рекурсивно применяем к дочерним виджетам
        for child in widget.winfo_children():
            self.apply_theme_to_settings_widgets(child)
    
    def load_config(self):
        """Загружает конфигурацию из файла"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить конфигурацию: {e}")
            return {
                "server_host": "localhost",
                "server_port": 8888,
                "theme": "light",
                "auto_refresh": True,
                "refresh_interval": 30,
                "export_format": "csv",
                "export_path": ""
            }
    
    def save_config(self):
        """Сохраняет конфигурацию в файл"""
        try:
            original_config = dict(self.config)
            # Обновляем конфигурацию из полей ввода
            self.config['server_host'] = self.server_host_entry.get()
            try:
                self.config['server_port'] = int(self.server_port_entry.get())
            except ValueError:
                messagebox.showerror("Ошибка", "Порт должен быть числом")
                return False
            
            self.config['theme'] = self.theme_var.get()
            try:
                self.config['font_size'] = int(self.font_size_var.get())
            except ValueError:
                messagebox.showerror("Ошибка", "Размер шрифта должен быть числом")
                return False
            self.config['auto_refresh'] = self.auto_refresh_var.get()
            try:
                self.config['refresh_interval'] = int(self.refresh_interval_entry.get())
            except ValueError:
                messagebox.showerror("Ошибка", "Интервал обновления должен быть числом")
                return False

            # Если изменений нет — ничего не сохраняем и не путаем пользователя
            if self.config == original_config:
                messagebox.showinfo("Информация", "Изменений нет.")
                return True
            
            # Сохраняем в файл
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            
            # Применяем тему и размер шрифта сразу, если это главное окно
            if hasattr(self.parent, 'apply_theme'):
                self.parent.config = self.config
                self.parent.apply_theme()
                self.parent.apply_font_size()
                messagebox.showinfo("Успех", "Настройки сохранены!\n\nТема интерфейса и размер шрифта применены.")
            else:
                messagebox.showinfo("Успех", "Настройки сохранены!\n\nДля применения изменений перезапустите приложение.")
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить настройки: {e}")
            return False
    
    def create_interface(self):
        """Создает интерфейс настроек"""
        # Notebook для вкладок
        notebook = ttk.Notebook(self.window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Вкладка "Подключение"
        connection_frame = ttk.Frame(notebook)
        notebook.add(connection_frame, text="Подключение к серверу")
        self.create_connection_tab(connection_frame)
        
        # Вкладка "Интерфейс"
        interface_frame = ttk.Frame(notebook)
        notebook.add(interface_frame, text="Интерфейс")
        self.create_interface_tab(interface_frame)
        
        # Вкладка "Разрешенное ПО"
        if self.network_client:
            allowed_software_frame = ttk.Frame(notebook)
            notebook.add(allowed_software_frame, text="Управление разрешенным ПО")
            self.create_allowed_software_tab(allowed_software_frame)
        
        # Кнопки внизу окна
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(button_frame, text="Сохранить", command=self.save_and_close).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Отмена", command=self.window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def create_connection_tab(self, parent):
        """Создает вкладку настроек подключения"""
        frame = ttk.LabelFrame(parent, text="Параметры подключения", padding=10)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Хост сервера
        ttk.Label(frame, text="Адрес сервера:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.server_host_entry = ttk.Entry(frame, width=30)
        self.server_host_entry.insert(0, self.config.get('server_host', 'localhost'))
        self.server_host_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        
        # Порт сервера
        ttk.Label(frame, text="Порт сервера:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.server_port_entry = ttk.Entry(frame, width=30)
        self.server_port_entry.insert(0, str(self.config.get('server_port', 8888)))
        self.server_port_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        
        # Информация
        info_label = ttk.Label(
            frame, 
            text="Внимание: Изменение этих параметров требует\nперезапуска приложения для применения.",
            foreground="gray"
        )
        info_label.grid(row=2, column=0, columnspan=2, pady=10, sticky=tk.W)
    
    def create_interface_tab(self, parent):
        """Создает вкладку настроек интерфейса"""
        frame = ttk.LabelFrame(parent, text="Настройки отображения", padding=10)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Тема
        ttk.Label(frame, text="Тема интерфейса:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.theme_var = tk.StringVar(value=self.config.get('theme', 'light'))
        theme_frame = ttk.Frame(frame)
        theme_frame.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        ttk.Radiobutton(theme_frame, text="Светлая", variable=self.theme_var, value="light").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(theme_frame, text="Темная", variable=self.theme_var, value="dark").pack(side=tk.LEFT, padx=5)
        
        # Размер шрифта
        ttk.Label(frame, text="Размер шрифта:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.font_size_var = tk.StringVar(value=str(self.config.get('font_size', 10)))
        font_size_frame = ttk.Frame(frame)
        font_size_frame.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        font_sizes = ['8', '9', '10', '11', '12', '14', '16', '18', '20']
        self.font_size_combo = ttk.Combobox(font_size_frame, textvariable=self.font_size_var, 
                                            values=font_sizes, width=10, state='readonly')
        self.font_size_combo.pack(side=tk.LEFT, padx=5)
        ttk.Label(font_size_frame, text="пт").pack(side=tk.LEFT, padx=2)
        
        # Автообновление
        self.auto_refresh_var = tk.BooleanVar(value=self.config.get('auto_refresh', True))
        ttk.Checkbutton(
            frame, 
            text="Автоматическое обновление данных",
            variable=self.auto_refresh_var
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Интервал обновления
        ttk.Label(frame, text="Интервал обновления (секунды):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.refresh_interval_entry = ttk.Entry(frame, width=30)
        self.refresh_interval_entry.insert(0, str(self.config.get('refresh_interval', 30)))
        self.refresh_interval_entry.grid(row=3, column=1, sticky=tk.W, pady=5, padx=5)
        
        # Информация
        info_label = ttk.Label(
            frame, 
            text="Примечание: Автообновление работает только\nна главной странице приложения.",
            foreground="gray"
        )
        info_label.grid(row=4, column=0, columnspan=2, pady=10, sticky=tk.W)
    
    def create_allowed_software_tab(self, parent):
        """Создает вкладку управления разрешенным ПО"""
        # Кнопки управления
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(button_frame, text="➕ Добавить", command=self.add_allowed_software).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="✏️ Изменить", command=self.edit_allowed_software).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🗑 Удалить", command=self.delete_allowed_software).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="🔄 Обновить", command=self.refresh_allowed_software).pack(side=tk.LEFT, padx=5)
        
        # Фрейм для фильтров и поиска
        filter_frame = ttk.LabelFrame(parent, text="Фильтры и поиск", padding=10)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Поиск
        search_frame = ttk.Frame(filter_frame)
        search_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT, padx=5)
        self.allowed_software_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.allowed_software_search_var, width=25)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind('<KeyRelease>', lambda e: self.apply_allowed_software_filter())
        
        # Фильтр по производителю
        manufacturer_frame = ttk.Frame(filter_frame)
        manufacturer_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(manufacturer_frame, text="Производитель:").pack(side=tk.LEFT, padx=5)
        self.allowed_software_manufacturer_filter_var = tk.StringVar()
        self.allowed_software_manufacturer_filter_combo = ttk.Combobox(
            manufacturer_frame, 
            textvariable=self.allowed_software_manufacturer_filter_var, 
            width=25, 
            state='readonly'
        )
        self.allowed_software_manufacturer_filter_combo.pack(side=tk.LEFT, padx=5)
        
        # Фильтр по категории
        category_frame = ttk.Frame(filter_frame)
        category_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(category_frame, text="Категория:").pack(side=tk.LEFT, padx=5)
        self.allowed_software_category_filter_var = tk.StringVar()
        self.allowed_software_category_filter_combo = ttk.Combobox(
            category_frame, 
            textvariable=self.allowed_software_category_filter_var, 
            width=25, 
            state='readonly'
        )
        self.allowed_software_category_filter_combo.pack(side=tk.LEFT, padx=5)
        
        # Кнопка сброса фильтров
        def reset_filter():
            self.allowed_software_manufacturer_filter_var.set("")
            self.allowed_software_category_filter_var.set("")
            self.allowed_software_search_var.set("")
            self.apply_allowed_software_filter()
        
        ttk.Button(filter_frame, text="Сбросить фильтры", command=reset_filter).pack(side=tk.LEFT, padx=5)
        
        # Переменная для хранения всех элементов ПО (для фильтрации)
        self.all_allowed_software_items = []
        
        # Таблица разрешенного ПО
        table_frame = ttk.LabelFrame(parent, text="Перечень разрешенного ПО", padding=5)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("ID", "Название", "Производитель", "Версия/Шаблон", "Категория", "Описание")
        self.allowed_software_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        
        for col in columns:
            self.allowed_software_tree.heading(col, text=col)
            self.allowed_software_tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.allowed_software_tree.yview)
        self.allowed_software_tree.configure(yscrollcommand=scrollbar.set)
        
        self.allowed_software_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Привязываем изменение фильтров и поиска
        self.allowed_software_manufacturer_filter_var.trace('w', lambda *args: self.apply_allowed_software_filter())
        self.allowed_software_category_filter_var.trace('w', lambda *args: self.apply_allowed_software_filter())
        self.allowed_software_search_var.trace('w', lambda *args: self.apply_allowed_software_filter())
        
        # Загружаем данные
        self.refresh_allowed_software()
    
    def apply_allowed_software_filter(self):
        """Применяет фильтры и поиск для разрешенного ПО"""
        if not hasattr(self, 'all_allowed_software_items'):
            return
        
        if not hasattr(self, 'allowed_software_tree'):
            return
        
        # Безопасно получаем значения фильтров
        selected_manufacturer = ""
        selected_category = ""
        search_text = ""
        
        if hasattr(self, 'allowed_software_manufacturer_filter_var'):
            selected_manufacturer = self.allowed_software_manufacturer_filter_var.get()
        if hasattr(self, 'allowed_software_category_filter_var'):
            selected_category = self.allowed_software_category_filter_var.get()
        if hasattr(self, 'allowed_software_search_var'):
            search_text = self.allowed_software_search_var.get().lower().strip()
        
        # Очищаем таблицу
        for item in self.allowed_software_tree.get_children():
            self.allowed_software_tree.delete(item)
        
        # Добавляем элементы в соответствии с фильтрами
        for sw_data in self.all_allowed_software_items:
            manufacturer = sw_data.get('manufacturer', '') or ''
            category = sw_data.get('category', '') or ''
            
            # Проверяем фильтр по производителю
            manufacturer_match = not selected_manufacturer or manufacturer == selected_manufacturer
            
            # Проверяем фильтр по категории
            category_match = not selected_category or category == selected_category
            
            # Проверяем поиск (по всем полям)
            search_match = True
            if search_text:
                # Проверяем поиск во всех значениях строки
                values_str = ' '.join(str(v) for v in sw_data['values']).lower()
                search_match = search_text in values_str
            
            # Если все условия выполнены, добавляем элемент
            if manufacturer_match and category_match and search_match:
                self.allowed_software_tree.insert("", tk.END, values=sw_data['values'])
    
    def refresh_allowed_software(self):
        """Обновляет список разрешенного ПО"""
        if not self.network_client:
            return
        
        # Проверяем, что дерево создано
        if not hasattr(self, 'allowed_software_tree'):
            return
        
        # Очищаем текущие данные
        if not hasattr(self, 'all_allowed_software_items'):
            self.all_allowed_software_items = []
        self.all_allowed_software_items.clear()
        
        # Очищаем таблицу
        for item in self.allowed_software_tree.get_children():
            self.allowed_software_tree.delete(item)
        
        # Загружаем данные с сервера
        response = self.network_client.get_allowed_software()
        if response and response.get('status') == 'OK':
            software_list = response.get('data', [])
            manufacturers_set = set()
            categories_set = set()
            
            for sw in software_list:
                values = (
                    sw.get('id', ''),
                    sw.get('name', ''),
                    sw.get('manufacturer', '') or '',
                    sw.get('version_pattern', '') or '',
                    sw.get('category', '') or '',
                    sw.get('description', '') or ''
                )
                
                # Сохраняем данные для фильтрации
                self.all_allowed_software_items.append({
                    'values': values,
                    'manufacturer': sw.get('manufacturer', '') or '',
                    'category': sw.get('category', '') or ''
                })
                
                # Собираем уникальные значения для фильтров
                manufacturer = sw.get('manufacturer', '') or ''
                if manufacturer:
                    manufacturers_set.add(manufacturer)
                
                category = sw.get('category', '') or ''
                if category:
                    categories_set.add(category)
            
            # Обновляем списки фильтров
            if hasattr(self, 'allowed_software_manufacturer_filter_combo'):
                manufacturer_list = [''] + sorted(list(manufacturers_set))
                self.allowed_software_manufacturer_filter_combo['values'] = manufacturer_list
            
            if hasattr(self, 'allowed_software_category_filter_combo'):
                category_list = [''] + sorted(list(categories_set))
                self.allowed_software_category_filter_combo['values'] = category_list
            
            # Применяем фильтры (это заполнит таблицу)
            self.apply_allowed_software_filter()
    
    def add_allowed_software(self):
        """Открывает диалог добавления разрешенного ПО"""
        dialog = AllowedSoftwareDialog(self.window, self.network_client, None)
        if dialog.result:
            self.refresh_allowed_software()
    
    def edit_allowed_software(self):
        """Открывает диалог редактирования разрешенного ПО"""
        selection = self.allowed_software_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите запись для редактирования")
            return
        
        item = self.allowed_software_tree.item(selection[0])
        software_id = item['values'][0]
        
        # Получаем данные с сервера
        if not self.network_client:
            return
        
        response = self.network_client.get_allowed_software()
        if response and response.get('status') == 'OK':
            software_list = response.get('data', [])
            software = next((sw for sw in software_list if sw.get('id') == software_id), None)
            if software:
                dialog = AllowedSoftwareDialog(self.window, self.network_client, software)
                if dialog.result:
                    self.refresh_allowed_software()
    
    def delete_allowed_software(self):
        """Удаляет выбранное разрешенное ПО"""
        selection = self.allowed_software_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите запись для удаления")
            return
        
        item = self.allowed_software_tree.item(selection[0])
        software_id = item['values'][0]
        software_name = item['values'][1]
        
        confirm = messagebox.askyesno(
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить разрешенное ПО?\n\n{software_name}"
        )
        
        if not confirm:
            return
        
        if self.network_client:
            response = self.network_client.delete_allowed_software(software_id)
            if response and response.get('status') == 'OK':
                messagebox.showinfo("Успех", "Разрешенное ПО успешно удалено")
                self.refresh_allowed_software()
            else:
                error_msg = response.get('message', 'Ошибка при удалении') if response else 'Не удалось подключиться к серверу'
                messagebox.showerror("Ошибка", f"Не удалось удалить разрешенное ПО:\n{error_msg}")
    
    def save_and_close(self):
        """Сохраняет настройки и закрывает окно"""
        if self.save_config():
            self.window.destroy()


class AllowedSoftwareDialog:
    """Диалог для добавления/редактирования разрешенного ПО"""
    
    def __init__(self, parent, network_client, software_data=None):
        self.network_client = network_client
        self.result = False
        self.software_data = software_data
        
        self.window = tk.Toplevel(parent)
        self.window.title("Добавить разрешенное ПО" if not software_data else "Изменить разрешенное ПО")
        self.window.geometry("500x350")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        
        self.create_interface()
        self.window.focus()
    
    def create_interface(self):
        """Создает интерфейс диалога"""
        frame = ttk.Frame(self.window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Название
        ttk.Label(frame, text="Название программы *:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(frame, width=40)
        self.name_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=5)
        if self.software_data:
            self.name_entry.insert(0, self.software_data.get('name', ''))
        
        # Производитель
        ttk.Label(frame, text="Производитель:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.manufacturer_entry = ttk.Entry(frame, width=40)
        self.manufacturer_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=5)
        if self.software_data:
            self.manufacturer_entry.insert(0, self.software_data.get('manufacturer', '') or '')
        
        # Версия/Шаблон
        ttk.Label(frame, text="Версия/Шаблон:").grid(row=2, column=0, sticky=tk.W, pady=5)
        version_frame = ttk.Frame(frame)
        version_frame.grid(row=2, column=1, sticky=tk.W, pady=5, padx=5)
        self.version_entry = ttk.Entry(version_frame, width=40)
        self.version_entry.pack(side=tk.TOP, fill=tk.X)
        if self.software_data:
            self.version_entry.insert(0, self.software_data.get('version_pattern', '') or '')
        
        # Подсказка для шаблонов версий
        hint_text = (
            "Примеры: '2019' (точная версия), '2016-2021' (диапазон), "
            "'>=2019' (минимум), '<=2021' (максимум), '202*' (шаблон)"
        )
        hint_label = tk.Label(version_frame, text=hint_text, 
                             font=("Arial", 8), fg="gray", wraplength=320, justify=tk.LEFT)
        hint_label.pack(side=tk.TOP, anchor=tk.W, pady=(2, 0))
        
        # Категория
        ttk.Label(frame, text="Категория:").grid(row=3, column=0, sticky=tk.W, pady=5)
        category_values = ['Офисные приложения', 'Средства разработки', 'Системные утилиты', 
                          'Медиа-приложения', 'Браузеры', 'Антивирусы', 'Другое']
        self.category_entry = ttk.Combobox(frame, width=37, values=category_values, state='readonly')
        self.category_entry.grid(row=3, column=1, sticky=tk.W, pady=5, padx=5)
        if self.software_data:
            category_value = self.software_data.get('category', '') or ''
            if category_value:
                self.category_entry.set(category_value)
        
        # Описание
        ttk.Label(frame, text="Описание:").grid(row=4, column=0, sticky=tk.W+tk.N, pady=5)
        self.description_text = tk.Text(frame, width=40, height=5)
        self.description_text.grid(row=4, column=1, sticky=tk.W, pady=5, padx=5)
        if self.software_data:
            self.description_text.insert('1.0', self.software_data.get('description', '') or '')
        
        # Кнопки
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Сохранить", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Отмена", command=self.window.destroy).pack(side=tk.LEFT, padx=5)
    
    def save(self):
        """Сохраняет разрешенное ПО"""
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("Ошибка", "Название программы обязательно для заполнения")
            return
        
        manufacturer = self.manufacturer_entry.get().strip() or None
        version_pattern = self.version_entry.get().strip() or None
        category = self.category_entry.get().strip() or None
        description = self.description_text.get('1.0', tk.END).strip() or None
        
        if not self.network_client:
            messagebox.showerror("Ошибка", "Нет подключения к серверу")
            return
        
        try:
            if self.software_data:
                # Обновление
                response = self.network_client.update_allowed_software(
                    software_id=self.software_data['id'],
                    name=name,
                    manufacturer=manufacturer,
                    version_pattern=version_pattern,
                    category=category,
                    description=description
                )
            else:
                # Добавление
                response = self.network_client.add_allowed_software(
                    name=name,
                    manufacturer=manufacturer,
                    version_pattern=version_pattern,
                    category=category,
                    description=description
                )
            
            if response and response.get('status') == 'OK':
                messagebox.showinfo("Успех", "Разрешенное ПО успешно сохранено")
                self.result = True
                self.window.destroy()
            else:
                error_msg = response.get('message', 'Ошибка при сохранении') if response else 'Не удалось подключиться к серверу'
                messagebox.showerror("Ошибка", f"Не удалось сохранить разрешенное ПО:\n{error_msg}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при сохранении: {e}")


class ReportsWindow:
    """Окно отчетов"""
    
    def __init__(self, parent, network_client):
        self.parent = parent
        self.network_client = network_client
        self.window = tk.Toplevel(parent)
        self.window.title("Отчеты")
        self.window.geometry("900x700")
        self.window.resizable(True, True)
        
        # Загружаем конфигурацию для использования настроек экспорта
        self.config_path = get_resource_path("config/client_config.json")
        self.config = self.load_config()
        
        # Флаг для предотвращения вызовов методов загрузки во время инициализации
        self._initializing = False
        # Флаги для предотвращения одновременных вызовов методов загрузки
        self._loading_statistics = False
        self._loading_software_statistics = False
        self._loading_violations_statistics = False
        
        # Проверяем доступность календаря
        if not CALENDAR_AVAILABLE:
            messagebox.showwarning(
                "Предупреждение",
                "Библиотека tkcalendar не установлена.\n"
                "Для использования календарных виджетов установите её:\n"
                "pip install tkcalendar\n\n"
                "Пока будут использоваться текстовые поля."
            )
        
        # Центрируем окно
        self.window.transient(parent)
        
        self.create_interface()
        self.window.focus()
    
    def load_config(self):
        """Загружает конфигурацию из файла"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            return {
                "export_format": "csv",
                "export_path": ""
            }
    
    def get_export_filename(self, default_extension, file_types):
        """
        Получает имя файла для экспорта с учетом настроек
        
        Args:
            default_extension: Расширение файла по умолчанию (например, ".csv")
            file_types: Список типов файлов для диалога
        
        Returns:
            str: Путь к файлу или пустая строка, если пользователь отменил
        """
        import os
        
        # Получаем путь из настроек
        export_path = self.config.get('export_path', '').strip()
        
        # Генерируем имя файла по умолчанию с временной меткой
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"report_{timestamp}{default_extension}"
        
        # Если путь указан и существует, используем его как начальную директорию
        initialdir = None
        initialfile = default_filename
        
        if export_path:
            export_path_obj = Path(export_path)
            if export_path_obj.exists() and export_path_obj.is_dir():
                initialdir = str(export_path_obj)
                initialfile = default_filename
            elif export_path_obj.exists() and export_path_obj.is_file():
                # Если указан файл, используем его директорию
                initialdir = str(export_path_obj.parent)
                initialfile = export_path_obj.name
            else:
                # Путь не существует, пробуем создать директорию
                try:
                    export_path_obj.mkdir(parents=True, exist_ok=True)
                    if export_path_obj.is_dir():
                        initialdir = str(export_path_obj)
                        initialfile = default_filename
                except Exception:
                    # Не удалось создать директорию, используем стандартное поведение
                    pass
        
        # Параметры для диалога сохранения
        dialog_kwargs = {
            'defaultextension': default_extension,
            'filetypes': file_types,
            'initialfile': initialfile
        }
        
        if initialdir:
            dialog_kwargs['initialdir'] = initialdir
        
        filename = filedialog.asksaveasfilename(**dialog_kwargs)
        return filename
    
    def create_date_entry(self, parent, default_date=None, width=12):
        """Создает календарный виджет для выбора даты"""
        if CALENDAR_AVAILABLE:
            if default_date is None:
                default_date = datetime.now()
            elif isinstance(default_date, str):
                try:
                    default_date = datetime.strptime(default_date, "%Y-%m-%d")
                except:
                    default_date = datetime.now()
            
            date_entry = DateEntry(
                parent,
                width=width,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='yyyy-mm-dd',
                year=default_date.year,
                month=default_date.month,
                day=default_date.day
            )
            return date_entry
        else:
            # Fallback на обычное текстовое поле
            var = tk.StringVar(value=default_date.strftime("%Y-%m-%d") if isinstance(default_date, datetime) else (default_date or datetime.now().strftime("%Y-%m-%d")))
            entry = ttk.Entry(parent, textvariable=var, width=width)
            # Сохраняем ссылку на переменную для получения значения
            entry._date_var = var
            return entry
    
    def get_date_from_widget(self, widget):
        """Получает дату из виджета в формате строки YYYY-MM-DD"""
        if CALENDAR_AVAILABLE and isinstance(widget, DateEntry):
            return widget.get_date().strftime("%Y-%m-%d")
        else:
            # Для обычного Entry получаем значение через textvariable
            if hasattr(widget, '_date_var'):
                return widget._date_var.get()
            else:
                return widget.get()
    
    def _on_filter_change(self, callback):
        """Обработчик изменения фильтров, предотвращает вызовы во время инициализации"""
        if not self._initializing:
            callback()
    
    def _safe_load(self, flag_name, callback):
        """Безопасная загрузка данных с защитой от повторных вызовов"""
        if getattr(self, flag_name, False):
            return  # Уже выполняется
        try:
            setattr(self, flag_name, True)
            callback()
        finally:
            setattr(self, flag_name, False)
    
    def create_interface(self):
        """Создает интерфейс окна отчетов"""
        # Заголовок
        header_frame = ttk.Frame(self.window)
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        ttk.Label(header_frame, text="Отчеты системы", font=("Arial", 14, "bold")).pack()
        
        # Вкладки для разных типов отчетов
        notebook = ttk.Notebook(self.window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Вкладка "Общая статистика"
        stats_frame = ttk.Frame(notebook)
        notebook.add(stats_frame, text="Общая статистика")
        self.create_statistics_tab(stats_frame)
        
        # Вкладка "Статистика по ПО"
        software_frame = ttk.Frame(notebook)
        notebook.add(software_frame, text="Статистика по ПО")
        self.create_software_tab(software_frame)
        
        # Вкладка "Несоответствия"
        violations_frame = ttk.Frame(notebook)
        notebook.add(violations_frame, text="Несоответствия")
        self.create_violations_tab(violations_frame)
        
        # Вкладка "По сотрудникам"
        employees_frame = ttk.Frame(notebook)
        notebook.add(employees_frame, text="По сотрудникам")
        self.create_employees_tab(employees_frame)
        
        # Загружаем списки для фильтров после создания всех вкладок
        # Устанавливаем флаг инициализации, чтобы trace-обработчики не вызывали методы загрузки
        self._initializing = True
        self.load_filter_options()
        self._initializing = False
        
        # Загружаем данные после инициализации фильтров
        self.load_statistics()
        self.load_software_statistics()
        self.load_violations_statistics()
        
        # Кнопки экспорта
        button_frame = ttk.Frame(self.window)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Кнопка "Экспорт" с выпадающим меню
        export_menu = tk.Menu(self.window, tearoff=0)
        export_menu.add_command(label="Экспорт в CSV", command=self.export_csv)
        export_menu.add_command(label="Экспорт в JSON", command=self.export_json)
        export_menu.add_command(label="Экспорт в PDF", command=self.export_pdf)
        
        # Создаем кнопку с меню
        export_button = ttk.Menubutton(button_frame, text="Экспорт", direction='below')
        export_button.menu = export_menu
        export_button['menu'] = export_menu
        export_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Закрыть", command=self.window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def load_filter_options(self):
        """Загружает списки для фильтров подразделений и категорий ПО"""
        # Получаем список подразделений из рабочих станций
        workstations_response = self.network_client.get_workstations()
        departments = set()
        if workstations_response and workstations_response.get('status') == 'OK':
            workstations = workstations_response.get('data', [])
            for ws in workstations:
                dept = ws.get('department')
                if dept:
                    departments.add(dept)
        
        departments_list = ['Все подразделения'] + sorted(list(departments))
        
        # Получаем список категорий из разрешенного ПО
        allowed_software_response = self.network_client.get_allowed_software()
        categories = set()
        if allowed_software_response and allowed_software_response.get('status') == 'OK':
            allowed_list = allowed_software_response.get('data', [])
            for item in allowed_list:
                category = item.get('category')
                if category:
                    categories.add(category)
        
        categories_list = ['Все категории'] + sorted(list(categories))
        
        # Устанавливаем значения в комбобоксы
        if hasattr(self, 'department_filter_combo'):
            self.department_filter_combo['values'] = departments_list
            self.department_filter_combo.set('Все подразделения')
        
        if hasattr(self, 'category_filter_combo'):
            self.category_filter_combo['values'] = categories_list
            self.category_filter_combo.set('Все категории')
        
        if hasattr(self, 'sw_department_filter_combo'):
            self.sw_department_filter_combo['values'] = departments_list
            self.sw_department_filter_combo.set('Все подразделения')
        
        if hasattr(self, 'sw_category_filter_combo'):
            self.sw_category_filter_combo['values'] = categories_list
            self.sw_category_filter_combo.set('Все категории')
        
        if hasattr(self, 'v_department_filter_combo'):
            self.v_department_filter_combo['values'] = departments_list
            self.v_department_filter_combo.set('Все подразделения')
        
        if hasattr(self, 'v_category_filter_combo'):
            self.v_category_filter_combo['values'] = categories_list
            self.v_category_filter_combo.set('Все категории')
    
    def set_period_from_preset(self, preset_name, start_widget, end_widget):
        """Устанавливает период из предустановленного значения"""
        today = datetime.now()
        
        if preset_name == "За последний месяц":
            start_date = today - timedelta(days=30)
            end_date = today
        elif preset_name == "За последний квартал":
            start_date = today - timedelta(days=90)
            end_date = today
        elif preset_name == "За последний год":
            start_date = today - timedelta(days=365)
            end_date = today
        else:
            return  # Неизвестный период
        
        # Устанавливаем даты в виджеты
        if CALENDAR_AVAILABLE:
            # Для DateEntry из tkcalendar
            try:
                from tkcalendar import DateEntry
                if isinstance(start_widget, DateEntry):
                    start_widget.set_date(start_date)
                if isinstance(end_widget, DateEntry):
                    end_widget.set_date(end_date)
            except:
                # Если не удалось установить через set_date, пробуем через date
                try:
                    if hasattr(start_widget, 'date'):
                        start_widget.date = start_date
                    if hasattr(end_widget, 'date'):
                        end_widget.date = end_date
                except:
                    pass
        else:
            # Для текстовых полей (Entry)
            if isinstance(start_widget, tk.Entry) or isinstance(start_widget, ttk.Entry):
                if hasattr(start_widget, '_date_var'):
                    start_widget._date_var.set(start_date.strftime("%Y-%m-%d"))
                else:
                    start_widget.delete(0, tk.END)
                    start_widget.insert(0, start_date.strftime("%Y-%m-%d"))
            if isinstance(end_widget, tk.Entry) or isinstance(end_widget, ttk.Entry):
                if hasattr(end_widget, '_date_var'):
                    end_widget._date_var.set(end_date.strftime("%Y-%m-%d"))
                else:
                    end_widget.delete(0, tk.END)
                    end_widget.insert(0, end_date.strftime("%Y-%m-%d"))
    
    def create_statistics_tab(self, parent):
        """Создает вкладку общей статистики"""
        # Период
        period_frame = ttk.LabelFrame(parent, text="Период", padding=10)
        period_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Предустановленные периоды
        ttk.Label(period_frame, text="Быстрый выбор:").grid(row=0, column=0, padx=5, sticky=tk.W)
        preset_var = tk.StringVar()
        preset_combo = ttk.Combobox(period_frame, textvariable=preset_var, 
                                    values=["Выберите период...", "За последний месяц", 
                                           "За последний квартал", "За последний год"],
                                    state="readonly", width=20)
        preset_combo.set("Выберите период...")
        preset_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        
        def on_preset_change(event=None):
            selected = preset_var.get()
            if selected and selected != "Выберите период...":
                self.set_period_from_preset(selected, self.start_date_widget, self.end_date_widget)
                preset_combo.set("Выберите период...")  # Сбрасываем выбор
        
        preset_combo.bind("<<ComboboxSelected>>", on_preset_change)
        
        # Ручной ввод дат
        ttk.Label(period_frame, text="С:").grid(row=1, column=0, padx=5, pady=5)
        self.start_date_widget = self.create_date_entry(
            period_frame, 
            default_date=datetime.now() - timedelta(days=30),
            width=12
        )
        self.start_date_widget.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(period_frame, text="По:").grid(row=1, column=2, padx=5, pady=5)
        self.end_date_widget = self.create_date_entry(
            period_frame,
            default_date=datetime.now(),
            width=12
        )
        self.end_date_widget.grid(row=1, column=3, padx=5, pady=5)
        
        ttk.Button(period_frame, text="Обновить", command=self.load_statistics).grid(row=1, column=4, padx=10, pady=5)
        
        # Фильтры
        filters_frame = ttk.LabelFrame(parent, text="Фильтры", padding=10)
        filters_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Фильтр по подразделению
        ttk.Label(filters_frame, text="Подразделение:").grid(row=0, column=0, padx=5, sticky=tk.W)
        self.department_filter_var = tk.StringVar()
        self.department_filter_combo = ttk.Combobox(filters_frame, textvariable=self.department_filter_var, 
                                                    width=25, state='readonly')
        self.department_filter_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        self.department_filter_var.trace('w', lambda *args: self._on_filter_change(self.load_statistics))
        
        # Фильтр по категории ПО
        ttk.Label(filters_frame, text="Категория ПО:").grid(row=0, column=2, padx=5, sticky=tk.W)
        self.category_filter_var = tk.StringVar()
        self.category_filter_combo = ttk.Combobox(filters_frame, textvariable=self.category_filter_var, 
                                                  width=25, state='readonly')
        self.category_filter_combo.grid(row=0, column=3, padx=5, sticky=tk.W)
        self.category_filter_var.trace('w', lambda *args: self._on_filter_change(self.load_statistics))
        
        # Кнопка сброса фильтров
        def reset_filters():
            self.department_filter_var.set("")
            self.category_filter_var.set("")
            self.load_statistics()
        
        ttk.Button(filters_frame, text="Сбросить фильтры", command=reset_filters).grid(row=0, column=4, padx=5, sticky=tk.W)
        
        # Статистика
        stats_frame = ttk.LabelFrame(parent, text="Статистика", padding=10)
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Кнопка очистки статистики
        stats_button_frame = ttk.Frame(stats_frame)
        stats_button_frame.pack(fill=tk.X, pady=(0, 5))
        
        def clear_statistics():
            self.stats_text.config(state=tk.NORMAL)
            self.stats_text.delete('1.0', tk.END)
            self.stats_text.config(state=tk.DISABLED)
        
        ttk.Button(stats_button_frame, text="Очистить статистику", command=clear_statistics).pack(side=tk.LEFT, padx=5)
        
        self.stats_text = tk.Text(stats_frame, height=20, wrap=tk.WORD)
        stats_scroll = ttk.Scrollbar(stats_frame, orient=tk.VERTICAL, command=self.stats_text.yview)
        self.stats_text.configure(yscrollcommand=stats_scroll.set)
        self.stats_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        stats_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_software_tab(self, parent):
        """Создает вкладку статистики по ПО"""
        # Период
        period_frame = ttk.LabelFrame(parent, text="Период", padding=10)
        period_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Предустановленные периоды
        ttk.Label(period_frame, text="Быстрый выбор:").grid(row=0, column=0, padx=5, sticky=tk.W)
        sw_preset_var = tk.StringVar()
        sw_preset_combo = ttk.Combobox(period_frame, textvariable=sw_preset_var, 
                                       values=["Выберите период...", "За последний месяц", 
                                              "За последний квартал", "За последний год"],
                                       state="readonly", width=20)
        sw_preset_combo.set("Выберите период...")
        sw_preset_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        
        def on_sw_preset_change(event=None):
            selected = sw_preset_var.get()
            if selected and selected != "Выберите период...":
                self.set_period_from_preset(selected, self.sw_start_date_widget, self.sw_end_date_widget)
                sw_preset_combo.set("Выберите период...")
        
        sw_preset_combo.bind("<<ComboboxSelected>>", on_sw_preset_change)
        
        # Ручной ввод дат
        ttk.Label(period_frame, text="С:").grid(row=1, column=0, padx=5, pady=5)
        self.sw_start_date_widget = self.create_date_entry(
            period_frame,
            default_date=datetime.now() - timedelta(days=30),
            width=12
        )
        self.sw_start_date_widget.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(period_frame, text="По:").grid(row=1, column=2, padx=5, pady=5)
        self.sw_end_date_widget = self.create_date_entry(
            period_frame,
            default_date=datetime.now(),
            width=12
        )
        self.sw_end_date_widget.grid(row=1, column=3, padx=5, pady=5)
        
        ttk.Button(period_frame, text="Обновить", command=self.load_software_statistics).grid(row=1, column=4, padx=10, pady=5)
        
        # Фильтры
        sw_filters_frame = ttk.LabelFrame(parent, text="Фильтры", padding=10)
        sw_filters_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Фильтр по подразделению
        ttk.Label(sw_filters_frame, text="Подразделение:").grid(row=0, column=0, padx=5, sticky=tk.W)
        self.sw_department_filter_var = tk.StringVar()
        self.sw_department_filter_combo = ttk.Combobox(sw_filters_frame, textvariable=self.sw_department_filter_var, 
                                                       width=25, state='readonly')
        self.sw_department_filter_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        self.sw_department_filter_var.trace('w', lambda *args: self._on_filter_change(self.load_software_statistics))
        
        # Фильтр по категории ПО
        ttk.Label(sw_filters_frame, text="Категория ПО:").grid(row=0, column=2, padx=5, sticky=tk.W)
        self.sw_category_filter_var = tk.StringVar()
        self.sw_category_filter_combo = ttk.Combobox(sw_filters_frame, textvariable=self.sw_category_filter_var, 
                                                     width=25, state='readonly')
        self.sw_category_filter_combo.grid(row=0, column=3, padx=5, sticky=tk.W)
        self.sw_category_filter_var.trace('w', lambda *args: self._on_filter_change(self.load_software_statistics))
        
        # Таблица
        table_frame = ttk.LabelFrame(parent, text="Топ программ", padding=5)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("Название", "Производитель", "Количество установок")
        self.software_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        for col in columns:
            self.software_tree.heading(col, text=col)
            self.software_tree.column(col, width=200)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.software_tree.yview)
        self.software_tree.configure(yscrollcommand=scrollbar.set)
        self.software_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_violations_tab(self, parent):
        """Создает вкладку несоответствий"""
        # Период
        period_frame = ttk.LabelFrame(parent, text="Период", padding=10)
        period_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Предустановленные периоды
        ttk.Label(period_frame, text="Быстрый выбор:").grid(row=0, column=0, padx=5, sticky=tk.W)
        v_preset_var = tk.StringVar()
        v_preset_combo = ttk.Combobox(period_frame, textvariable=v_preset_var, 
                                     values=["Выберите период...", "За последний месяц", 
                                            "За последний квартал", "За последний год"],
                                     state="readonly", width=20)
        v_preset_combo.set("Выберите период...")
        v_preset_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        
        def on_v_preset_change(event=None):
            selected = v_preset_var.get()
            if selected and selected != "Выберите период...":
                self.set_period_from_preset(selected, self.v_start_date_widget, self.v_end_date_widget)
                v_preset_combo.set("Выберите период...")
        
        v_preset_combo.bind("<<ComboboxSelected>>", on_v_preset_change)
        
        # Ручной ввод дат
        ttk.Label(period_frame, text="С:").grid(row=1, column=0, padx=5, pady=5)
        self.v_start_date_widget = self.create_date_entry(
            period_frame,
            default_date=datetime.now() - timedelta(days=30),
            width=12
        )
        self.v_start_date_widget.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(period_frame, text="По:").grid(row=1, column=2, padx=5, pady=5)
        self.v_end_date_widget = self.create_date_entry(
            period_frame,
            default_date=datetime.now(),
            width=12
        )
        self.v_end_date_widget.grid(row=1, column=3, padx=5, pady=5)
        
        ttk.Button(period_frame, text="Обновить", command=self.load_violations_statistics).grid(row=1, column=4, padx=10, pady=5)
        
        # Фильтры
        v_filters_frame = ttk.LabelFrame(parent, text="Фильтры", padding=10)
        v_filters_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Фильтр по подразделению
        ttk.Label(v_filters_frame, text="Подразделение:").grid(row=0, column=0, padx=5, sticky=tk.W)
        self.v_department_filter_var = tk.StringVar()
        self.v_department_filter_combo = ttk.Combobox(v_filters_frame, textvariable=self.v_department_filter_var, 
                                                       width=25, state='readonly')
        self.v_department_filter_combo.grid(row=0, column=1, padx=5, sticky=tk.W)
        self.v_department_filter_var.trace('w', lambda *args: self._on_filter_change(self.load_violations_statistics))
        
        # Фильтр по категории ПО
        ttk.Label(v_filters_frame, text="Категория ПО:").grid(row=0, column=2, padx=5, sticky=tk.W)
        self.v_category_filter_var = tk.StringVar()
        self.v_category_filter_combo = ttk.Combobox(v_filters_frame, textvariable=self.v_category_filter_var, 
                                                     width=25, state='readonly')
        self.v_category_filter_combo.grid(row=0, column=3, padx=5, sticky=tk.W)
        self.v_category_filter_var.trace('w', lambda *args: self._on_filter_change(self.load_violations_statistics))
        
        # Статистика
        stats_frame = ttk.LabelFrame(parent, text="Статистика по несоответствиям", padding=10)
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.violations_text = tk.Text(stats_frame, height=20, wrap=tk.WORD)
        violations_scroll = ttk.Scrollbar(stats_frame, orient=tk.VERTICAL, command=self.violations_text.yview)
        self.violations_text.configure(yscrollcommand=violations_scroll.set)
        self.violations_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        violations_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_employees_tab(self, parent):
        """Создает вкладку статистики по сотрудникам"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Button(button_frame, text="Обновить", command=self.load_employee_statistics).pack(side=tk.LEFT, padx=5)
        
        # Таблица
        table_frame = ttk.LabelFrame(parent, text="Статистика по сотрудникам", padding=5)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("IP", "Пользователь", "Компьютер", "Подразделение", "ПО", "Несоответствия", "Последнее обновление")
        self.employees_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)
        for col in columns:
            self.employees_tree.heading(col, text=col)
            self.employees_tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.employees_tree.yview)
        self.employees_tree.configure(yscrollcommand=scrollbar.set)
        self.employees_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Загружаем данные
        self.load_employee_statistics()
    
    def load_statistics(self):
        """Загружает общую статистику"""
        if self._loading_statistics:
            return
        self._loading_statistics = True
        try:
            self.stats_text.delete('1.0', tk.END)
            self.stats_text.config(state=tk.NORMAL)
            
            start_date = self.get_date_from_widget(self.start_date_widget)
            end_date = self.get_date_from_widget(self.end_date_widget)
            
            # Получаем значения фильтров
            department_filter = None
            if hasattr(self, 'department_filter_var'):
                dept = self.department_filter_var.get()
                if dept and dept != 'Все подразделения':
                    department_filter = dept
            
            category_filter = None
            if hasattr(self, 'category_filter_var'):
                cat = self.category_filter_var.get()
                if cat and cat != 'Все категории':
                    category_filter = cat
            
            # Получаем статистику по ПО
            sw_stats = self.network_client.get_software_statistics(start_date, end_date)
            # Получаем статистику по несоответствиям
            v_stats = self.network_client.get_violations_statistics(start_date, end_date)
            # Получаем список рабочих станций
            workstations = self.network_client.get_workstations()
            
            text = "ОБЩАЯ СТАТИСТИКА\n"
            text += "=" * 50 + "\n\n"
            text += f"Период: {start_date} - {end_date}\n"
            if department_filter:
                text += f"Подразделение: {department_filter}\n"
            if category_filter:
                text += f"Категория ПО: {category_filter}\n"
            text += "\n"
            
            if workstations and workstations.get('status') == 'OK':
                ws_list = workstations.get('data', [])
                # Применяем фильтр по подразделению
                if department_filter:
                    ws_list = [ws for ws in ws_list if ws.get('department') == department_filter]
                total_ws = len(ws_list)
                active_ws = sum(1 for ws in ws_list if ws.get('agent_status') == 'active')
                text += f"Всего рабочих станций: {total_ws}\n"
                text += f"Активных агентов: {active_ws}\n\n"
            
            if sw_stats and sw_stats.get('status') == 'OK':
                data = sw_stats.get('data', {})
                text += "СТАТИСТИКА ПО ПРОГРАММНОМУ ОБЕСПЕЧЕНИЮ:\n"
                text += f"  Всего установок ПО: {data.get('total_software', 0)}\n"
                text += f"  Уникальных программ: {data.get('unique_software', 0)}\n\n"
            
            if v_stats and v_stats.get('status') == 'OK':
                data = v_stats.get('data', {})
                text += "СТАТИСТИКА ПО НЕСООТВЕТСТВИЯМ:\n"
                text += f"  Всего несоответствий: {data.get('total_violations', 0)}\n"
                
                by_status = data.get('by_status', {})
                if by_status:
                    text += "\n  По статусам:\n"
                    for status, count in by_status.items():
                        text += f"    {status}: {count}\n"
                
                by_type = data.get('by_type', {})
                if by_type:
                    text += "\n  По типам:\n"
                    for v_type, count in by_type.items():
                        text += f"    {v_type}: {count}\n"
            
            self.stats_text.insert('1.0', text)
            self.stats_text.config(state=tk.DISABLED)
        finally:
            self._loading_statistics = False
    
    def load_software_statistics(self):
        """Загружает статистику по ПО"""
        if self._loading_software_statistics:
            return
        self._loading_software_statistics = True
        try:
            for item in self.software_tree.get_children():
                self.software_tree.delete(item)
            
            start_date = self.get_date_from_widget(self.sw_start_date_widget)
            end_date = self.get_date_from_widget(self.sw_end_date_widget)
            
            # Получаем значения фильтров
            department_filter = None
            if hasattr(self, 'sw_department_filter_var'):
                dept = self.sw_department_filter_var.get()
                if dept and dept != 'Все подразделения':
                    department_filter = dept
            
            category_filter = None
            if hasattr(self, 'sw_category_filter_var'):
                cat = self.sw_category_filter_var.get()
                if cat and cat != 'Все категории':
                    category_filter = cat
            
            response = self.network_client.get_software_statistics(start_date, end_date)
            if response and response.get('status') == 'OK':
                data = response.get('data', {})
                top_software = data.get('top_software', [])
                
                # Применяем фильтры (базовая реализация - фильтрация на клиенте)
                # В полной реализации фильтрация должна выполняться на сервере
                filtered_software = []
                if department_filter or category_filter:
                    # Получаем список разрешенного ПО для фильтрации по категории
                    allowed_software_by_category = {}
                    if category_filter:
                        allowed_response = self.network_client.get_allowed_software()
                        if allowed_response and allowed_response.get('status') == 'OK':
                            allowed_list = allowed_response.get('data', [])
                            for item in allowed_list:
                                if item.get('category') == category_filter:
                                    name = item.get('name', '').lower()
                                    manufacturer = item.get('manufacturer', '').lower() if item.get('manufacturer') else None
                                    allowed_software_by_category[(name, manufacturer)] = item
                    
                    # Получаем список рабочих станций для фильтрации по подразделению
                    workstation_ids_by_dept = set()
                    if department_filter:
                        workstations_response = self.network_client.get_workstations()
                        if workstations_response and workstations_response.get('status') == 'OK':
                            workstations = workstations_response.get('data', [])
                            for ws in workstations:
                                if ws.get('department') == department_filter:
                                    workstation_ids_by_dept.add(ws.get('id'))
                    
                    # Фильтруем ПО
                    for sw in top_software:
                        # Проверяем фильтр по категории (упрощенная проверка)
                        if category_filter:
                            sw_name = sw.get('name', '').lower()
                            sw_manufacturer = sw.get('manufacturer', '').lower() if sw.get('manufacturer') else None
                            key = (sw_name, sw_manufacturer)
                            if key not in allowed_software_by_category:
                                continue
                        
                        filtered_software.append(sw)
                else:
                    filtered_software = top_software
                
                # Отображаем статистику
                for sw in filtered_software:
                    self.software_tree.insert("", tk.END, values=(
                        sw.get('name', ''),
                        sw.get('manufacturer', '') or '',
                        sw.get('install_count', 0)
                    ))
        finally:
            self._loading_software_statistics = False
    
    def load_violations_statistics(self):
        """Загружает статистику по несоответствиям"""
        if self._loading_violations_statistics:
            return
        self._loading_violations_statistics = True
        try:
            self.violations_text.delete('1.0', tk.END)
            self.violations_text.config(state=tk.NORMAL)
            
            start_date = self.get_date_from_widget(self.v_start_date_widget)
            end_date = self.get_date_from_widget(self.v_end_date_widget)
            
            # Получаем значения фильтров
            department_filter = None
            if hasattr(self, 'v_department_filter_var'):
                dept = self.v_department_filter_var.get()
                if dept and dept != 'Все подразделения':
                    department_filter = dept
            
            category_filter = None
            if hasattr(self, 'v_category_filter_var'):
                cat = self.v_category_filter_var.get()
                if cat and cat != 'Все категории':
                    category_filter = cat
            
            response = self.network_client.get_violations_statistics(start_date, end_date)
            text = "СТАТИСТИКА ПО НЕСООТВЕТСТВИЯМ\n"
            text += "=" * 50 + "\n\n"
            text += f"Период: {start_date} - {end_date}\n"
            if department_filter:
                text += f"Подразделение: {department_filter}\n"
            if category_filter:
                text += f"Категория ПО: {category_filter}\n"
            text += "\n"
            
            if response and response.get('status') == 'OK':
                data = response.get('data', {})
                
                # Применяем фильтры к данным несоответствий (базовая реализация)
                # В полной реализации фильтрация должна выполняться на сервере
                if department_filter or category_filter:
                    # Получаем полный список несоответствий для фильтрации
                    violations_response = self.network_client.get_violations()
                    violations_list = []
                    if violations_response and violations_response.get('status') == 'OK':
                        all_violations = violations_response.get('data', [])
                        
                        # Получаем список рабочих станций для фильтрации по подразделению
                        workstation_ids_by_dept = set()
                        if department_filter:
                            workstations_response = self.network_client.get_workstations()
                            if workstations_response and workstations_response.get('status') == 'OK':
                                workstations = workstations_response.get('data', [])
                                for ws in workstations:
                                    if ws.get('department') == department_filter:
                                        workstation_ids_by_dept.add(ws.get('id'))
                        
                        # Получаем список разрешенного ПО для фильтрации по категории
                        allowed_software_by_category = {}
                        if category_filter:
                            allowed_response = self.network_client.get_allowed_software()
                            if allowed_response and allowed_response.get('status') == 'OK':
                                allowed_list = allowed_response.get('data', [])
                                for item in allowed_list:
                                    if item.get('category') == category_filter:
                                        name = item.get('name', '').lower()
                                        manufacturer = item.get('manufacturer', '').lower() if item.get('manufacturer') else None
                                        allowed_software_by_category[(name, manufacturer)] = item
                        
                        # Фильтруем несоответствия
                        for violation in all_violations:
                            ws_id = violation.get('workstation_id')
                            
                            # Проверяем фильтр по подразделению
                            if department_filter and ws_id not in workstation_ids_by_dept:
                                continue
                            
                            # Для фильтра по категории нужна более сложная логика
                            # Упрощенная версия - пропускаем проверку категории для базовой реализации
                            violations_list.append(violation)
                    
                    # Пересчитываем статистику с учетом фильтров
                    total_violations = len(violations_list)
                    by_status = {}
                    by_type = {}
                    for v in violations_list:
                        status = v.get('status', 'new')
                        v_type = v.get('violation_type', 'unknown')
                        by_status[status] = by_status.get(status, 0) + 1
                        by_type[v_type] = by_type.get(v_type, 0) + 1
                else:
                    # Без фильтров используем данные из статистики
                    total_violations = data.get('total_violations', 0)
                    by_status = data.get('by_status', {})
                    by_type = data.get('by_type', {})
                
                text += f"Всего несоответствий: {total_violations}\n\n"
                
                if by_status:
                    text += "По статусам:\n"
                    for status, count in by_status.items():
                        text += f"  {status}: {count}\n"
                    text += "\n"
                
                if by_type:
                    text += "По типам нарушений:\n"
                    for v_type, count in by_type.items():
                        text += f"  {v_type}: {count}\n"
            else:
                text += "Не удалось загрузить данные"
            
            self.violations_text.insert('1.0', text)
            self.violations_text.config(state=tk.DISABLED)
        finally:
            self._loading_violations_statistics = False
    
    def load_employee_statistics(self):
        """Загружает статистику по сотрудникам"""
        for item in self.employees_tree.get_children():
            self.employees_tree.delete(item)
        
        response = self.network_client.get_employee_statistics()
        if response and response.get('status') == 'OK':
            employees = response.get('data', [])
            
            for emp in employees:
                self.employees_tree.insert("", tk.END, values=(
                    emp.get('ip_address', ''),
                    emp.get('username', '') or '',
                    emp.get('computer_name', '') or '',
                    emp.get('department', '') or '',
                    emp.get('software_count', 0),
                    emp.get('violations_count', 0),
                    emp.get('last_update', '') or ''
                ))
    
    def export_csv(self):
        """Экспортирует данные в CSV"""
        filename = self.get_export_filename(
            ".csv",
            [("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not filename:
            return
        
        # Проверяем существование файла и предупреждаем о перезаписи
        import os
        if os.path.exists(filename):
            if not messagebox.askyesno(
                "Подтверждение",
                f"Файл {os.path.basename(filename)} уже существует.\n"
                "Хотите перезаписать его?"
            ):
                return
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Общая статистика
                start_date = self.get_date_from_widget(self.start_date_widget)
                end_date = self.get_date_from_widget(self.end_date_widget)
                
                writer.writerow(["ОБЩАЯ СТАТИСТИКА"])
                writer.writerow(["Период", start_date, "по", end_date])
                writer.writerow([])
                
                # Рабочие станции
                workstations = self.network_client.get_workstations()
                if workstations and workstations.get('status') == 'OK':
                    ws_list = workstations.get('data', [])
                    writer.writerow(["РАБОЧИЕ СТАНЦИИ"])
                    writer.writerow(["IP", "Пользователь", "Компьютер", "ОС", "Статус", "Последнее обновление"])
                    for ws in ws_list:
                        writer.writerow([
                            ws.get('ip_address', ''),
                            ws.get('username', '') or '',
                            ws.get('computer_name', '') or '',
                            ws.get('os_info', '') or '',
                            ws.get('agent_status', ''),
                            ws.get('last_update', '') or ''
                        ])
                    writer.writerow([])
                
                # Статистика по ПО
                sw_stats = self.network_client.get_software_statistics(
                    start_date, end_date
                )
                if sw_stats and sw_stats.get('status') == 'OK':
                    data = sw_stats.get('data', {})
                    writer.writerow(["СТАТИСТИКА ПО ПО"])
                    writer.writerow(["Всего установок", data.get('total_software', 0)])
                    writer.writerow(["Уникальных программ", data.get('unique_software', 0)])
                    writer.writerow([])
                    writer.writerow(["Топ программ", "", ""])
                    writer.writerow(["Название", "Производитель", "Количество установок"])
                    for sw in data.get('top_software', []):
                        writer.writerow([
                            sw.get('name', ''),
                            sw.get('manufacturer', '') or '',
                            sw.get('install_count', 0)
                        ])
                    writer.writerow([])
                
                # Несоответствия
                v_stats = self.network_client.get_violations_statistics(
                    start_date, end_date
                )
                if v_stats and v_stats.get('status') == 'OK':
                    data = v_stats.get('data', {})
                    writer.writerow(["НЕСООТВЕТСТВИЯ"])
                    writer.writerow(["Всего", data.get('total_violations', 0)])
                    for status, count in data.get('by_status', {}).items():
                        writer.writerow([f"Статус: {status}", count])
                    for v_type, count in data.get('by_type', {}).items():
                        writer.writerow([f"Тип: {v_type}", count])
                
                messagebox.showinfo("Успех", f"Данные экспортированы в {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные: {e}")
    
    def export_json(self):
        """Экспортирует данные в JSON"""
        filename = self.get_export_filename(
            ".json",
            [("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not filename:
            return
        
        # Проверяем существование файла и предупреждаем о перезаписи
        import os
        if os.path.exists(filename):
            if not messagebox.askyesno(
                "Подтверждение",
                f"Файл {os.path.basename(filename)} уже существует.\n"
                "Хотите перезаписать его?"
            ):
                return
        
        try:
            start_date = self.get_date_from_widget(self.start_date_widget)
            end_date = self.get_date_from_widget(self.end_date_widget)
            
            data = {
                'period': {
                    'start': start_date,
                    'end': end_date
                },
                'generated_at': datetime.now().isoformat()
            }
            
            # Рабочие станции
            workstations = self.network_client.get_workstations()
            if workstations and workstations.get('status') == 'OK':
                data['workstations'] = workstations.get('data', [])
            
            # Статистика по ПО
            sw_stats = self.network_client.get_software_statistics(
                start_date, end_date
            )
            if sw_stats and sw_stats.get('status') == 'OK':
                data['software_statistics'] = sw_stats.get('data', {})
            
            # Статистика по несоответствиям
            v_stats = self.network_client.get_violations_statistics(
                start_date, end_date
            )
            if v_stats and v_stats.get('status') == 'OK':
                data['violations_statistics'] = v_stats.get('data', {})
            
            # Статистика по сотрудникам
            emp_stats = self.network_client.get_employee_statistics()
            if emp_stats and emp_stats.get('status') == 'OK':
                data['employee_statistics'] = emp_stats.get('data', [])
            
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            
            messagebox.showinfo("Успех", f"Данные экспортированы в {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные: {e}")
    
    def export_pdf(self):
        """Экспортирует данные в PDF"""
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.lib import colors
            from reportlab.lib.units import cm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            import os
            import platform
        except ImportError:
            messagebox.showerror(
                "Ошибка", 
                "Для экспорта в PDF требуется библиотека reportlab.\n"
                "Установите её командой: pip install reportlab"
            )
            return
        
        filename = self.get_export_filename(
            ".pdf",
            [("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if not filename:
            return
        
        # Проверяем существование файла и предупреждаем о перезаписи
        import os
        if os.path.exists(filename):
            if not messagebox.askyesno(
                "Подтверждение",
                f"Файл {os.path.basename(filename)} уже существует.\n"
                "Хотите перезаписать его?"
            ):
                return
        
        try:
            # Регистрируем шрифт с поддержкой кириллицы
            font_name = 'CyrillicFont'
            font_bold_name = 'CyrillicFontBold'
            font_registered = False
            
            # Пытаемся найти системные шрифты Windows с поддержкой кириллицы
            if platform.system() == 'Windows':
                font_paths = [
                    r'C:\Windows\Fonts\arial.ttf',
                    r'C:\Windows\Fonts\Arial.ttf',
                    r'C:\Windows\Fonts\arialbd.ttf',
                    r'C:\Windows\Fonts\Arialbd.ttf',
                ]
                
                arial_path = None
                arial_bold_path = None
                
                for path in font_paths:
                    if os.path.exists(path):
                        if 'bd' in path.lower() or 'bold' in path.lower():
                            arial_bold_path = path
                        else:
                            arial_path = path
                        if arial_path and arial_bold_path:
                            break
                
                if arial_path:
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, arial_path))
                        if arial_bold_path:
                            pdfmetrics.registerFont(TTFont(font_bold_name, arial_bold_path))
                        else:
                            # Если жирный шрифт не найден, используем обычный
                            pdfmetrics.registerFont(TTFont(font_bold_name, arial_path))
                        font_registered = True
                    except Exception as e:
                        pass
            
            # Если не удалось зарегистрировать системный шрифт, используем встроенные шрифты
            if not font_registered:
                # Используем стандартные шрифты reportlab, которые поддерживают Unicode
                # Но для кириллицы лучше использовать DejaVu Sans, если доступен
                try:
                    # Пытаемся использовать DejaVu Sans (часто доступен в системах)
                    dejavu_paths = [
                        '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
                        '/usr/share/fonts/TTF/DejaVuSans.ttf',
                        'C:\\Windows\\Fonts\\DejaVuSans.ttf',
                    ]
                    
                    for path in dejavu_paths:
                        if os.path.exists(path):
                            pdfmetrics.registerFont(TTFont(font_name, path))
                            pdfmetrics.registerFont(TTFont(font_bold_name, path))
                            font_registered = True
                            break
                except Exception:
                    pass
            
            # Если все еще не зарегистрирован, используем стандартные шрифты с Unicode
            if not font_registered:
                font_name = 'Helvetica'
                font_bold_name = 'Helvetica-Bold'
            
            doc = SimpleDocTemplate(filename, pagesize=A4)
            story = []
            styles = getSampleStyleSheet()
            
            # Создаем стили с правильным шрифтом
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontName=font_name,
                fontSize=10
            )
            
            heading2_style = ParagraphStyle(
                'CustomHeading2',
                parent=styles['Heading2'],
                fontName=font_bold_name,
                fontSize=12
            )
            
            # Заголовок
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontName=font_bold_name,
                fontSize=16,
                textColor=colors.HexColor('#1f4788'),
                spaceAfter=30,
                alignment=1  # Центрирование
            )
            story.append(Paragraph("ОТЧЕТ СИСТЕМЫ 'ПРОГРАММНЫЙ СТРАЖ'", title_style))
            story.append(Spacer(1, 0.5*cm))
            
            # Период
            start_date = self.get_date_from_widget(self.start_date_widget)
            end_date = self.get_date_from_widget(self.end_date_widget)
            
            period_text = f"Период: {start_date} - {end_date}"
            story.append(Paragraph(period_text, normal_style))
            story.append(Paragraph(f"Дата формирования: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", normal_style))
            story.append(Spacer(1, 0.5*cm))
            
            # Рабочие станции
            workstations = self.network_client.get_workstations()
            if workstations and workstations.get('status') == 'OK':
                ws_list = workstations.get('data', [])
                story.append(Paragraph("<b>РАБОЧИЕ СТАНЦИИ</b>", heading2_style))
                story.append(Spacer(1, 0.2*cm))
                
                total_ws = len(ws_list)
                active_ws = sum(1 for ws in ws_list if ws.get('agent_status') == 'active')
                story.append(Paragraph(f"Всего рабочих станций: {total_ws}", normal_style))
                story.append(Paragraph(f"Активных агентов: {active_ws}", normal_style))
                story.append(Spacer(1, 0.3*cm))
            
            # Статистика по ПО
            sw_stats = self.network_client.get_software_statistics(
                start_date, end_date
            )
            if sw_stats and sw_stats.get('status') == 'OK':
                data = sw_stats.get('data', {})
                story.append(Paragraph("<b>СТАТИСТИКА ПО ПРОГРАММНОМУ ОБЕСПЕЧЕНИЮ</b>", heading2_style))
                story.append(Spacer(1, 0.2*cm))
                story.append(Paragraph(f"Всего установок ПО: {data.get('total_software', 0)}", normal_style))
                story.append(Paragraph(f"Уникальных программ: {data.get('unique_software', 0)}", normal_style))
                story.append(Spacer(1, 0.3*cm))
                
                # Топ программ
                top_software = data.get('top_software', [])
                if top_software:
                    story.append(Paragraph("<b>Топ-10 программ по количеству установок:</b>", normal_style))
                    story.append(Spacer(1, 0.2*cm))
                    
                    table_data = [['Название', 'Производитель', 'Установок']]
                    for sw in top_software:
                        table_data.append([
                            sw.get('name', ''),
                            sw.get('manufacturer', '') or '',
                            str(sw.get('install_count', 0))
                        ])
                    
                    table = Table(table_data, colWidths=[6*cm, 5*cm, 2*cm])
                    table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                        ('FONTNAME', (0, 0), (-1, 0), font_bold_name),
                        ('FONTSIZE', (0, 0), (-1, 0), 10),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('FONTNAME', (0, 1), (-1, -1), font_name),
                        ('FONTSIZE', (0, 1), (-1, -1), 9),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ]))
                    story.append(table)
                    story.append(Spacer(1, 0.5*cm))
            
            # Несоответствия
            v_stats = self.network_client.get_violations_statistics(
                start_date, end_date
            )
            if v_stats and v_stats.get('status') == 'OK':
                data = v_stats.get('data', {})
                story.append(Paragraph("<b>НЕСООТВЕТСТВИЯ</b>", heading2_style))
                story.append(Spacer(1, 0.2*cm))
                story.append(Paragraph(f"Всего несоответствий: {data.get('total_violations', 0)}", normal_style))
                story.append(Spacer(1, 0.2*cm))
                
                by_status = data.get('by_status', {})
                if by_status:
                    story.append(Paragraph("<b>По статусам:</b>", normal_style))
                    for status, count in by_status.items():
                        story.append(Paragraph(f"  {status}: {count}", normal_style))
                    story.append(Spacer(1, 0.2*cm))
                
                by_type = data.get('by_type', {})
                if by_type:
                    story.append(Paragraph("<b>По типам нарушений:</b>", normal_style))
                    for v_type, count in by_type.items():
                        story.append(Paragraph(f"  {v_type}: {count}", normal_style))
            
            doc.build(story)
            messagebox.showinfo("Успех", f"Данные экспортированы в PDF: {filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось экспортировать в PDF: {e}")


class MainWindow:
    """Главное окно приложения"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Программный Страж - Контроль состава ПО")
        self.root.geometry("1000x600")
        
        self.user = None
        self.client = None
        self.workstation_ids = {}  # Словарь для хранения ID рабочих станций (ip -> id)
        self.all_workstations_data = []  # Сохраняем все данные рабочих станций для фильтрации
        self.violations_workstation_ids = set()  # Сохраняем ID станций с нарушениями
        
        # Загружаем конфигурацию и применяем тему и размер шрифта
        self.config_path = get_resource_path("config/client_config.json")
        self.config = self.load_config()
        self.apply_theme()
        self.apply_font_size()
        
        # Создаем меню
        self.create_menu()
        
        # Создаем интерфейс
        self.create_interface()
        
        # Обработчик закрытия окна
        self.root.protocol("WM_DELETE_WINDOW", self.exit_app)
        
        # Показываем окно входа
        self.show_login()
    
    def load_config(self):
        """Загружает конфигурацию из файла"""
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            return {
                "server_host": "localhost",
                "server_port": 8888,
                "theme": "light",
                "auto_refresh": True,
                "refresh_interval": 30,
                "export_format": "csv",
                "export_path": ""
            }
    
    def apply_theme(self):
        """Применяет тему интерфейса"""
        theme = self.config.get('theme', 'light')
        style = ttk.Style()
        
        if theme == 'dark':
            # Тёмная тема
            self.root.configure(bg='#2b2b2b')
            style.theme_use('clam')  # Используем базовую тему
            
            # Настраиваем цвета для тёмной темы
            style.configure('TFrame', background='#2b2b2b')
            style.configure('TLabel', background='#2b2b2b', foreground='#ffffff')
            style.configure('TLabelFrame', background='#2b2b2b', foreground='#ffffff')
            style.configure('TLabelFrame.Label', background='#2b2b2b', foreground='#ffffff')
            style.configure('TButton', background='#3c3c3c', foreground='#ffffff')
            style.map('TButton', background=[('active', '#4c4c4c')])
            style.configure('TEntry', fieldbackground='#3c3c3c', foreground='#ffffff')
            style.configure('TCombobox', fieldbackground='#3c3c3c', foreground='#ffffff')
            style.configure('TNotebook', background='#2b2b2b')
            style.configure('TNotebook.Tab', background='#3c3c3c', foreground='#ffffff')
            style.map('TNotebook.Tab', background=[('selected', '#4c4c4c')])
            style.configure('Treeview', background='#3c3c3c', foreground='#ffffff', fieldbackground='#3c3c3c')
            style.configure('Treeview.Heading', background='#4c4c4c', foreground='#ffffff')
            style.map('Treeview', background=[('selected', '#555555')])
            style.configure('Vertical.TScrollbar', background='#3c3c3c', troughcolor='#2b2b2b')
            style.configure('Horizontal.TScrollbar', background='#3c3c3c', troughcolor='#2b2b2b')
            style.configure('TCheckbutton', background='#2b2b2b', foreground='#ffffff')
            style.configure('TRadiobutton', background='#2b2b2b', foreground='#ffffff')
            style.configure('TText', background='#3c3c3c', foreground='#ffffff')
        else:
            # Светлая тема (по умолчанию)
            self.root.configure(bg='SystemButtonFace')
            style.theme_use('clam')
            
            # Сбрасываем настройки к стандартным
            style.configure('TFrame', background='SystemButtonFace')
            style.configure('TLabel', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TLabelFrame', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TLabelFrame.Label', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TButton', background='SystemButtonFace', foreground='SystemWindowText')
            style.map('TButton', background=[('active', 'SystemHighlight')])
            style.configure('TEntry', fieldbackground='SystemWindow', foreground='SystemWindowText')
            style.configure('TCombobox', fieldbackground='SystemWindow', foreground='SystemWindowText')
            style.configure('TNotebook', background='SystemButtonFace')
            style.configure('TNotebook.Tab', background='SystemButtonFace', foreground='SystemWindowText')
            style.map('TNotebook.Tab', background=[('selected', 'SystemHighlight')])
            style.configure('Treeview', background='SystemWindow', foreground='SystemWindowText', fieldbackground='SystemWindow')
            style.configure('Treeview.Heading', background='SystemButtonFace', foreground='SystemWindowText')
            style.map('Treeview', background=[('selected', 'SystemHighlight')])
            style.configure('Vertical.TScrollbar', background='SystemButtonFace', troughcolor='SystemButtonFace')
            style.configure('Horizontal.TScrollbar', background='SystemButtonFace', troughcolor='SystemButtonFace')
            style.configure('TCheckbutton', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TRadiobutton', background='SystemButtonFace', foreground='SystemWindowText')
            style.configure('TText', background='SystemWindow', foreground='SystemWindowText')
        
        # Применяем тему к текстовым полям (tk.Text)
        self.apply_theme_to_widgets(self.root)
        
        # Применяем размер шрифта
        self.apply_font_size()
    
    def apply_font_size(self):
        """Применяет размер шрифта к интерфейсу"""
        font_size = self.config.get('font_size', 10)
        style = ttk.Style()
        
        # Применяем размер шрифта к различным элементам
        style.configure('TLabel', font=('TkDefaultFont', font_size))
        style.configure('TButton', font=('TkDefaultFont', font_size))
        style.configure('TEntry', font=('TkDefaultFont', font_size))
        style.configure('TCombobox', font=('TkDefaultFont', font_size))
        style.configure('Treeview', font=('TkDefaultFont', font_size))
        style.configure('Treeview.Heading', font=('TkDefaultFont', font_size))
        style.configure('TCheckbutton', font=('TkDefaultFont', font_size))
        style.configure('TRadiobutton', font=('TkDefaultFont', font_size))
        style.configure('TLabelFrame', font=('TkDefaultFont', font_size))
        style.configure('TLabelFrame.Label', font=('TkDefaultFont', font_size))
        style.configure('TNotebook', font=('TkDefaultFont', font_size))
        style.configure('TNotebook.Tab', font=('TkDefaultFont', font_size))
        
        # Применяем размер шрифта к текстовым полям
        self.apply_font_size_to_widgets(self.root, font_size)
    
    def apply_font_size_to_widgets(self, widget, font_size):
        """Рекурсивно применяет размер шрифта ко всем виджетам"""
        if isinstance(widget, tk.Text):
            widget.configure(font=('TkDefaultFont', font_size))
        elif isinstance(widget, tk.Label):
            widget.configure(font=('TkDefaultFont', font_size))
        elif isinstance(widget, tk.Button):
            widget.configure(font=('TkDefaultFont', font_size))
        elif isinstance(widget, tk.Entry):
            widget.configure(font=('TkDefaultFont', font_size))
        
        # Рекурсивно применяем к дочерним виджетам
        for child in widget.winfo_children():
            self.apply_font_size_to_widgets(child, font_size)
    
    def apply_theme_to_widgets(self, widget):
        """Рекурсивно применяет тему ко всем виджетам"""
        theme = self.config.get('theme', 'light')
        
        if isinstance(widget, tk.Text):
            if theme == 'dark':
                widget.configure(bg='#3c3c3c', fg='#ffffff', insertbackground='#ffffff')
            else:
                widget.configure(bg='SystemWindow', fg='SystemWindowText', insertbackground='SystemWindowText')
        elif isinstance(widget, tk.Frame) or isinstance(widget, tk.Toplevel):
            if theme == 'dark':
                widget.configure(bg='#2b2b2b')
            else:
                widget.configure(bg='SystemButtonFace')
        
        # Рекурсивно применяем к дочерним виджетам
        for child in widget.winfo_children():
            self.apply_theme_to_widgets(child)
    
    def create_menu(self):
        """Создает меню навигации"""
        menu_frame = ttk.Frame(self.root)
        menu_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(menu_frame, text="ДОМОЙ", command=self.show_home).pack(side=tk.LEFT, padx=5)
        ttk.Button(menu_frame, text="ПРОГРАММЫ", command=self.show_programs).pack(side=tk.LEFT, padx=5)
        ttk.Button(menu_frame, text="ОТЧЁТЫ", command=self.show_reports).pack(side=tk.LEFT, padx=5)
        ttk.Button(menu_frame, text="НАСТРОЙКИ", command=self.show_settings).pack(side=tk.LEFT, padx=5)
        
        # Выход
        ttk.Button(menu_frame, text="ВЫХОД", command=self.exit_app).pack(side=tk.RIGHT, padx=5)

    def exit_app(self):
        """Закрывает приложение"""
        if messagebox.askyesno("Выход", "Выйти из приложения?"):
            try:
                self.root.quit()
                self.root.destroy()
            except Exception:
                pass
    
    def create_interface(self):
        """Создает основной интерфейс"""
        # Область для контента
        self.content_frame = ttk.Frame(self.root)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Статистика
        stats_frame = ttk.LabelFrame(self.content_frame, text="Статистика")
        stats_frame.pack(fill=tk.X, pady=5)
        
        # Фрейм для статистики и кнопки обновления
        stats_content_frame = ttk.Frame(stats_frame)
        stats_content_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.stats_label = ttk.Label(stats_content_frame, text="Войдите в систему для просмотра статистики")
        self.stats_label.pack(side=tk.LEFT, padx=5)
        
        # Кнопка обновления
        refresh_button = ttk.Button(stats_content_frame, text="🔄 Обновить", command=self.refresh_data)
        refresh_button.pack(side=tk.RIGHT, padx=5)
        
        # Поиск рабочей станции (в верхней части главной страницы)
        search_frame = ttk.LabelFrame(self.content_frame, text="Поиск рабочей станции")
        search_frame.pack(fill=tk.X, padx=5, pady=5)
        
        search_content_frame = ttk.Frame(search_frame)
        search_content_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(search_content_frame, text="Поиск:").pack(side=tk.LEFT, padx=5)
        self.search_entry = ttk.Entry(search_content_frame, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind('<KeyRelease>', self.on_search)
        
        # Кнопка очистки поиска
        clear_search_button = ttk.Button(search_content_frame, text="Очистить", command=self.clear_search)
        clear_search_button.pack(side=tk.LEFT, padx=5)
        
        # Основной контейнер: левая часть (таблица) и правая часть (диаграммы)
        main_container = ttk.Frame(self.content_frame)
        main_container.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Левая часть - таблица рабочих станций
        left_frame = ttk.Frame(main_container)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Таблица рабочих станций
        table_frame = ttk.LabelFrame(left_frame, text="Рабочие станции")
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Кнопка удаления над таблицей
        button_frame = ttk.Frame(table_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(button_frame, text="🗑 Удалить выбранную рабочую станцию", 
                  command=self.delete_selected_workstation).pack(side=tk.LEFT, padx=5)
        
        # Фрейм для таблицы и скроллбара
        tree_frame = ttk.Frame(table_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Создаем Treeview
        columns = ("IP", "Имя пользователя", "Компьютер", "Подразделение", "ОС", "Статус", "Последнее обновление")
        self.workstations_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
        
        for col in columns:
            self.workstations_tree.heading(col, text=col)
            self.workstations_tree.column(col, width=150)
        
        # Скроллбар
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.workstations_tree.yview)
        self.workstations_tree.configure(yscrollcommand=scrollbar.set)
        
        self.workstations_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Двойной клик для просмотра ПО
        self.workstations_tree.bind('<Double-1>', self.on_workstation_select)
        
        # Контекстное меню для удаления (правый клик)
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Удалить рабочую станцию", command=self.delete_selected_workstation)
        self.workstations_tree.bind('<Button-3>', self.show_context_menu)  # Правый клик
        
        # Правая часть - диаграммы
        right_frame = ttk.Frame(main_container)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, padx=(5, 0))
        right_frame.pack_propagate(False)  # Запрещаем изменение размера
        right_frame.config(width=400)
        
        # Диаграмма 1: Количество установленных агентов
        chart1_frame = ttk.LabelFrame(right_frame, text="Количество установленных агентов")
        chart1_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        self.agents_figure = Figure(figsize=(3.5, 3), dpi=80)
        self.agents_canvas = FigureCanvasTkAgg(self.agents_figure, chart1_frame)
        self.agents_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Диаграмма 2: Агенты имеющие потенциально опасное ПО
        chart2_frame = ttk.LabelFrame(right_frame, text="Агенты имеющие потенциально опасное ПО")
        chart2_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        self.violations_figure = Figure(figsize=(3.5, 3), dpi=80)
        self.violations_canvas = FigureCanvasTkAgg(self.violations_figure, chart2_frame)
        self.violations_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Инициализируем пустые диаграммы
        self.update_agents_chart([])
        self.update_violations_chart([], set())
        
        # Убеждаемся, что диаграммы видны и обновлены
        self.agents_canvas.draw()
        self.violations_canvas.draw()
        
        # Принудительно обновляем layout для правильного отображения
        right_frame.update_idletasks()
        
        # Кнопка "Просмотр программ" в нижней части страницы
        bottom_button_frame = ttk.Frame(self.content_frame)
        bottom_button_frame.pack(fill=tk.X, pady=10)
        ttk.Button(bottom_button_frame, text="Просмотр программ", 
                  command=self.show_programs).pack(pady=5)
    
    def show_login(self):
        """Показывает окно входа"""
        LoginWindow(self.root, self.on_login_success)
    
    def on_login_success(self, user, client):
        """Обработчик успешного входа"""
        self.user = user
        self.client = client
        self.refresh_data()
    
    def refresh_data(self):
        """Обновляет данные"""
        if not self.client:
            # Обновляем диаграммы с пустыми данными
            self.update_agents_chart([])
            self.update_violations_chart([], set())
            return
        
        # Получаем список рабочих станций
        response = self.client.get_workstations()
        if response and response.get('status') == 'OK':
            workstations = response.get('data', [])
            
            # Получаем список нарушений для выделения
            violations_response = self.client.get_violations()
            violations_workstation_ids = set()
            if violations_response and violations_response.get('status') == 'OK':
                violations = violations_response.get('data', [])
                violations_workstation_ids = {v.get('workstation_id') for v in violations if v.get('status') != 'resolved'}
            
            # Сохраняем данные для фильтрации
            self.all_workstations_data = workstations
            self.violations_workstation_ids = violations_workstation_ids
            
            # Применяем фильтр поиска, если он активен
            self.apply_search_filter()
            
            self.update_statistics(workstations)  # Это обновит диаграмму агентов
            self.update_violations_chart(workstations, violations_workstation_ids)  # Обновляем диаграмму нарушений
        else:
            # Если не удалось получить данные, показываем пустые диаграммы
            self.update_agents_chart([])
            self.update_violations_chart([], set())
    
    def update_workstations_table(self, workstations, violations_workstation_ids=None):
        """Обновляет таблицу рабочих станций"""
        if violations_workstation_ids is None:
            violations_workstation_ids = set()
            
        # Очищаем таблицу и словарь
        for item in self.workstations_tree.get_children():
            self.workstations_tree.delete(item)
        self.workstation_ids.clear()
        
        # Заполняем данными
        for ws in workstations:
            # Проверяем наличие несоответствий (красный цвет)
            tags = []
            ws_id = ws.get('id')
            
            if ws_id in violations_workstation_ids:
                tags.append('violation')  # Красный цвет для нарушений
            elif ws.get('agent_status') != 'active':
                tags.append('inactive')  # Светло-красный для неактивных
            
            ip_address = ws.get('ip_address', '')
            # Сохраняем ID для быстрого доступа
            self.workstation_ids[ip_address] = ws_id
            
            item = self.workstations_tree.insert("", tk.END, values=(
                ip_address,
                ws.get('username', ''),
                ws.get('computer_name', ''),
                ws.get('department', '') or '',
                ws.get('os_info', ''),
                ws.get('agent_status', ''),
                ws.get('last_update', '')
            ), tags=tags)
        
        # Настраиваем цвета
        self.workstations_tree.tag_configure('violation', background='#ff0000', foreground='white')  # Красный для нарушений
        self.workstations_tree.tag_configure('inactive', background='#ffcccc')  # Светло-красный для неактивных
    
    def update_statistics(self, workstations):
        """Обновляет статистику"""
        total = len(workstations)
        active = sum(1 for ws in workstations if ws.get('agent_status') == 'active')
        
        stats_text = f"Всего рабочих станций: {total} | Активных агентов: {active}"
        self.stats_label.config(text=stats_text)
        
        # Обновляем диаграммы
        self.update_agents_chart(workstations)
    
    def update_agents_chart(self, workstations):
        """Обновляет диаграмму количества установленных агентов"""
        active = sum(1 for ws in workstations if ws.get('agent_status') == 'active')
        inactive = len(workstations) - active
        
        # Очищаем предыдущую диаграмму
        self.agents_figure.clear()
        ax = self.agents_figure.add_subplot(111)
        
        total = len(workstations)
        if total > 0:
            # Формируем данные для диаграммы
            labels = []
            sizes = []
            colors = []
            explode = []
            
            if active > 0:
                labels.append('Активные')
                sizes.append(active)
                colors.append('#4CAF50')
                explode.append(0.05)
            
            if inactive > 0:
                labels.append('Неактивные')
                sizes.append(inactive)
                colors.append('#F44336')
                explode.append(0.05)
            
            if len(sizes) > 0:
                # Используем explode только если есть оба сегмента
                use_explode = explode if len(explode) > 1 else None
                ax.pie(sizes, explode=use_explode, labels=labels, colors=colors,
                      autopct='%1.1f%%', shadow=True, startangle=90)
            else:
                ax.text(0.5, 0.5, 'Нет данных', ha='center', va='center', fontsize=12, 
                       transform=ax.transAxes)
            
            ax.set_title('Количество установленных агентов', fontsize=10, fontweight='bold', pad=10)
        else:
            ax.text(0.5, 0.5, 'Нет данных', ha='center', va='center', fontsize=12,
                   transform=ax.transAxes)
            ax.set_title('Количество установленных агентов', fontsize=10, fontweight='bold', pad=10)
        
        # Убираем оси для круговой диаграммы
        ax.axis('equal')
        self.agents_figure.tight_layout()
        self.agents_canvas.draw()
    
    def update_violations_chart(self, workstations, violations_workstation_ids):
        """Обновляет диаграмму агентов с потенциально опасным ПО"""
        with_violations = len(violations_workstation_ids)
        without_violations = len(workstations) - with_violations
        
        # Очищаем предыдущую диаграмму
        self.violations_figure.clear()
        ax = self.violations_figure.add_subplot(111)
        
        total = len(workstations)
        if total > 0:
            # Формируем данные для диаграммы
            labels = []
            sizes = []
            colors = []
            explode = []
            
            if with_violations > 0:
                labels.append('С нарушениями')
                sizes.append(with_violations)
                colors.append('#F44336')
                explode.append(0.1)
            
            if without_violations > 0:
                labels.append('Без нарушений')
                sizes.append(without_violations)
                colors.append('#4CAF50')
                explode.append(0)
            
            if len(sizes) > 0:
                # Используем explode только если есть нарушения и оба сегмента
                use_explode = explode if len(explode) > 1 and with_violations > 0 else None
                ax.pie(sizes, explode=use_explode, labels=labels, colors=colors,
                      autopct='%1.1f%%', shadow=True, startangle=90)
            else:
                ax.text(0.5, 0.5, 'Нет данных', ha='center', va='center', fontsize=12,
                       transform=ax.transAxes)
            
            ax.set_title('Агенты с потенциально опасным ПО', fontsize=10, fontweight='bold', pad=10)
        else:
            ax.text(0.5, 0.5, 'Нет данных', ha='center', va='center', fontsize=12,
                   transform=ax.transAxes)
            ax.set_title('Агенты с потенциально опасным ПО', fontsize=10, fontweight='bold', pad=10)
        
        # Убираем оси для круговой диаграммы
        ax.axis('equal')
        self.violations_figure.tight_layout()
        self.violations_canvas.draw()
    
    def on_workstation_select(self, event):
        """Обработчик выбора рабочей станции"""
        selection = self.workstations_tree.selection()
        if not selection:
            return
        
        item = self.workstations_tree.item(selection[0])
        ip_address = item['values'][0]
        
        # Получаем ID из словаря
        workstation_id = self.workstation_ids.get(ip_address)
        if not workstation_id:
            messagebox.showerror("Ошибка", "Не удалось определить ID рабочей станции")
            return
        
        # Создаем объект рабочей станции для отображения
        workstation = {
            'id': workstation_id,
            'ip_address': ip_address,
            'username': item['values'][1],
            'computer_name': item['values'][2],
            'os_info': item['values'][4]  # Индекс изменился после добавления колонки "Подразделение"
        }
        self.show_software_window(workstation)
    
    def show_software_window(self, workstation):
        """Показывает окно с программным обеспечением"""
        window = tk.Toplevel(self.root)
        window.title(f"Программное обеспечение - {workstation.get('ip_address')}")
        window.geometry("1200x550")
        
        # Фрейм для фильтров
        filter_frame = ttk.LabelFrame(window, text="Фильтры и поиск", padding=10)
        filter_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Поиск
        search_frame = ttk.Frame(filter_frame)
        search_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(search_frame, text="Поиск:").pack(side=tk.LEFT, padx=5)
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=search_var, width=25)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind('<KeyRelease>', lambda e: apply_filter())
        
        # Фильтр по производителю
        manufacturer_frame = ttk.Frame(filter_frame)
        manufacturer_frame.pack(side=tk.LEFT, padx=5)
        ttk.Label(manufacturer_frame, text="Фильтр по производителю:").pack(side=tk.LEFT, padx=5)
        manufacturer_filter_var = tk.StringVar()
        manufacturer_filter_combo = ttk.Combobox(manufacturer_frame, textvariable=manufacturer_filter_var, 
                                                  width=30, state='readonly')
        manufacturer_filter_combo.pack(side=tk.LEFT, padx=5)
        
        # Кнопка сброса фильтров
        def reset_filter():
            manufacturer_filter_var.set("")
            search_var.set("")
            apply_filter()
        
        ttk.Button(filter_frame, text="Сбросить фильтры", command=reset_filter).pack(side=tk.LEFT, padx=5)
        
        # Таблица ПО
        columns = ("Название", "Версия", "Производитель", "Размер", "Дата установки", "Статус соответствия", "Путь")
        tree = ttk.Treeview(window, columns=columns, show="headings", height=20)
        
        # Переменная для хранения всех элементов ПО (для фильтрации)
        all_software_items = []
        
        # Переменная для хранения направления сортировки
        sort_reverse = {}
        
        def _parse_date_for_sort(date_str):
            """Парсит дату для сортировки"""
            if date_str == "N/A" or not date_str:
                return datetime.min
            try:
                # Пытаемся распарсить дату в формате "YYYY-MM-DD HH:MM"
                return datetime.strptime(date_str[:16], "%Y-%m-%d %H:%M")
            except:
                try:
                    # Пытаемся распарсить ISO формат
                    return datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                except:
                    return datetime.min
        
        def treeview_sort_column(col, reverse):
            """Сортирует дерево по указанной колонке"""
            # Получаем все элементы
            items = [(tree.set(item, col), item) for item in tree.get_children('')]
            
            # Определяем тип данных для сортировки
            try:
                # Пытаемся отсортировать как числа (для колонки "Размер")
                if col == "Размер":
                    items.sort(key=lambda t: float(t[0].split()[0]) if t[0] != "N/A" and len(t[0].split()) > 0 and t[0].split()[0].replace('.', '').isdigit() else 0, reverse=reverse)
                # Пытаемся отсортировать как даты (для колонки "Дата установки")
                elif col == "Дата установки":
                    items.sort(key=lambda t: _parse_date_for_sort(t[0]), reverse=reverse)
                else:
                    # Сортировка как строки
                    items.sort(key=lambda t: t[0].lower() if t[0] else '', reverse=reverse)
            except:
                # Если не удалось отсортировать, используем строковую сортировку
                items.sort(key=lambda t: str(t[0]).lower() if t[0] else '', reverse=reverse)
            
            # Перемещаем элементы в отсортированном порядке
            for index, (val, item) in enumerate(items):
                tree.move(item, '', index)
            
            # Меняем направление сортировки для следующего клика
            sort_reverse[col] = not reverse
            
            # Обновляем заголовок с указанием направления сортировки
            for c in columns:
                if c == col:
                    arrow = " ▼" if reverse else " ▲"
                    tree.heading(c, text=c + arrow, command=lambda _col=c: treeview_sort_column(_col, sort_reverse.get(_col, False)))
                else:
                    tree.heading(c, text=c, command=lambda _col=c: treeview_sort_column(_col, sort_reverse.get(_col, False)))
        
        # Настраиваем ширину колонок и привязываем сортировку
        for col in columns:
            tree.heading(col, text=col, command=lambda _col=col: treeview_sort_column(_col, sort_reverse.get(_col, False)))
            if col == "Название":
                tree.column(col, width=200)
            elif col == "Версия":
                tree.column(col, width=120)
            elif col == "Производитель":
                tree.column(col, width=150)
            elif col == "Размер":
                tree.column(col, width=100)
            elif col == "Дата установки":
                tree.column(col, width=150)
            elif col == "Статус соответствия":
                tree.column(col, width=150)
            elif col == "Путь":
                tree.column(col, width=300)
        
        scrollbar = ttk.Scrollbar(window, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        def apply_filter():
            """Применяет фильтры и поиск"""
            selected_manufacturer = manufacturer_filter_var.get()
            search_text = search_var.get().lower().strip()
            
            # Очищаем таблицу
            for item in tree.get_children():
                tree.delete(item)
            
            # Добавляем элементы в соответствии с фильтрами
            for sw_data in all_software_items:
                manufacturer = sw_data.get('manufacturer', '') or ''
                
                # Проверяем фильтр по производителю
                manufacturer_match = not selected_manufacturer or manufacturer == selected_manufacturer
                
                # Проверяем поиск (по всем полям)
                search_match = True
                if search_text:
                    # Проверяем поиск во всех значениях строки
                    values_str = ' '.join(str(v) for v in sw_data['values']).lower()
                    search_match = search_text in values_str
                
                # Если оба условия выполнены, добавляем элемент
                if manufacturer_match and search_match:
                    tree.insert("", tk.END, values=sw_data['values'], tags=sw_data['tags'])
        
        # Привязываем изменение фильтров и поиска
        manufacturer_filter_var.trace('w', lambda *args: apply_filter())
        search_var.trace('w', lambda *args: apply_filter())
        
        def load_software_data():
            """Загружает ПО и перечень разрешённого ПО с сервера и заполняет таблицу"""
            if not self.client:
                return
            
            # Очищаем текущие данные
            all_software_items.clear()
            
            response = self.client.get_software(workstation['id'])
            allowed_software_response = self.client.get_allowed_software()
            allowed_software_dict = {}
            if allowed_software_response and allowed_software_response.get('status') == 'OK':
                allowed_software_list = allowed_software_response.get('data', [])
                for item in allowed_software_list:
                    name = item.get('name', '').lower()
                    manufacturer = item.get('manufacturer', '').lower() if item.get('manufacturer') else None
                    key = (name, manufacturer)
                    allowed_software_dict[key] = item
                    if manufacturer:
                        key_name_only = (name, None)
                        if key_name_only not in allowed_software_dict:
                            allowed_software_dict[key_name_only] = item
            
            if response and response.get('status') == 'OK':
                software_list = response.get('data', [])
                manufacturers_set = set()  # Для сбора уникальных производителей
                
                for sw in software_list:
                    size = sw.get('size', 0)
                    size_str = f"{size / 1024 / 1024:.2f} МБ" if size else "N/A"
                    
                    # Форматируем дату установки
                    install_date = sw.get('install_date', '')
                    if install_date:
                        try:
                            # Пытаемся распарсить дату и отформатировать
                            if isinstance(install_date, str):
                                # Если дата в формате строки, пытаемся её распарсить
                                dt = None
                                # Пробуем разные форматы даты
                                date_formats = [
                                    '%Y-%m-%d %H:%M:%S',
                                    '%Y-%m-%dT%H:%M:%S',
                                    '%Y-%m-%dT%H:%M:%SZ',
                                    '%Y-%m-%d',
                                    '%d.%m.%Y %H:%M:%S',
                                    '%d.%m.%Y',
                                    '%Y%m%d%H%M%S',
                                    '%Y%m%d',
                                ]
                                
                                # Сначала пробуем ISO формат
                                try:
                                    dt = datetime.fromisoformat(install_date.replace('Z', '+00:00'))
                                except:
                                    # Пробуем другие форматы
                                    for fmt in date_formats:
                                        try:
                                            dt = datetime.strptime(install_date, fmt)
                                            break
                                        except:
                                            continue
                                
                                if dt:
                                    install_date_str = dt.strftime("%d.%m.%Y %H:%M")
                                else:
                                    # Если не удалось распарсить, пытаемся извлечь числа и добавить разделители
                                    # Убираем все нецифровые символы
                                    digits_only = ''.join(filter(str.isdigit, install_date))
                                    if len(digits_only) >= 8:  # Минимум YYYYMMDD
                                        try:
                                            # Пытаемся извлечь дату из строки цифр
                                            if len(digits_only) >= 14:  # YYYYMMDDHHMMSS
                                                dt = datetime.strptime(digits_only[:14], '%Y%m%d%H%M%S')
                                            elif len(digits_only) >= 8:  # YYYYMMDD
                                                dt = datetime.strptime(digits_only[:8], '%Y%m%d')
                                            install_date_str = dt.strftime("%d.%m.%Y %H:%M") if len(digits_only) >= 14 else dt.strftime("%d.%m.%Y")
                                        except:
                                            # Если не удалось, добавляем разделители вручную
                                            if len(digits_only) >= 8:
                                                year = digits_only[:4]
                                                month = digits_only[4:6]
                                                day = digits_only[6:8]
                                                if len(digits_only) >= 14:
                                                    hour = digits_only[8:10]
                                                    minute = digits_only[10:12]
                                                    install_date_str = f"{day}.{month}.{year} {hour}:{minute}"
                                                else:
                                                    install_date_str = f"{day}.{month}.{year}"
                                            else:
                                                install_date_str = install_date
                                    else:
                                        install_date_str = install_date
                            else:
                                # Если это не строка, конвертируем в строку и пытаемся распарсить
                                install_date_str = str(install_date)
                                try:
                                    dt = datetime.fromisoformat(install_date_str.replace('Z', '+00:00'))
                                    install_date_str = dt.strftime("%d.%m.%Y %H:%M")
                                except:
                                    pass
                        except Exception as e:
                            # В случае любой ошибки, пытаемся хотя бы добавить разделители
                            install_date_str = str(install_date) if install_date else "N/A"
                            # Если это строка из цифр, пытаемся добавить разделители
                            if isinstance(install_date, str):
                                digits_only = ''.join(filter(str.isdigit, install_date_str))
                                if len(digits_only) >= 8:
                                    try:
                                        year = digits_only[:4]
                                        month = digits_only[4:6]
                                        day = digits_only[6:8]
                                        if len(digits_only) >= 14:
                                            hour = digits_only[8:10]
                                            minute = digits_only[10:12]
                                            install_date_str = f"{day}.{month}.{year} {hour}:{minute}"
                                        elif len(digits_only) >= 8:
                                            install_date_str = f"{day}.{month}.{year}"
                                    except:
                                        pass
                    else:
                        install_date_str = "N/A"
                    
                    # Определяем статус соответствия по перечню разрешённого ПО (источник истины).
                    # Принцип "запрещено по умолчанию": ПО соответствует только если явно добавлено в список разрешённого.
                    # Если ПО в списке разрешённого — «Соответствует», даже при старых записях в violations.
                    sw_name = sw.get('name', '').lower()
                    sw_manufacturer = sw.get('manufacturer', '').lower() if sw.get('manufacturer') else None
                    key = (sw_name, sw_manufacturer)
                    key_name_only = (sw_name, None)
                    if key in allowed_software_dict or key_name_only in allowed_software_dict:
                        is_compliant = True  # В списке разрешённого — соответствует
                    else:
                        is_compliant = False  # Не в перечне — не соответствует (даже если перечень пуст)
                    
                    if is_compliant:
                        compliance_status = "Соответствует"
                        tags = []
                    else:
                        compliance_status = "Не соответствует"
                        tags = ['violation']
                    
                    manufacturer = sw.get('manufacturer', '') or ''
                    if manufacturer:
                        manufacturers_set.add(manufacturer)
                    
                    # Сохраняем данные элемента для фильтрации
                    item_values = (
                        sw.get('name', ''),
                        sw.get('version', ''),
                        manufacturer,
                        size_str,
                        install_date_str,
                        compliance_status,
                        sw.get('executable_path', '')
                    )
                    all_software_items.append({
                        'values': item_values,
                        'tags': tags,
                        'manufacturer': manufacturer
                    })
                
                # Заполняем список производителей в Combobox
                manufacturers_list = sorted(list(manufacturers_set))
                manufacturer_filter_combo['values'] = manufacturers_list
                
                # Настраиваем цвета для нарушений
                tree.tag_configure('violation', foreground='red')
                
                # Применяем фильтр для первоначального отображения
                apply_filter()
        
        # Кнопка обновления данных
        def reload_data():
            """Перезагружает данные с сервера и обновляет таблицу"""
            # Очищаем таблицу
            for item in tree.get_children():
                tree.delete(item)
            # Загружаем заново
            load_software_data()
        
        ttk.Button(filter_frame, text="🔄 Обновить", command=reload_data).pack(side=tk.LEFT, padx=5)
        
        # Загружаем данные при открытии окна
        load_software_data()
    
    def show_home(self):
        """Показывает главную страницу"""
        self.refresh_data()
    
    def show_programs(self):
        """Показывает страницу программ"""
        selection = self.workstations_tree.selection()
        if selection:
            self.on_workstation_select(None)
        else:
            messagebox.showinfo("Информация", "Выберите рабочую станцию из списка")
    
    def show_reports(self):
        """Показывает страницу отчетов"""
        if not self.client:
            messagebox.showwarning("Предупреждение", "Сначала войдите в систему")
            return
        ReportsWindow(self.root, self.client)
    
    def show_settings(self):
        """Показывает страницу настроек"""
        SettingsWindow(self.root, self.client)
    
    def apply_search_filter(self):
        """Применяет фильтр поиска к таблице рабочих станций"""
        # Если данных еще нет, ничего не делаем
        if not hasattr(self, 'all_workstations_data') or not self.all_workstations_data:
            return
        
        search_text = ""
        if hasattr(self, 'search_entry'):
            search_text = self.search_entry.get().lower().strip()
        
        # Фильтруем рабочие станции по поисковому запросу
        filtered_workstations = self.all_workstations_data
        if search_text:
            filtered_workstations = [
                ws for ws in self.all_workstations_data
                if any(search_text in str(v).lower() for v in [
                    ws.get('ip_address', ''),
                    ws.get('username', ''),
                    ws.get('computer_name', ''),
                    ws.get('department', '')
                ])
            ]
        
        # Обновляем таблицу с отфильтрованными данными
        violations_ids = getattr(self, 'violations_workstation_ids', set())
        self.update_workstations_table(filtered_workstations, violations_ids)
    
    def on_search(self, event):
        """Обработчик поиска рабочей станции"""
        # Применяем фильтр поиска
        self.apply_search_filter()
    
    def clear_search(self):
        """Очищает поле поиска"""
        if hasattr(self, 'search_entry'):
            self.search_entry.delete(0, tk.END)
            # Применяем фильтр (покажет все станции)
            self.apply_search_filter()
    
    def show_context_menu(self, event):
        """Показывает контекстное меню"""
        item = self.workstations_tree.identify_row(event.y)
        if item:
            self.workstations_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def delete_selected_workstation(self):
        """Удаляет выбранную рабочую станцию"""
        selection = self.workstations_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите рабочую станцию для удаления")
            return
        
        item = self.workstations_tree.item(selection[0])
        ip_address = item['values'][0]
        username = item['values'][1]
        computer_name = item['values'][2]
        
        # Получаем ID из словаря
        workstation_id = self.workstation_ids.get(ip_address)
        if not workstation_id:
            messagebox.showerror("Ошибка", "Не удалось определить ID рабочей станции")
            return
        
        # Подтверждение удаления
        confirm = messagebox.askyesno(
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить рабочую станцию?\n\n"
            f"IP: {ip_address}\n"
            f"Пользователь: {username}\n"
            f"Компьютер: {computer_name}\n\n"
            f"Внимание: Будут удалены все связанные данные о программном обеспечении!"
        )
        
        if not confirm:
            return
        
        # Удаляем через сервер
        if self.client:
            delete_response = self.client.delete_workstation(workstation_id)
            if delete_response and delete_response.get('status') == 'OK':
                messagebox.showinfo("Успех", "Рабочая станция успешно удалена")
                self.refresh_data()  # Обновляем список
            else:
                error_msg = delete_response.get('message', 'Ошибка при удалении') if delete_response else 'Не удалось подключиться к серверу'
                messagebox.showerror("Ошибка", f"Не удалось удалить рабочую станцию:\n{error_msg}")


def main():
    """Главная функция клиента"""
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()

