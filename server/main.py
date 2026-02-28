"""
Главный модуль сервера
"""
import socket
import json
import threading
import sys
import os
import hashlib
from pathlib import Path

# Добавляем корневую директорию проекта в путь
sys.path.insert(0, str(Path(__file__).parent.parent))

from common.path_helper import get_resource_path, get_base_path, ensure_directory
from common.database_manager import DatabaseManager
from server.request_handler import RequestHandler
from common.repositories import UserRepository
from database.init_db import init_database

def load_config():
    """Загружает конфигурацию сервера"""
    config_path = get_resource_path("config/server_config.json")
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def hash_password(password: str) -> str:
    """Хеширует пароль"""
    return hashlib.sha256(password.encode()).hexdigest()

def create_admin_user(db_manager: DatabaseManager):
    """Создает учетную запись администратора при первом запуске"""
    user_repo = UserRepository(db_manager)
    
    # Проверяем, есть ли уже пользователи
    existing = db_manager.fetch_all("SELECT * FROM users")
    if existing:
        print("Пользователи уже существуют в базе данных")
        return
    
    # Пробуем создать GUI окно для ввода данных администратора
    try:
        import tkinter as tk
        from tkinter import messagebox, ttk
        
        result = {'login': None, 'password': None, 'confirmed': False}
        
        def on_submit():
            login = login_entry.get().strip() or "admin"
            password = password_entry.get()
            password_confirm = password_confirm_entry.get()
            
            if not password:
                messagebox.showerror("Ошибка", "Пароль не может быть пустым!")
                return
            
            if password != password_confirm:
                messagebox.showerror("Ошибка", "Пароли не совпадают!")
                return
            
            if len(password) < 6:
                if not messagebox.askyesno("Предупреждение", 
                    "Пароль слишком короткий (менее 6 символов).\n"
                    "Рекомендуется использовать пароль длиной не менее 8 символов.\n\n"
                    "Продолжить с этим паролем?"):
                    return
            
            result['login'] = login
            result['password'] = password
            result['confirmed'] = True
            root.destroy()
        
        def on_cancel():
            if messagebox.askyesno("Подтверждение", 
                "Вы уверены, что хотите отменить создание администратора?\n"
                "Без администратора вы не сможете войти в систему."):
                root.destroy()
        
        root = tk.Tk()
        root.title("Создание учетной записи администратора")
        root.geometry("500x320")
        root.resizable(False, False)
        root.attributes('-topmost', True)  # Поверх всех окон
        
        # Центрируем окно
        root.update_idletasks()
        x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
        y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
        root.geometry(f"+{x}+{y}")
        
        # Заголовок
        title_label = tk.Label(root, text="Создание учетной записи администратора", 
                              font=("Arial", 12, "bold"))
        title_label.pack(pady=10)
        
        info_label = tk.Label(root, 
                             text="При первом запуске необходимо создать учетную запись администратора",
                             font=("Arial", 9))
        info_label.pack(pady=5)
        
        # Фрейм для полей ввода
        frame = ttk.Frame(root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Логин
        ttk.Label(frame, text="Логин:").grid(row=0, column=0, sticky=tk.W, pady=5)
        login_entry = ttk.Entry(frame, width=30)
        login_entry.insert(0, "admin")
        login_entry.grid(row=0, column=1, pady=5, padx=5)
        
        # Пароль
        ttk.Label(frame, text="Пароль:").grid(row=1, column=0, sticky=tk.W, pady=5)
        password_entry = ttk.Entry(frame, width=30, show="*")
        password_entry.grid(row=1, column=1, pady=5, padx=5)
        
        # Подтверждение пароля
        ttk.Label(frame, text="Подтверждение:").grid(row=2, column=0, sticky=tk.W, pady=5)
        password_confirm_entry = ttk.Entry(frame, width=30, show="*")
        password_confirm_entry.grid(row=2, column=1, pady=5, padx=5)
        
        # Подсказка
        hint_label = tk.Label(frame, 
                             text="Рекомендуется использовать сложный пароль\n(не менее 8 символов, буквы, цифры)",
                             font=("Arial", 8), fg="gray")
        hint_label.grid(row=3, column=0, columnspan=2, pady=5)
        
        # Кнопки
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        create_button = ttk.Button(button_frame, text="Создать", command=on_submit)
        create_button.pack(side=tk.LEFT, padx=5)
        cancel_button = ttk.Button(button_frame, text="Отмена", command=on_cancel)
        cancel_button.pack(side=tk.LEFT, padx=5)
        
        # Привязываем Enter к функции создания для всех полей
        login_entry.bind('<Return>', lambda e: on_submit())
        password_entry.bind('<Return>', lambda e: on_submit())
        password_confirm_entry.bind('<Return>', lambda e: on_submit())
        
        # Устанавливаем кнопку "Создать" как кнопку по умолчанию
        root.bind('<Return>', lambda e: on_submit())
        
        # Фокус на поле пароля
        password_entry.focus()
        
        root.mainloop()
        
        if result['confirmed'] and result['login'] and result['password']:
            password_hash = hash_password(result['password'])
            user_repo.create(result['login'], password_hash, "admin", "Системный администратор")
            print(f"Учетная запись администратора '{result['login']}' создана успешно!")
            messagebox.showinfo("Успех", 
                f"Учетная запись администратора '{result['login']}' создана успешно!\n\n"
                "Теперь вы можете использовать эти данные для входа в систему.")
        else:
            print("Создание администратора отменено")
            if not result['confirmed']:
                # Если пользователь закрыл окно без создания, используем консольный ввод как запасной вариант
                print("Использование консольного ввода...")
                print("Введите логин администратора (по умолчанию: admin):")
                login = input().strip() or "admin"
                print("Введите пароль администратора:")
                password = input().strip()
                
                if password:
                    password_hash = hash_password(password)
                    user_repo.create(login, password_hash, "admin", "Системный администратор")
                    print(f"Учетная запись администратора '{login}' создана успешно!")
                else:
                    print("Пароль не может быть пустым!")
    
    except Exception as e:
        # Если GUI недоступен, используем консольный ввод
        print(f"Не удалось открыть GUI окно: {e}")
        print("Использование консольного ввода...")
        print("Создание учетной записи администратора...")
        print("Введите логин администратора (по умолчанию: admin):")
        login = input().strip() or "admin"
        print("Введите пароль администратора:")
        password = input().strip()
        
        if not password:
            print("Пароль не может быть пустым!")
            return
        
        password_hash = hash_password(password)
        user_repo.create(login, password_hash, "admin", "Системный администратор")
        print(f"Учетная запись администратора '{login}' создана успешно!")

def handle_client(client_socket, address, db_path):
    """Обрабатывает подключение клиента"""
    # Создаем новое соединение с БД для этого потока
    db_manager = DatabaseManager(str(db_path))
    db_manager.connect()
    
    try:
        handler = RequestHandler(db_manager)
        client_socket.settimeout(10)
        
        # Пробуем определить формат запроса
        # Читаем первые байты
        first_bytes = client_socket.recv(20)
        if not first_bytes:
            return
        
        # Если первый символ - цифра, это запрос от агента с размером
        if first_bytes[0] in b'0123456789':
            # Это агент - читаем размер до символа новой строки
            size_line = first_bytes
            while b'\n' not in size_line:
                chunk = client_socket.recv(1)
                if not chunk:
                    return
                size_line += chunk
            
            try:
                # Извлекаем размер из строки
                size_str = size_line.decode('utf-8').split('\n')[0].strip()
                data_size = int(size_str)
            except (ValueError, IndexError):
                print(f"Ошибка: не удалось прочитать размер данных от {address}")
                return
            
            # Получаем данные по частям (учитываем, что часть данных уже в size_line)
            # Находим позицию после \n
            newline_pos = size_line.find(b'\n')
            if newline_pos >= 0:
                data = size_line[newline_pos + 1:]
            else:
                data = b''
            
            while len(data) < data_size:
                chunk = client_socket.recv(min(8192, data_size - len(data)))
                if not chunk:
                    break
                data += chunk
            
            request = json.loads(data.decode('utf-8'))
        else:
            # Это клиент - читаем JSON напрямую
            data = first_bytes
            
            # Получаем остальные данные и пробуем распарсить
            while True:
                try:
                    request = json.loads(data.decode('utf-8'))
                    break
                except json.JSONDecodeError:
                    # Получаем еще данные
                    chunk = client_socket.recv(8192)
                    if not chunk:
                        # Если данных больше нет, пробуем распарсить то, что есть
                        try:
                            request = json.loads(data.decode('utf-8'))
                            break
                        except:
                            raise ValueError("Не удалось распарсить JSON")
                    data += chunk
        
        # Определяем тип запроса
        if 'system_info' in request:
            # Запрос от агента
            response = handler.handle_agent_data(request)
            client_socket.sendall(response.encode('utf-8'))
        else:
            # Запрос от клиента
            response = handler.handle_client_request(request)
            client_socket.sendall(json.dumps(response).encode('utf-8'))
            
    except json.JSONDecodeError as e:
        error_msg = f"Ошибка парсинга JSON: {e}"
        print(f"Ошибка при обработке запроса от {address}: {error_msg}")
        try:
            client_socket.sendall(json.dumps({'status': 'ERROR', 'message': error_msg}).encode('utf-8'))
        except:
            pass
    except Exception as e:
        print(f"Ошибка при обработке запроса от {address}: {e}")
        import traceback
        traceback.print_exc()
        try:
            client_socket.sendall(json.dumps({'status': 'ERROR', 'message': str(e)}).encode('utf-8'))
        except:
            pass
    finally:
        client_socket.close()
        db_manager.disconnect()

def main():
    """Главная функция сервера"""
    print("Запуск сервера «Программный Страж»...")
    
    # Загружаем конфигурацию
    config = load_config()
    
    # Инициализируем базу данных
    # База данных всегда должна быть рядом с исполняемым файлом
    db_path_str = config['database_path']
    if not os.path.isabs(db_path_str):
        # Используем базовый путь (рядом с exe или корень проекта)
        base_path = get_base_path()
        db_path = base_path / db_path_str
    else:
        db_path = Path(db_path_str)
    
    ensure_directory(db_path)
    
    if not db_path.exists():
        print("Инициализация базы данных...")
        init_database(str(db_path))
        print("База данных создана")
    
    # Подключаемся к базе данных для проверки и создания администратора
    db_manager = DatabaseManager(str(db_path))
    db_manager.connect()
    
    # Создаем администратора при первом запуске (только если нет пользователей)
    existing_users = db_manager.fetch_all("SELECT * FROM users")
    if not existing_users:
        create_admin_user(db_manager)
    
    db_manager.disconnect()
    
    # Создаем сокет
    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    server_socket.bind((config['host'], config['port']))
    server_socket.listen(10)
    
    print(f"Сервер запущен на {config['host']}:{config['port']}")
    print("Ожидание подключений...")
    
    try:
        while True:
            client_socket, address = server_socket.accept()
            print(f"Подключение от {address}")
            
            # Обрабатываем в отдельном потоке, передаем путь к БД вместо менеджера
            thread = threading.Thread(
                target=handle_client,
                args=(client_socket, address, db_path)
            )
            thread.daemon = True
            thread.start()
            
    except KeyboardInterrupt:
        print("\nОстановка сервера...")
    finally:
        server_socket.close()

if __name__ == "__main__":
    main()

