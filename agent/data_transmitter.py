"""
Модуль для передачи данных на сервер
"""
import socket
import json
import time
from typing import Dict, Optional

class DataTransmitter:
    """Класс для передачи данных на сервер"""
    
    def __init__(self, server_host: str, server_port: int, retry_attempts: int = 3, retry_delay: int = 60):
        self.server_host = server_host
        self.server_port = server_port
        self.retry_attempts = retry_attempts
        self.retry_delay = retry_delay
    
    def send_data(self, system_info: Dict, software_list: list) -> bool:
        """Отправляет данные на сервер"""
        data = {
            'system_info': system_info,
            'software_list': software_list,
            'timestamp': time.time()
        }
        
        for attempt in range(self.retry_attempts):
            try:
                # Создаем сокет
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(30)  # Увеличиваем таймаут для больших данных
                
                # Подключаемся к серверу
                sock.connect((self.server_host, self.server_port))
                
                # Сериализуем данные с обработкой специальных символов
                try:
                    message = json.dumps(data, ensure_ascii=False).encode('utf-8')
                except Exception as e:
                    print(f"Ошибка при сериализации данных: {e}")
                    sock.close()
                    return False
                
                # Отправляем размер данных перед самими данными
                size = len(message)
                sock.sendall(f"{size}\n".encode('utf-8'))
                
                # Отправляем данные
                sock.sendall(message)
                
                # Получаем подтверждение
                response = sock.recv(1024).decode('utf-8')
                sock.close()
                
                if response == "OK":
                    return True
                else:
                    print(f"Сервер вернул ошибку: {response}")
                    
            except Exception as e:
                print(f"Ошибка при отправке данных (попытка {attempt + 1}/{self.retry_attempts}): {e}")
                if attempt < self.retry_attempts - 1:
                    time.sleep(self.retry_delay)
        
        return False

