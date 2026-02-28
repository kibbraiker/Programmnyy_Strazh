"""
Модуль для сбора информации о системе и рабочей станции
"""
import socket
import platform
import os
from typing import Dict

class SystemInfoCollector:
    """Класс для сбора информации о системе"""
    
    def collect_system_info(self) -> Dict:
        """Собирает информацию о системе и рабочей станции"""
        info = {
            'ip_address': self._get_ip_address(),
            'username': os.getenv('USERNAME') or os.getenv('USER'),
            'computer_name': platform.node(),
            'os_info': self._get_os_info(),
        }
        return info
    
    def _get_ip_address(self) -> str:
        """Получает IP-адрес рабочей станции"""
        try:
            # Подключаемся к внешнему адресу для определения локального IP
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except:
            return "127.0.0.1"
    
    def _get_os_info(self) -> str:
        """Получает информацию об операционной системе"""
        os_name = platform.system()
        os_version = platform.version()
        os_arch = platform.machine()
        return f"{os_name} {os_version} ({os_arch})"

