"""
Главный модуль агента мониторинга
"""
import json
import time
import os
import sys
from pathlib import Path

# Добавляем корневую директорию проекта в путь
sys.path.insert(0, str(Path(__file__).parent.parent))

from common.path_helper import get_resource_path
from agent.software_scanner import SoftwareScanner
from agent.system_info_collector import SystemInfoCollector
from agent.data_transmitter import DataTransmitter

def load_config():
    """Загружает конфигурацию агента"""
    config_path = get_resource_path("config/agent_config.json")
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def main():
    """Главная функция агента"""
    print("Запуск агента мониторинга «Программный Страж»...")
    
    # Загружаем конфигурацию
    config = load_config()
    
    # Инициализируем компоненты
    scanner = SoftwareScanner()
    info_collector = SystemInfoCollector()
    transmitter = DataTransmitter(
        config['server_host'],
        config['server_port'],
        config.get('retry_attempts', 3),
        config.get('retry_delay', 60)
    )
    
    scan_interval = config.get('scan_interval', 3600)  # По умолчанию 1 час
    
    print(f"Интервал сканирования: {scan_interval} секунд")
    print(f"Сервер: {config['server_host']}:{config['server_port']}")
    
    while True:
        try:
            print("\nНачало сканирования...")
            
            # Собираем информацию о системе
            system_info = info_collector.collect_system_info()
            print(f"Рабочая станция: {system_info['computer_name']} ({system_info['ip_address']})")
            
            # Сканируем установленное ПО
            print("Сканирование установленного программного обеспечения...")
            software_list = scanner.scan_installed_software()
            print(f"Обнаружено программ: {len(software_list)}")
            
            # Отправляем данные на сервер
            print("Отправка данных на сервер...")
            if transmitter.send_data(system_info, software_list):
                print("Данные успешно отправлены на сервер")
            else:
                print("Ошибка при отправке данных на сервер")
            
            print(f"Ожидание следующего сканирования ({scan_interval} секунд)...")
            time.sleep(scan_interval)
            
        except KeyboardInterrupt:
            print("\nОстановка агента...")
            break
        except Exception as e:
            print(f"Критическая ошибка: {e}")
            time.sleep(60)  # Ждем минуту перед повтором

if __name__ == "__main__":
    main()

