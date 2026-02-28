"""
Клиент для взаимодействия с сервером
"""
import socket
import json
from typing import Dict, Optional

class NetworkClient:
    """Класс для взаимодействия с сервером"""
    
    def __init__(self, server_host: str, server_port: int):
        self.server_host = server_host
        self.server_port = server_port
    
    def send_request(self, request: Dict) -> Optional[Dict]:
        """Отправляет запрос на сервер и возвращает ответ"""
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(15)  # Увеличиваем таймаут
            sock.connect((self.server_host, self.server_port))
            
            # Отправляем запрос (клиентские запросы маленькие, отправляем без размера)
            message = json.dumps(request, ensure_ascii=False).encode('utf-8')
            sock.sendall(message)
            
            # Получаем ответ - сервер отправляет JSON напрямую
            response_data = b''
            sock.settimeout(5)  # Таймаут для получения ответа
            
            try:
                while True:
                    chunk = sock.recv(8192)
                    if not chunk:
                        break
                    response_data += chunk
                    # Пробуем распарсить JSON
                    try:
                        response = json.loads(response_data.decode('utf-8'))
                        sock.close()
                        return response
                    except json.JSONDecodeError:
                        # Продолжаем получать данные
                        continue
            except socket.timeout:
                # Если получили данные, пробуем распарсить
                if response_data:
                    try:
                        return json.loads(response_data.decode('utf-8'))
                    except:
                        pass
            
            sock.close()
            if response_data:
                return json.loads(response_data.decode('utf-8'))
            return {'status': 'ERROR', 'message': 'No response from server'}
        except socket.timeout:
            return {'status': 'ERROR', 'message': 'Connection timed out'}
        except Exception as e:
            return {'status': 'ERROR', 'message': str(e)}
    
    def authenticate(self, login: str, password: str) -> Optional[Dict]:
        """Аутентификация пользователя"""
        import hashlib
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        
        request = {
            'type': 'authenticate',
            'login': login,
            'password_hash': password_hash
        }
        return self.send_request(request)
    
    def get_workstations(self) -> Optional[Dict]:
        """Получает список рабочих станций"""
        request = {'type': 'get_workstations'}
        return self.send_request(request)
    
    def get_software(self, workstation_id: int) -> Optional[Dict]:
        """Получает список ПО для рабочей станции"""
        request = {
            'type': 'get_software',
            'workstation_id': workstation_id
        }
        return self.send_request(request)
    
    def get_violations(self) -> Optional[Dict]:
        """Получает список несоответствий"""
        request = {'type': 'get_violations'}
        return self.send_request(request)
    
    def delete_workstation(self, workstation_id: int) -> Optional[Dict]:
        """Удаляет рабочую станцию"""
        request = {
            'type': 'delete_workstation',
            'workstation_id': workstation_id
        }
        return self.send_request(request)
    
    def get_allowed_software(self) -> Optional[Dict]:
        """Получает список разрешенного ПО"""
        request = {'type': 'get_allowed_software'}
        return self.send_request(request)
    
    def add_allowed_software(self, name: str, manufacturer: str = None,
                            version_pattern: str = None, category: str = None,
                            description: str = None) -> Optional[Dict]:
        """Добавляет разрешенное ПО"""
        request = {
            'type': 'add_allowed_software',
            'name': name,
            'manufacturer': manufacturer,
            'version_pattern': version_pattern,
            'category': category,
            'description': description
        }
        return self.send_request(request)
    
    def update_allowed_software(self, software_id: int, name: str = None,
                               manufacturer: str = None, version_pattern: str = None,
                               category: str = None, description: str = None) -> Optional[Dict]:
        """Обновляет запись о разрешенном ПО"""
        request = {
            'type': 'update_allowed_software',
            'id': software_id,
            'name': name,
            'manufacturer': manufacturer,
            'version_pattern': version_pattern,
            'category': category,
            'description': description
        }
        return self.send_request(request)
    
    def delete_allowed_software(self, software_id: int) -> Optional[Dict]:
        """Удаляет запись о разрешенном ПО"""
        request = {
            'type': 'delete_allowed_software',
            'id': software_id
        }
        return self.send_request(request)
    
    def get_software_statistics(self, start_date: str = None, end_date: str = None) -> Optional[Dict]:
        """Получает статистику по программному обеспечению"""
        request = {
            'type': 'get_software_statistics',
            'start_date': start_date,
            'end_date': end_date
        }
        return self.send_request(request)
    
    def get_violations_statistics(self, start_date: str = None, end_date: str = None) -> Optional[Dict]:
        """Получает статистику по несоответствиям"""
        request = {
            'type': 'get_violations_statistics',
            'start_date': start_date,
            'end_date': end_date
        }
        return self.send_request(request)
    
    def get_employee_statistics(self) -> Optional[Dict]:
        """Получает статистику по сотрудникам"""
        request = {'type': 'get_employee_statistics'}
        return self.send_request(request)
    
    def get_software_by_period(self, start_date: str, end_date: str) -> Optional[Dict]:
        """Получает ПО за указанный период"""
        request = {
            'type': 'get_software_by_period',
            'start_date': start_date,
            'end_date': end_date
        }
        return self.send_request(request)
    
    def get_workstations_by_department(self, department: str = None) -> Optional[Dict]:
        """Получает рабочие станции по подразделению"""
        request = {
            'type': 'get_workstations_by_department',
            'department': department
        }
        return self.send_request(request)

