"""
Обработчик запросов от клиентов и агентов
"""
import json
from typing import Dict, Optional
from common.database_manager import DatabaseManager
from common.repositories import (
    WorkstationRepository, SoftwareRepository, UserRepository,
    AllowedSoftwareRepository, ComplianceRepository
)
from server.compliance_checker import ComplianceChecker

class RequestHandler:
    """Класс для обработки запросов"""
    
    def __init__(self, db_manager: DatabaseManager):
        self.db = db_manager
        self.workstation_repo = WorkstationRepository(db_manager)
        self.software_repo = SoftwareRepository(db_manager)
        self.user_repo = UserRepository(db_manager)
        self.allowed_repo = AllowedSoftwareRepository(db_manager)
        self.compliance_repo = ComplianceRepository(db_manager)
        self.compliance_checker = ComplianceChecker(self.allowed_repo, self.compliance_repo)
    
    def handle_agent_data(self, data: Dict) -> str:
        """Обрабатывает данные от агента"""
        try:
            self.db.begin_transaction()
            
            system_info = data.get('system_info', {})
            software_list = data.get('software_list', [])
            
            # Создаем или обновляем рабочую станцию
            workstation_id = self.workstation_repo.create_or_update(
                system_info.get('ip_address'),
                system_info.get('username'),
                system_info.get('computer_name'),
                system_info.get('os_info'),
                system_info.get('department')
            )
            
            # Удаляем старое ПО для этой станции
            self.software_repo.delete_by_workstation(workstation_id)
            
            # Добавляем новое ПО
            software_id_map = {}
            for software in software_list:
                software_id = self.software_repo.create_or_update(
                    workstation_id,
                    software.get('name'),
                    software.get('version'),
                    software.get('manufacturer'),
                    software.get('install_date'),
                    software.get('size'),
                    software.get('executable_path'),
                    software.get('license_status')
                )
                software_key = f"{software.get('name')}_{software.get('version')}"
                software_id_map[software_key] = software_id
            
            # Проверяем соответствие
            violations = self.compliance_checker.check_software(workstation_id, software_list)
            if violations:
                self.compliance_checker.create_violations(violations, software_id_map)
            
            self.db.commit()
            return "OK"
            
        except Exception as e:
            self.db.rollback()
            return f"ERROR: {str(e)}"
    
    def handle_client_request(self, request: Dict) -> Dict:
        """Обрабатывает запрос от клиента"""
        request_type = request.get('type')
        
        if request_type == 'get_workstations':
            workstations = self.workstation_repo.get_all()
            return {'status': 'OK', 'data': workstations}
        
        elif request_type == 'get_software':
            workstation_id = request.get('workstation_id')
            if workstation_id:
                software = self.software_repo.get_by_workstation(workstation_id)
                return {'status': 'OK', 'data': software}
            return {'status': 'ERROR', 'message': 'workstation_id required'}
        
        elif request_type == 'get_violations':
            violations = self.compliance_repo.get_all()
            return {'status': 'OK', 'data': violations}
        
        elif request_type == 'delete_workstation':
            workstation_id = request.get('workstation_id')
            if workstation_id:
                if self.workstation_repo.delete(workstation_id):
                    return {'status': 'OK', 'message': 'Workstation deleted'}
                return {'status': 'ERROR', 'message': 'Failed to delete workstation'}
            return {'status': 'ERROR', 'message': 'workstation_id required'}
        
        elif request_type == 'authenticate':
            login = request.get('login')
            password_hash = request.get('password_hash')
            user = self.user_repo.get_by_login(login)
            
            if user and user['password_hash'] == password_hash:
                self.user_repo.update_last_login(user['id'])
                return {'status': 'OK', 'user': {
                    'id': user['id'],
                    'login': user['login'],
                    'access_level': user['access_level'],
                    'full_name': user['full_name']
                }}
            return {'status': 'ERROR', 'message': 'Invalid credentials'}
        
        elif request_type == 'get_allowed_software':
            allowed_software = self.allowed_repo.get_all()
            return {'status': 'OK', 'data': allowed_software}
        
        elif request_type == 'add_allowed_software':
            try:
                name = request.get('name')
                manufacturer = request.get('manufacturer')
                software_id = self.allowed_repo.create(
                    name=name,
                    manufacturer=manufacturer,
                    version_pattern=request.get('version_pattern'),
                    category=request.get('category'),
                    description=request.get('description')
                )
                self.compliance_repo.resolve_for_matching_software(name, manufacturer)
                return {'status': 'OK', 'data': {'id': software_id}}
            except Exception as e:
                return {'status': 'ERROR', 'message': str(e)}
        
        elif request_type == 'update_allowed_software':
            software_id = request.get('id')
            if not software_id:
                return {'status': 'ERROR', 'message': 'id required'}
            try:
                name = request.get('name')
                manufacturer = request.get('manufacturer')
                self.allowed_repo.update(
                    software_id=software_id,
                    name=name,
                    manufacturer=manufacturer,
                    version_pattern=request.get('version_pattern'),
                    category=request.get('category'),
                    description=request.get('description')
                )
                if name is None or manufacturer is None:
                    current = self.allowed_repo.get_by_id(software_id)
                    if current:
                        name = name if name is not None else current.get('name')
                        manufacturer = manufacturer if manufacturer is not None else current.get('manufacturer')
                if name:
                    self.compliance_repo.resolve_for_matching_software(name, manufacturer)
                return {'status': 'OK', 'message': 'Updated successfully'}
            except Exception as e:
                return {'status': 'ERROR', 'message': str(e)}
        
        elif request_type == 'delete_allowed_software':
            software_id = request.get('id')
            if not software_id:
                return {'status': 'ERROR', 'message': 'id required'}
            try:
                if self.allowed_repo.delete(software_id):
                    return {'status': 'OK', 'message': 'Deleted successfully'}
                return {'status': 'ERROR', 'message': 'Failed to delete'}
            except Exception as e:
                return {'status': 'ERROR', 'message': str(e)}
        
        # Запросы отчетов
        elif request_type == 'get_software_statistics':
            start_date = request.get('start_date')
            end_date = request.get('end_date')
            try:
                stats = self.software_repo.get_statistics(start_date, end_date)
                return {'status': 'OK', 'data': stats}
            except Exception as e:
                return {'status': 'ERROR', 'message': str(e)}
        
        elif request_type == 'get_violations_statistics':
            start_date = request.get('start_date')
            end_date = request.get('end_date')
            try:
                # Используем динамический расчет несоответствий путем сравнения всего ПО с разрешенным списком
                # Это обеспечивает актуальные данные независимо от записей в таблице compliance_violations
                stats = self.compliance_repo.get_dynamic_statistics(
                    self.software_repo,
                    self.allowed_repo,
                    self.compliance_checker
                )
                return {'status': 'OK', 'data': stats}
            except Exception as e:
                return {'status': 'ERROR', 'message': str(e)}
        
        elif request_type == 'get_employee_statistics':
            try:
                stats = self.workstation_repo.get_employee_statistics()
                return {'status': 'OK', 'data': stats}
            except Exception as e:
                return {'status': 'ERROR', 'message': str(e)}
        
        elif request_type == 'get_software_by_period':
            start_date = request.get('start_date')
            end_date = request.get('end_date')
            if not start_date or not end_date:
                return {'status': 'ERROR', 'message': 'start_date and end_date required'}
            try:
                software = self.software_repo.get_by_period(start_date, end_date)
                return {'status': 'OK', 'data': software}
            except Exception as e:
                return {'status': 'ERROR', 'message': str(e)}
        
        elif request_type == 'get_workstations_by_department':
            department = request.get('department')
            try:
                workstations = self.workstation_repo.get_by_department(department)
                return {'status': 'OK', 'data': workstations}
            except Exception as e:
                return {'status': 'ERROR', 'message': str(e)}
        
        else:
            return {'status': 'ERROR', 'message': 'Unknown request type'}

