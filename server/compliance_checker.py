"""
Модуль для проверки соответствия установленного ПО перечню разрешенного
"""
import re
from typing import Dict, Optional, List
from difflib import SequenceMatcher
from common.repositories import AllowedSoftwareRepository, ComplianceRepository

class ComplianceChecker:
    """Класс для проверки соответствия установленного ПО"""
    
    def __init__(self, allowed_software_repo: AllowedSoftwareRepository,
                 compliance_repo: ComplianceRepository):
        self.allowed_repo = allowed_software_repo
        self.compliance_repo = compliance_repo
    
    def _fuzzy_match(self, str1: str, str2: str, threshold: float = 0.8) -> bool:
        """Нечеткое сопоставление строк с использованием SequenceMatcher"""
        if not str1 or not str2:
            return False
        similarity = SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
        return similarity >= threshold
    
    def _normalize_version(self, version: str) -> List[int]:
        """Нормализует версию в список чисел для сравнения"""
        if not version:
            return []
        # Извлекаем числа из версии (например, "1.2.3" -> [1, 2, 3])
        parts = re.findall(r'\d+', version)
        return [int(p) for p in parts]
    
    def _parse_version_range(self, pattern: str) -> Optional[Dict]:
        """Парсит диапазон версий (например, "2016-2021" или ">=2019")"""
        if not pattern:
            return None
        
        pattern = pattern.strip()
        
        # Диапазон вида "2016-2021" или "2016..2021"
        range_match = re.match(r'(\d+(?:\.\d+)*)\s*[-.]{2,3}\s*(\d+(?:\.\d+)*)', pattern)
        if range_match:
            return {
                'type': 'range',
                'min': self._normalize_version(range_match.group(1)),
                'max': self._normalize_version(range_match.group(2))
            }
        
        # Минимальная версия ">=2019" или ">= 2019"
        min_match = re.match(r'>=\s*(\d+(?:\.\d+)*)', pattern)
        if min_match:
            return {
                'type': 'min',
                'version': self._normalize_version(min_match.group(1))
            }
        
        # Максимальная версия "<=2021" или "<= 2021"
        max_match = re.match(r'<=\s*(\d+(?:\.\d+)*)', pattern)
        if max_match:
            return {
                'type': 'max',
                'version': self._normalize_version(max_match.group(1))
            }
        
        # Шаблон с звездочкой "202*" или "20.*"
        if '*' in pattern:
            return {
                'type': 'pattern',
                'pattern': pattern.replace('*', '.*')
            }
        
        return None
    
    def _check_version_pattern(self, installed_version: str, pattern: str) -> bool:
        """Проверяет соответствие версии шаблону"""
        if not pattern:
            return True  # Если шаблон не указан, версия не проверяется
        
        installed_normalized = self._normalize_version(installed_version)
        if not installed_normalized:
            return False
        
        # Парсим диапазон или шаблон
        version_range = self._parse_version_range(pattern)
        
        if version_range:
            if version_range['type'] == 'range':
                min_ver = version_range['min']
                max_ver = version_range['max']
                return self._compare_versions(installed_normalized, min_ver) >= 0 and \
                       self._compare_versions(installed_normalized, max_ver) <= 0
            elif version_range['type'] == 'min':
                return self._compare_versions(installed_normalized, version_range['version']) >= 0
            elif version_range['type'] == 'max':
                return self._compare_versions(installed_normalized, version_range['version']) <= 0
            elif version_range['type'] == 'pattern':
                # Регулярное выражение для шаблона
                try:
                    pattern_regex = re.compile(version_range['pattern'], re.IGNORECASE)
                    return bool(pattern_regex.match(installed_version))
                except:
                    return False
        
        # Точное совпадение или нечеткое сопоставление
        if installed_version.lower() == pattern.lower():
            return True
        
        # Проверяем нечеткое совпадение версий
        return self._fuzzy_match(installed_version, pattern, threshold=0.7)
    
    def _compare_versions(self, v1: List[int], v2: List[int]) -> int:
        """Сравнивает две нормализованные версии. Возвращает -1, 0 или 1"""
        max_len = max(len(v1), len(v2))
        v1_padded = v1 + [0] * (max_len - len(v1))
        v2_padded = v2 + [0] * (max_len - len(v2))
        
        for i in range(max_len):
            if v1_padded[i] < v2_padded[i]:
                return -1
            elif v1_padded[i] > v2_padded[i]:
                return 1
        return 0
    
    def _find_matching_allowed_software(self, name: str, manufacturer: Optional[str], 
                                       version: Optional[str], allowed_list: List[Dict]) -> Optional[Dict]:
        """Находит разрешенное ПО с учетом нечеткого сопоставления"""
        name_lower = name.lower() if name else ""
        manufacturer_lower = manufacturer.lower() if manufacturer else None
        
        best_match = None
        best_score = 0.0
        
        for allowed in allowed_list:
            allowed_name = allowed.get('name', '').lower()
            allowed_manufacturer = allowed.get('manufacturer', '').lower() if allowed.get('manufacturer') else None
            
            # Проверяем совпадение имени
            name_score = SequenceMatcher(None, name_lower, allowed_name).ratio()
            
            # Проверяем совпадение производителя
            manufacturer_match = True
            if manufacturer_lower and allowed_manufacturer:
                manufacturer_match = self._fuzzy_match(manufacturer_lower, allowed_manufacturer, threshold=0.7)
            elif manufacturer_lower or allowed_manufacturer:
                # Если один указан, а другой нет - не считаем это совпадением
                manufacturer_match = False
            
            # Общий балл (имя важнее производителя)
            if manufacturer_match:
                score = name_score * 0.7 + (1.0 if manufacturer_match else 0.0) * 0.3
            else:
                score = name_score * 0.5  # Штраф за несовпадение производителя
            
            if score > best_score and score >= 0.8:  # Порог для нечеткого сопоставления
                best_score = score
                best_match = allowed
        
        return best_match
    
    def check_software(self, workstation_id: int, software_list: List[Dict]) -> List[Dict]:
        """Проверяет список ПО на соответствие и возвращает список нарушений"""
        violations = []
        allowed_software = self.allowed_repo.get_all()
        
        for software in software_list:
            name = software.get('name', '')
            manufacturer = software.get('manufacturer')
            version = software.get('version')
            
            if not name:
                continue
            
            # Ищем точное совпадение
            matched_allowed = None
            name_lower = name.lower()
            manufacturer_lower = manufacturer.lower() if manufacturer else None
            
            # Сначала проверяем точное совпадение
            for allowed in allowed_software:
                allowed_name = allowed.get('name', '').lower()
                allowed_manufacturer = allowed.get('manufacturer', '').lower() if allowed.get('manufacturer') else None
                
                # Точное совпадение имени и производителя
                if name_lower == allowed_name:
                    if (manufacturer_lower == allowed_manufacturer) or \
                       (not manufacturer_lower and not allowed_manufacturer) or \
                       (not manufacturer_lower and allowed_manufacturer):
                        matched_allowed = allowed
                        break
            
            # Если точного совпадения нет, пробуем нечеткое сопоставление
            if not matched_allowed:
                matched_allowed = self._find_matching_allowed_software(name, manufacturer, version, allowed_software)
            
            if matched_allowed:
                # Проверяем версию, если указана
                version_pattern = matched_allowed.get('version_pattern')
                if version_pattern and version:
                    if not self._check_version_pattern(version, version_pattern):
                        # Несоответствие версии
                        violation = {
                            'workstation_id': workstation_id,
                            'software': software,
                            'violation_type': 'version_mismatch',
                            'message': f"Несоответствие версии: {name} {version} (разрешено: {version_pattern})"
                        }
                        violations.append(violation)
                # Если версия соответствует или не проверяется - ПО разрешено
            else:
                # ПО не найдено в перечне - создаем нарушение
                violation = {
                    'workstation_id': workstation_id,
                    'software': software,
                    'violation_type': 'unauthorized_software',
                    'message': f"Неразрешенное ПО: {name}"
                }
                violations.append(violation)
        
        return violations
    
    def create_violations(self, violations: List[Dict], software_id_map: Dict[str, int]):
        """Создает записи о нарушениях в базе данных"""
        for violation in violations:
            software = violation['software']
            software_key = f"{software.get('name')}_{software.get('version')}"
            software_id = software_id_map.get(software_key)
            
            if software_id:
                self.compliance_repo.create(
                    violation['workstation_id'],
                    software_id,
                    violation['violation_type'],
                    violation.get('message')
                )

