"""
Модуль для сканирования установленного программного обеспечения
"""
import winreg
import os
from pathlib import Path
from typing import List, Dict

class SoftwareScanner:
    """Класс для сканирования установленного ПО через реестр Windows"""
    
    REGISTRY_PATHS = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]
    
    def scan_installed_software(self) -> List[Dict]:
        """Сканирует установленное программное обеспечение"""
        software_list = []
        
        for hkey, path in self.REGISTRY_PATHS:
            try:
                key = winreg.OpenKey(hkey, path)
                self._scan_registry_key(key, software_list)
                winreg.CloseKey(key)
            except Exception as e:
                print(f"Ошибка при сканировании {path}: {e}")
        
        # Удаляем дубликаты
        seen = set()
        unique_software = []
        for item in software_list:
            key = (item.get('name'), item.get('version'), item.get('manufacturer'))
            if key not in seen:
                seen.add(key)
                unique_software.append(item)
        
        return unique_software
    
    def _scan_registry_key(self, key, software_list: List[Dict]):
        """Сканирует раздел реестра"""
        i = 0
        while True:
            try:
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                
                try:
                    name = self._get_registry_value(subkey, "DisplayName")
                    if name:
                        software = {
                            'name': name,
                            'version': self._get_registry_value(subkey, "DisplayVersion"),
                            'manufacturer': self._get_registry_value(subkey, "Publisher"),
                            'install_date': self._get_registry_value(subkey, "InstallDate"),
                            'executable_path': self._get_registry_value(subkey, "InstallLocation"),
                        }
                        
                        # Вычисляем размер, если возможно
                        install_location = software.get('executable_path')
                        if install_location and os.path.exists(install_location):
                            try:
                                software['size'] = self._get_directory_size(install_location)
                            except:
                                software['size'] = None
                        else:
                            software['size'] = None
                        
                        software_list.append(software)
                finally:
                    winreg.CloseKey(subkey)
                
                i += 1
            except OSError:
                break
    
    def _get_registry_value(self, key, value_name: str) -> str:
        """Получает значение из реестра"""
        try:
            value, _ = winreg.QueryValueEx(key, value_name)
            return str(value) if value else None
        except:
            return None
    
    def _get_directory_size(self, path: str) -> int:
        """Вычисляет размер директории в байтах"""
        total_size = 0
        try:
            for dirpath, dirnames, filenames in os.walk(path):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    try:
                        total_size += os.path.getsize(filepath)
                    except:
                        pass
        except:
            pass
        return total_size

