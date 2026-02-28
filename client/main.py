"""
Главный модуль клиентского приложения
"""
import sys
import os
from pathlib import Path

# Для работы в собранном виде (PyInstaller)
if getattr(sys, 'frozen', False):
    # Если приложение собрано в исполняемый файл
    # sys._MEIPASS - это временная папка, куда PyInstaller распаковывает файлы
    if hasattr(sys, '_MEIPASS'):
        application_path = sys._MEIPASS
        sys.path.insert(0, application_path)
    
    # Также добавляем директорию, где находится исполняемый файл
    exe_dir = Path(sys.executable).parent
    sys.path.insert(0, str(exe_dir))
else:
    # Режим разработки - добавляем корневую директорию проекта в путь
    project_root = Path(__file__).parent.parent
    sys.path.insert(0, str(project_root))

# Импорт главной функции
# В собранном виде PyInstaller должен включить все модули в sys._MEIPASS
# и они должны быть доступны через обычные импорты
try:
    # Основной вариант импорта
    from client.gui import main as gui_main
except ImportError:
    try:
        # Если client.gui не работает, пробуем через sys.path
        # Добавляем client директорию в путь
        client_dir = Path(__file__).parent
        if client_dir not in [Path(p) for p in sys.path]:
            sys.path.insert(0, str(client_dir))
        
        # Пробуем импортировать gui напрямую
        import gui
        gui_main = gui.main
    except ImportError:
        # Последняя попытка - импорт из текущей директории
        try:
            # Добавляем родительскую директорию
            parent_dir = Path(__file__).parent.parent
            if str(parent_dir) not in sys.path:
                sys.path.insert(0, str(parent_dir))
            
            from client.gui import main as gui_main
        except ImportError as e:
            # Если ничего не помогло, показываем ошибку
            import tkinter.messagebox as mb
            error_msg = (
                f"Критическая ошибка: не удалось загрузить модуль GUI.\n\n"
                f"Ошибка: {e}\n\n"
                f"sys.path: {sys.path}\n"
                f"frozen: {getattr(sys, 'frozen', False)}\n"
                f"_MEIPASS: {getattr(sys, '_MEIPASS', 'N/A')}"
            )
            try:
                mb.showerror("Ошибка запуска", error_msg)
            except:
                print(error_msg)
            sys.exit(1)

if __name__ == "__main__":
    gui_main()
