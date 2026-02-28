@echo off
chcp 65001 >nul
echo ============================================================
echo Регистрация агента как службы Windows
echo ============================================================
echo.

REM Проверяем права администратора
net session >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Требуются права администратора!
    echo Пожалуйста, запустите этот файл от имени администратора:
    echo 1. Щелкните правой кнопкой мыши по файлу
    echo 2. Выберите "Запуск от имени администратора"
    pause
    exit /b 1
)

cd /d "%~dp0\.."

REM Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не найден!
    echo Пожалуйста, установите Python версии 3.8 или выше.
    pause
    exit /b 1
)

echo Создание задачи для автозапуска агента...
echo.

REM Создаем задачу планировщика для автозапуска агента
set TASK_NAME=Programmnyy_Strazh_Agent
set PYTHON_PATH=python
set SCRIPT_PATH=%~dp0main.py

REM Удаляем существующую задачу, если есть
schtasks /delete /tn "%TASK_NAME%" /f >nul 2>&1

REM Создаем новую задачу
schtasks /create /tn "%TASK_NAME%" /tr "\"%PYTHON_PATH%\" \"%SCRIPT_PATH%\"" /sc onlogon /ru SYSTEM /rl HIGHEST /f

if errorlevel 1 (
    echo ОШИБКА: Не удалось создать задачу планировщика!
    echo.
    echo Альтернативный способ: добавьте ярлык агента в автозагрузку Windows:
    echo 1. Нажмите Win+R
    echo 2. Введите: shell:startup
    echo 3. Скопируйте ярлык start_agent.bat в открывшуюся папку
    pause
    exit /b 1
)

echo.
echo ============================================================
echo Агент успешно зарегистрирован как служба!
echo ============================================================
echo.
echo Задача планировщика "%TASK_NAME%" создана.
echo Агент будет автоматически запускаться при входе пользователя в систему.
echo.
echo Для проверки статуса задачи используйте:
echo   schtasks /query /tn "%TASK_NAME%"
echo.
echo Для удаления задачи используйте:
echo   schtasks /delete /tn "%TASK_NAME%" /f
echo.
pause
