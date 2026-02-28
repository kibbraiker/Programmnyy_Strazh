@echo off
chcp 65001 >nul
echo ============================================================
echo Установка клиентского приложения "Программный Страж"
echo ============================================================
echo.

cd /d "%~dp0\.."

REM Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не найден!
    echo Пожалуйста, установите Python версии 3.8 или выше.
    echo Скачайте с официального сайта: https://www.python.org/downloads/
    echo При установке обязательно установите флажок "Add Python to PATH"
    pause
    exit /b 1
)

echo Установка зависимостей Python...
pip install -r requirements.txt
if errorlevel 1 (
    echo ОШИБКА: Не удалось установить зависимости!
    echo Проверьте подключение к интернету и повторите попытку.
    pause
    exit /b 1
)

echo.
echo Настройка конфигурации...
echo.
echo ВАЖНО: Убедитесь, что в файле config\client_config.json указан
echo правильный IP-адрес и порт сервера!
echo.
echo Текущие настройки:
type config\client_config.json
echo.

echo ============================================================
echo Установка клиента завершена успешно!
echo ============================================================
echo.
echo Для запуска клиента используйте:
echo   start_client.bat
echo или
echo   python client\main.py
echo.
pause
