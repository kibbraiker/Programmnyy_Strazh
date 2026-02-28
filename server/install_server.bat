@echo off
chcp 65001 >nul
echo ============================================================
echo Установка серверного компонента "Программный Страж"
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
echo Инициализация базы данных...
python database\init_db.py
if errorlevel 1 (
    echo ОШИБКА: Не удалось инициализировать базу данных!
    echo Проверьте права доступа к папке database\
    pause
    exit /b 1
)

echo.
echo ============================================================
echo Установка сервера завершена успешно!
echo ============================================================
echo.
echo Для запуска сервера используйте:
echo   start_server.bat
echo или
echo   python server\main.py
echo.
pause
