@echo off
chcp 65001 >nul
title Установка зависимостей - Обработчик мероприятий

echo.
echo                ╔══════════════════════════════════════════════════════════════╗
echo                ║                УСТАНОВКА ЗАВИСИМОСТЕЙ                        ║
echo                ║                Обработчик мероприятий                        ║
echo                ╚══════════════════════════════════════════════════════════════╝
echo.

echo  Проверка Python...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo  Python не установлен!
    echo.
    echo Скачайте и установите Python с официального сайта:
    echo https://www.python.org/downloads/
    echo.
    echo Обязательно отметьте "Add Python to PATH" при установке!
    pause
    exit /b 1
)

python --version
echo  Python найден

echo.
echo  Проверка pip...
pip --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo  pip не найден!
    echo Переустановите Python с включенным pip
    pause
    exit /b 1
)

echo  pip найден

echo.
echo  Обновление pip до последней версии...
python -m pip install --upgrade pip
if %ERRORLEVEL% neq 0 (
    echo   Не удалось обновить pip, продолжаем с текущей версией
)

echo.
echo  Установка зависимостей из requirements.txt...
echo ═══════════════════════════════════════════════════════════════

if not exist "requirements.txt" (
    echo Файл requirements.txt не найден!
)

pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo.
    echo  Ошибка при установке зависимостей!
    echo.
    echo Попробуйте установить вручную:
    echo pip install pandas openpyxl python-docx fuzzywuzzy python-Levenshtein
    echo.
    pause
    exit /b 1
)

echo.
echo  Все зависимости установлены успешно!

echo.
echo  Проверка установленных библиотек...
python -c "import pandas; print(' pandas:', pandas.__version__)" 2>nul || echo " pandas не работает"
python -c "import openpyxl; print(' openpyxl:', openpyxl.__version__)" 2>nul || echo " openpyxl не работает"
python -c "import docx; print(' python-docx: OK')" 2>nul || echo " python-docx не работает"
python -c "import fuzzywuzzy; print(' fuzzywuzzy: OK')" 2>nul || echo " fuzzywuzzy не работает"

echo.

echo.
echo                ╔══════════════════════════════════════════════════════════════╗
echo                ║                     УСТАНОВКА ЗАВЕРШЕНА!                     ║
echo                ╚══════════════════════════════════════════════════════════════╝


pause