@echo off
chcp 65001 >nul
title Установка зависимостей - Обработчик мероприятий

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║                УСТАНОВКА ЗАВИСИМОСТЕЙ                       ║
echo ║              Обработчик мероприятий                         ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

echo 🔍 Проверка Python...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ❌ Python не установлен!
    echo.
    echo Скачайте и установите Python с официального сайта:
    echo https://www.python.org/downloads/
    echo.
    echo Обязательно отметьте "Add Python to PATH" при установке!
    pause
    exit /b 1
)

python --version
echo ✅ Python найден

echo.
echo 📦 Проверка pip...
pip --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo ❌ pip не найден!
    echo Переустановите Python с включенным pip
    pause
    exit /b 1
)

echo ✅ pip найден

echo.
echo 🔄 Обновление pip до последней версии...
python -m pip install --upgrade pip
if %ERRORLEVEL% neq 0 (
    echo ⚠️  Не удалось обновить pip, продолжаем с текущей версией
)

echo.
echo 📋 Установка зависимостей из requirements.txt...
echo ═══════════════════════════════════════════════════════════════

if not exist "requirements.txt" (
    echo ❌ Файл requirements.txt не найден!
    echo Создаём файл requirements.txt...
    
    echo # Обработчик мероприятий - Зависимости > requirements.txt
    echo pandas^>=1.5.0 >> requirements.txt
    echo openpyxl^>=3.0.0 >> requirements.txt
    echo python-docx^>=0.8.11 >> requirements.txt
    echo fuzzywuzzy^>=0.18.0 >> requirements.txt
    echo python-Levenshtein^>=0.12.0 >> requirements.txt
    
    echo ✅ Файл requirements.txt создан
)

pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ Ошибка при установке зависимостей!
    echo.
    echo Попробуйте установить вручную:
    echo pip install pandas openpyxl python-docx fuzzywuzzy python-Levenshtein
    echo.
    pause
    exit /b 1
)

echo.
echo ✅ Все зависимости установлены успешно!

echo.
echo 🧪 Проверка установленных библиотек...
python -c "import pandas; print('✅ pandas:', pandas.__version__)" 2>nul || echo "❌ pandas не работает"
python -c "import openpyxl; print('✅ openpyxl:', openpyxl.__version__)" 2>nul || echo "❌ openpyxl не работает"
python -c "import docx; print('✅ python-docx: OK')" 2>nul || echo "❌ python-docx не работает"
python -c "import fuzzywuzzy; print('✅ fuzzywuzzy: OK')" 2>nul || echo "❌ fuzzywuzzy не работает"

echo.
echo 📁 Проверка основных файлов...
if exist "unified_processor.py" (
    echo ✅ unified_processor.py найден
) else (
    echo ❌ unified_processor.py НЕ найден!
)

if exist "gui_processor.py" (
    echo ✅ gui_processor.py найден
) else (
    echo ❌ gui_processor.py НЕ найден!
)

if exist "test.docx" (
    echo ✅ test.docx найден
) else (
    echo ⚠️  test.docx НЕ найден - поместите ваш DOCX файл и переименуйте в test.docx
)

if exist "kod_tipovogo.xlsx" (
    echo ✅ kod_tipovogo.xlsx найден
) else (
    echo ⚠️  kod_tipovogo.xlsx НЕ найден - поместите файл с кодами мероприятий
)

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║                     УСТАНОВКА ЗАВЕРШЕНА!                    ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo 🎉 Теперь вы можете запускать обработчик:
echo.
echo • Двойной клик на launcher.bat - главное меню
echo • python gui_processor.py - графический интерфейс  
echo • python unified_processor.py - консольная версия
echo.

pause