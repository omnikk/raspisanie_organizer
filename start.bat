@echo off
title Обработчик мероприятий

echo 🚀 Запуск GUI интерфейса...

REM Быстрая проверка
if not exist "gui_events.py" (
    echo ❌ gui_events.py не найден!
    pause
    exit /b 1
)

REM Запуск
python gui_events.py

REM Если закрылся с ошибкой, покажем паузу
if errorlevel 1 pause