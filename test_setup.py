def test_python():
    """Проверяет версию Python"""
    import sys
    print(f" Python версия: {sys.version}")
    
    if sys.version_info < (3, 7):
        print(" Требуется Python 3.7 или выше!")
        return False
    else:
        print(" Версия Python подходит")
        return True

def test_libraries():
    """Проверяет все необходимые библиотеки"""
    print("\n ПРОВЕРКА БИБЛИОТЕК")
    print("=" * 40)
    
    libraries = [
        ("pandas", "import pandas as pd"),
        ("python-docx", "import docx"),
        ("fuzzywuzzy", "import fuzzywuzzy"),
        ("openpyxl", "import openpyxl"),
        ("jupyter", "import jupyter"),
        ("tkinter", "import tkinter"),
    ]
    
    success_count = 0
    
    for name, import_code in libraries:
        try:
            exec(import_code)
            print(f"УСПЕХ {name}")
            success_count += 1
        except ImportError as e:
            print(f"ОШИБКА {name}: не установлен")
        except Exception as e:
            print(f"ВНИМАНИЕ  {name}: {e}")
    
    print(f"\n Результат: {success_count}/{len(libraries)} библиотек готово")
    
    if success_count == len(libraries):
        print(" ВСЕ БИБЛИОТЕКИ УСТАНОВЛЕНЫ!")
        return True
    else:
        print(" Не все библиотеки установлены")
        print("Установите: pip install -r requirements.txt")
        return False

def test_files():
    """Проверяет наличие файлов"""
    print("\n ПРОВЕРКА ФАЙЛОВ")
    print("=" * 40)
    
    import os
    
    required_files = [
        ("gui_events.py", "Главный GUI интерфейс"),
        ("1.docxtocsv.ipynb", "Этап 1: DOCX → CSV"),
        ("3.kod_tipovogo.ipynb", "Этап 2: Коды мероприятий"),
        ("4.dopobrabokta.ipynb", "Этап 3: Excel обработка"),
    ]
    
    input_files = [
        ("test.docx", "DOCX с таблицей курсов"),
        ("kod_tipovogo.xlsx", "Excel с кодами мероприятий"),
    ]
    
    all_found = True
    
    print(" Системные файлы:")
    for filename, description in required_files:
        if os.path.exists(filename):
            print(f" {filename}")
        else:
            print(f" {filename} - {description}")
            all_found = False
    
    print("\n Входные файлы:")
    for filename, description in input_files:
        if os.path.exists(filename):
            print(f"УСПЕХ {filename}")
        else:
            print(f"ВНИМАНИЕ  {filename} - {description}")
    
    return all_found

def test_jupyter():
    """Проверяет Jupyter"""
    print("\n ПРОВЕРКА JUPYTER")
    print("=" * 40)
    
    try:
        import subprocess
        result = subprocess.run(['jupyter', '--version'], 
                              capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            print(f" Jupyter установлен")
            print(f" Версия: {result.stdout.strip()}")
            return True
        else:
            print(" Jupyter не работает")
            return False
            
    except FileNotFoundError:
        print(" Jupyter не найден")
        print(" Установите: pip install jupyter")
        return False
    except Exception as e:
        print(f" Ошибка проверки Jupyter: {e}")
        return False

def main():
    """Запуск всех проверок"""
    print(" ПРОВЕРКА ГОТОВНОСТИ СИСТЕМЫ")
    print("Для обработки мероприятий")
    print("=" * 50)
    
    all_good = True
    
    # Проверка Python
    if not test_python():
        all_good = False
    
    # Проверка библиотек
    if not test_libraries():
        all_good = False
    
    # Проверка файлов
    if not test_files():
        all_good = False
    
    # Проверка Jupyter
    if not test_jupyter():
        all_good = False
    
    print("\n" + "=" * 50)
    if all_good:
        print(" СИСТЕМА ГОТОВА К РАБОТЕ!")
        print(" Можете запускать run_events.bat")
    else:
        print(" СИСТЕМА НЕ ГОТОВА")
        print(" Исправьте ошибки выше и запустите тест снова")
    
    print("\n" + "=" * 50)
    print("Для выхода нажмите Enter...")
    input()
if __name__ == "__main__":
    main()