import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import os
import sys
import subprocess
import time
from datetime import datetime
from pathlib import Path

class EventProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🚀 Обработчик мероприятий v1.0")
        self.root.geometry("850x600")
        self.root.minsize(750, 500)  # Минимальный размер
        self.root.resizable(True, True)
        
        # Центрируем окно
        self.center_window()
        
        # Настройка стилей
        self.setup_styles()
        
        # Настройки (СНАЧАЛА переменные!)
        self.base_dir = Path.cwd()
        self.scripts = {
            '1': {
                'name': '1.docxtocsv.ipynb',
                'description': '📄 DOCX → Расписание CSV',
                'details': 'Извлекает таблицу из DOCX и создает расписание с занятиями'
            },
            '2': {
                'name': '3.kod_tipovogo.ipynb', 
                'description': '🔤 Добавление кодов мероприятий',
                'details': 'Сопоставляет мероприятия с кодами из справочника'
            },
            '3': {
                'name': '4.dopobrabokta.ipynb',
                'description': '📊 Финальная обработка Excel',
                'details': 'Создает отформатированные Excel файлы по мероприятиям'
            }
        }
        
        # Флаг обработки
        self.processing = False
        
        # Создаем интерфейс (ПОСЛЕ определения переменных!)
        self.create_widgets()
        
        # Проверяем файлы при запуске (автоматически)
        self.root.after(100, self.check_all_files)  # Через 100мс после запуска

    def center_window(self):
        """Центрирует окно на экране"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def setup_styles(self):
        """Настраивает стили интерфейса"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Кастомные стили
        style.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'), foreground='#2c3e50')
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'), foreground='#34495e')
        style.configure('Success.TLabel', foreground='#27ae60', font=('Segoe UI', 10))
        style.configure('Error.TLabel', foreground='#e74c3c', font=('Segoe UI', 10))
        style.configure('Warning.TLabel', foreground='#f39c12', font=('Segoe UI', 10))
        style.configure('Big.TButton', font=('Segoe UI', 11, 'bold'), padding=10)

    def create_widgets(self):
        """Создает виджеты интерфейса"""
        # Главный контейнер
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Заголовок с иконкой
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        title_label = ttk.Label(title_frame, text="🚀 Обработчик мероприятий", style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame, text="Автоматическая обработка данных о курсах", 
                                 font=('Segoe UI', 10), foreground='#7f8c8d')
        subtitle_label.pack(pady=(5, 0))
        
        # Фрейм для проверки файлов
        files_frame = ttk.LabelFrame(main_frame, text="📁 Статус системы", padding="10")
        files_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        files_frame.columnconfigure(1, weight=1)
        
        self.file_status_frame = ttk.Frame(files_frame)
        self.file_status_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        # Основные кнопки действий
        actions_frame = ttk.LabelFrame(main_frame, text="🚀 Действия", padding="10")
        actions_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Кнопка полной обработки
        self.full_process_btn = ttk.Button(actions_frame, text="🎯 Полная обработка (все этапы)", 
                                          command=self.start_full_process, style='Big.TButton')
        self.full_process_btn.pack(fill=tk.X, pady=(0, 10))
        
        # Кнопки отдельных этапов
        stages_frame = ttk.Frame(actions_frame)
        stages_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.stage_buttons = {}
        for i, (num, info) in enumerate(self.scripts.items()):
            btn = ttk.Button(stages_frame, text=f"Этап {num}: {info['description']}", 
                           command=lambda n=num: self.start_single_stage(n))
            btn.pack(fill=tk.X, pady=2)
            self.stage_buttons[num] = btn
        
        # Дополнительные действия
        extra_frame = ttk.Frame(actions_frame)
        extra_frame.pack(fill=tk.X, pady=(15, 0))
        
        ttk.Button(extra_frame, text="📁 Открыть результаты", 
                  command=self.open_results).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(extra_frame, text="🗑️ Очистить временные", 
                  command=self.cleanup_files).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(extra_frame, text="❓ Справка", 
                  command=self.show_help).pack(side=tk.LEFT)
        
        # Прогресс
        progress_frame = ttk.LabelFrame(main_frame, text="📊 Прогресс", padding="8")
        progress_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(progress_frame, text="Готов к работе", font=('Segoe UI', 10))
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(8, 0))
        
        # Лог
        log_frame = ttk.LabelFrame(main_frame, text="📋 Лог обработки", padding="8")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=80, 
                                                 font=('Consolas', 9), wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Кнопки для лога
        log_buttons = ttk.Frame(log_frame)
        log_buttons.grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        
        ttk.Button(log_buttons, text="💾 Сохранить лог", 
                  command=self.save_log).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(log_buttons, text="🗑️ Очистить лог", 
                  command=self.clear_log).pack(side=tk.LEFT)

    def log_message(self, message, level="INFO"):
        """Добавляет сообщение в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        level_icon = {"INFO": "ℹ️", "SUCCESS": "✅", "ERROR": "❌", "WARNING": "⚠️"}.get(level, "📝")
        formatted_message = f"[{timestamp}] {level_icon} {message}\n"
        
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def check_script_exists(self, script_name):
        """Проверяет существование скрипта"""
        return (self.base_dir / script_name).exists()

    def check_all_files(self):
        """Проверяет все необходимые файлы"""
        # Очищаем предыдущие статусы
        for widget in self.file_status_frame.winfo_children():
            widget.destroy()
        
        row = 0
        all_good = True
        
        # Проверяем notebook файлы
        ttk.Label(self.file_status_frame, text="📓 Notebook файлы:", style='Header.TLabel').grid(
            row=row, column=0, sticky=tk.W, pady=(0, 5))
        row += 1
        
        for script_info in self.scripts.values():
            if self.check_script_exists(script_info['name']):
                ttk.Label(self.file_status_frame, text=f"✅ {script_info['name']}", 
                         style='Success.TLabel').grid(row=row, column=0, sticky=tk.W, padx=(20, 0))
            else:
                ttk.Label(self.file_status_frame, text=f"❌ {script_info['name']}", 
                         style='Error.TLabel').grid(row=row, column=0, sticky=tk.W, padx=(20, 0))
                all_good = False
            row += 1
        
        # Проверяем входные файлы
        ttk.Label(self.file_status_frame, text="📄 Входные файлы:", style='Header.TLabel').grid(
            row=row, column=0, sticky=tk.W, pady=(10, 5))
        row += 1
        
        input_files = ['test.docx', 'kod_tipovogo.xlsx']
        for file in input_files:
            if (self.base_dir / file).exists():
                ttk.Label(self.file_status_frame, text=f"✅ {file}", 
                         style='Success.TLabel').grid(row=row, column=0, sticky=tk.W, padx=(20, 0))
            else:
                ttk.Label(self.file_status_frame, text=f"⚠️ {file} не найден", 
                         style='Warning.TLabel').grid(row=row, column=0, sticky=tk.W, padx=(20, 0))
            row += 1
        
        
        
        # Общий статус
        if all_good:
            self.log_message("Система готова к работе! Все файлы найдены.", "SUCCESS")
        else:
            self.log_message("Внимание: некоторые файлы отсутствуют", "WARNING")

    def set_processing_state(self, processing):
        """Устанавливает состояние обработки"""
        self.processing = processing
        
        if processing:
            self.full_process_btn.configure(state='disabled', text="⏳ Обрабатывается...")
            for btn in self.stage_buttons.values():
                btn.configure(state='disabled')
            self.progress_bar.start()
        else:
            self.full_process_btn.configure(state='normal', text="🎯 Полная обработка (все этапы)")
            for btn in self.stage_buttons.values():
                btn.configure(state='normal')
            self.progress_bar.stop()

    def update_status(self, message):
        """Обновляет статус"""
        self.status_label.configure(text=message)
        self.log_message(message)

    def run_notebook(self, notebook_name, description):
        """Запускает notebook"""
        self.update_status(f"Запуск: {description}")
        
        start_time = time.time()
        
        try:
            # Запускаем notebook через jupyter nbconvert
            result = subprocess.run([
                'jupyter', 'nbconvert', '--execute', '--to', 'notebook', 
                '--inplace', str(self.base_dir / notebook_name)
            ], capture_output=True, text=True, cwd=self.base_dir, timeout=600)
            
            elapsed_time = time.time() - start_time
            
            if result.returncode == 0:
                self.log_message(f"✅ {description} завершено успешно за {elapsed_time:.1f} сек", "SUCCESS")
                return True
            else:
                self.log_message(f"❌ Ошибка в {description}: {result.stderr}", "ERROR")
                return False
                
        except subprocess.TimeoutExpired:
            self.log_message(f"❌ Таймаут при выполнении {description}", "ERROR")
            return False
        except FileNotFoundError:
            self.log_message("❌ Jupyter не найден! Установите: pip install jupyter", "ERROR")
            return False
        except Exception as e:
            self.log_message(f"❌ Ошибка: {e}", "ERROR")
            return False

    def start_full_process(self):
        """Запускает полную обработку"""
        if self.processing:
            return
        
        def process_thread():
            try:
                self.set_processing_state(True)
                self.update_status("🚀 Начинаем полную обработку...")
                
                successful = 0
                total = len(self.scripts)
                
                for num, script_info in self.scripts.items():
                    self.update_status(f"📋 Этап {num}/{total}: {script_info['description']}")
                    
                    if self.run_notebook(script_info['name'], script_info['description']):
                        successful += 1
                    else:
                        break  # Останавливаемся при ошибке
                    
                    time.sleep(1)  # Пауза между этапами
                
                if successful == total:
                    self.update_status("🎉 Вся обработка завершена успешно!")
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Успех!", 
                        "Обработка завершена успешно!\n\nРезультаты находятся в папке '4.excel_final'"
                    ))
                else:
                    self.update_status(f"⚠️ Обработка завершена с ошибками ({successful}/{total})")
                
            except Exception as e:
                self.log_message(f"💥 Критическая ошибка: {e}", "ERROR")
                self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}"))
            
            finally:
                self.root.after(0, lambda: self.set_processing_state(False))
                self.root.after(0, lambda: self.update_status("Готов к работе"))
        
        threading.Thread(target=process_thread, daemon=True).start()

    def start_single_stage(self, stage_num):
        """Запускает отдельный этап"""
        if self.processing:
            return
        
        script_info = self.scripts[stage_num]
        
        def stage_thread():
            try:
                self.set_processing_state(True)
                
                if self.run_notebook(script_info['name'], script_info['description']):
                    self.root.after(0, lambda: messagebox.showinfo("Успех", f"Этап {stage_num} завершен успешно!"))
                else:
                    self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Ошибка в этапе {stage_num}"))
                
            except Exception as e:
                self.log_message(f"💥 Ошибка этапа {stage_num}: {e}", "ERROR")
                self.root.after(0, lambda: messagebox.showerror("Ошибка", f"Ошибка: {e}"))
            
            finally:
                self.root.after(0, lambda: self.set_processing_state(False))
                self.root.after(0, lambda: self.update_status("Готов к работе"))
        
        threading.Thread(target=stage_thread, daemon=True).start()

    def open_results(self):
        """Открывает папку с результатами"""
        results_dir = self.base_dir / "4.excel_final"
        if results_dir.exists():
            if sys.platform.startswith('win'):
                os.startfile(results_dir)
            elif sys.platform.startswith('darwin'):
                subprocess.run(['open', str(results_dir)])
            else:
                subprocess.run(['xdg-open', str(results_dir)])
        else:
            messagebox.showwarning("Предупреждение", "Папка с результатами не найдена")

    def cleanup_files(self):
        """Очищает временные файлы"""
        temp_items = ["split_events", "excel_events", "1.courses_data.csv", 
                     "1.reordered_file.csv", "3.new_code.csv"]
        
        cleaned = 0
        for item in temp_items:
            item_path = self.base_dir / item
            try:
                if item_path.is_dir():
                    import shutil
                    shutil.rmtree(item_path)
                    cleaned += 1
                elif item_path.is_file():
                    item_path.unlink()
                    cleaned += 1
            except Exception:
                pass
        
        if cleaned > 0:
            self.log_message(f"🗑️ Очищено {cleaned} временных элементов", "SUCCESS")
            messagebox.showinfo("Успех", f"Очищено {cleaned} временных файлов")
        else:
            messagebox.showinfo("Информация", "Временные файлы не найдены")

    def save_log(self):
        """Сохраняет лог в файл"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
            title="Сохранить лог"
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get(1.0, tk.END))
                messagebox.showinfo("Успех", "Лог сохранен")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка сохранения: {e}")

    def clear_log(self):
        """Очищает лог"""
        self.log_text.delete(1.0, tk.END)

    def show_help(self):
        """Показывает справку"""
        help_text = """
🚀 СПРАВКА ПО ИСПОЛЬЗОВАНИЮ

📋 ЭТАПЫ ОБРАБОТКИ:

1️⃣ DOCX → Расписание CSV
• Извлекает таблицу из test.docx
• Создает расписание с занятиями и группами
• Результат: 1.reordered_file.csv

2️⃣ Добавление кодов мероприятий
• Сопоставляет названия с кодами из kod_tipovogo.xlsx
• Использует нечеткое сопоставление
• Результат: 3.new_code.csv

3️⃣ Финальная обработка Excel
• Разделяет мероприятия на отдельные файлы
• Форматирует даты и заполняет категории
• Результат: папка 4.excel_final/

📁 ТРЕБУЕМЫЕ ФАЙЛЫ:
• test.docx - документ с таблицей курсов
• kod_tipovogo.xlsx - справочник кодов мероприятий
• Все 3 .ipynb файла в той же папке

🎯 РЕЗУЛЬТАТ:
Папка 4.excel_final/ с отформатированными Excel файлами,
где каждое мероприятие в отдельном файле.
        """
        
        messagebox.showinfo("Справка", help_text)


def main():
    """Главная функция"""
    root = tk.Tk()
    app = EventProcessorGUI(root)
    
    try:
        # Начальное сообщение
        app.log_message("Добро пожаловать в обработчик мероприятий!", "SUCCESS")
        
        root.mainloop()
    except KeyboardInterrupt:
        print("Программа прервана пользователем")
    except Exception as e:
        print(f"Ошибка: {e}")


if __name__ == "__main__":
    main()