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
        self.root.geometry("700x480")  # Уменьшили с 850x600
        self.root.minsize(650, 420)    # Уменьшили с 750x500
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
                'description': '📄 DOCX → CSV',
                'details': 'Извлекает таблицу из DOCX и создает расписание с занятиями'
            },
            '2': {
                'name': '3.kod_tipovogo.ipynb', 
                'description': '🔤 Коды мероприятий',
                'details': 'Сопоставляет мероприятия с кодами из справочника'
            },
            '3': {
                'name': '4.dopobrabokta.ipynb',
                'description': '📊 Финальная обработка',
                'details': 'Создает отформатированные Excel файлы по мероприятиям'
            }
        }
        
        # Флаг обработки
        self.processing = False
        
        # Создаем интерфейс (ПОСЛЕ определения переменных!)
        self.create_widgets()
        
        # Проверяем файлы при запуске (автоматически)
        self.root.after(100, self.check_all_files)

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
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground='#2c3e50')
        style.configure('Header.TLabel', font=('Segoe UI', 10, 'bold'), foreground='#34495e')
        style.configure('Success.TLabel', foreground='#27ae60', font=('Segoe UI', 9))
        style.configure('Error.TLabel', foreground='#e74c3c', font=('Segoe UI', 9))
        style.configure('Warning.TLabel', foreground='#f39c12', font=('Segoe UI', 9))
        style.configure('Big.TButton', font=('Segoe UI', 10, 'bold'), padding=(8, 6))
        style.configure('Small.TButton', font=('Segoe UI', 9), padding=(4, 2))

    def create_widgets(self):
        """Создает виджеты интерфейса"""
        # Главный контейнер с меньшими отступами
        main_frame = ttk.Frame(self.root, padding="8")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Компактный заголовок
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        
        title_label = ttk.Label(title_frame, text="🚀 Обработчик мероприятий", style='Title.TLabel')
        title_label.pack()
        
        # Детальный статус системы (компактно)
        status_frame = ttk.LabelFrame(main_frame, text="📁 Статус системы", padding="6")
        status_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        status_frame.columnconfigure(0, weight=1)
        
        # Статус notebook файлов
        self.nb_status_label = ttk.Label(status_frame, text="📓 Notebook файлы: Проверка...", 
                                        font=('Segoe UI', 9), style='Header.TLabel')
        self.nb_status_label.grid(row=0, column=0, sticky=tk.W)
        
        self.nb_files_label = ttk.Label(status_frame, text="", 
                                       font=('Segoe UI', 8))
        self.nb_files_label.grid(row=1, column=0, sticky=tk.W, padx=(15, 0))
        
        # Статус входных файлов  
        self.input_status_label = ttk.Label(status_frame, text="📄 Входные файлы:", 
                                           font=('Segoe UI', 9), style='Header.TLabel')
        self.input_status_label.grid(row=2, column=0, sticky=tk.W, pady=(4, 0))
        
        self.input_files_label = ttk.Label(status_frame, text="", 
                                          font=('Segoe UI', 8))
        self.input_files_label.grid(row=3, column=0, sticky=tk.W, padx=(15, 0))
        
        # Основные действия - более компактно
        actions_frame = ttk.LabelFrame(main_frame, text="🚀 Действия", padding="6")
        actions_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        actions_frame.columnconfigure(0, weight=1)
        
        # Главная кнопка
        self.full_process_btn = ttk.Button(actions_frame, text="🎯 Полная обработка", 
                                          command=self.start_full_process, style='Big.TButton')
        self.full_process_btn.pack(fill=tk.X, pady=(0, 6))
        
        # Кнопки этапов в горизонтальном ряду
        stages_frame = ttk.Frame(actions_frame)
        stages_frame.pack(fill=tk.X, pady=(2, 0))
        
        self.stage_buttons = {}
        for i, (num, info) in enumerate(self.scripts.items()):
            btn = ttk.Button(stages_frame, text=info['description'], 
                           command=lambda n=num: self.start_single_stage(n),
                           style='Small.TButton')
            btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 2) if i < len(self.scripts)-1 else 0)
            self.stage_buttons[num] = btn
        
        # Дополнительные действия в одну строку
        extra_frame = ttk.Frame(actions_frame)
        extra_frame.pack(fill=tk.X, pady=(6, 0))
        
        ttk.Button(extra_frame, text="📁 Результаты", 
                  command=self.open_results, style='Small.TButton').pack(side=tk.LEFT, padx=(0, 4))
        
        ttk.Button(extra_frame, text="🗑️ Очистить", 
                  command=self.cleanup_files, style='Small.TButton').pack(side=tk.LEFT, padx=(0, 4))
        
        ttk.Button(extra_frame, text="❓ Справка", 
                  command=self.show_help, style='Small.TButton').pack(side=tk.LEFT)
        
        # Компактный прогресс
        progress_frame = ttk.LabelFrame(main_frame, text="📊 Прогресс", padding="6")
        progress_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 8))
        progress_frame.columnconfigure(0, weight=1)
        
        self.status_label = ttk.Label(progress_frame, text="Готов к работе", font=('Segoe UI', 9))
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(4, 0))
        
        # Компактный лог
        log_frame = ttk.LabelFrame(main_frame, text="📋 Лог", padding="6")
        log_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Уменьшили высоту лога с 10 до 6
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, width=70, 
                                                 font=('Consolas', 8), wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Компактные кнопки для лога
        log_buttons = ttk.Frame(log_frame)
        log_buttons.grid(row=1, column=0, sticky=tk.W, pady=(6, 0))
        
        ttk.Button(log_buttons, text="💾 Сохранить", 
                  command=self.save_log, style='Small.TButton').pack(side=tk.LEFT, padx=(0, 4))
        
        ttk.Button(log_buttons, text="🗑️ Очистить", 
                  command=self.clear_log, style='Small.TButton').pack(side=tk.LEFT)

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
        """Проверяет все необходимые файлы с детальным статусом"""
        # Проверяем notebook файлы
        nb_files_status = []
        all_nb_ok = True
        
        for script_info in self.scripts.values():
            if self.check_script_exists(script_info['name']):
                nb_files_status.append(f"✅ {script_info['name']}")
            else:
                nb_files_status.append(f"❌ {script_info['name']}")
                all_nb_ok = False
        
        # Проверяем входные файлы
        input_files = ['test.docx', 'kod_tipovogo.xlsx']
        input_files_status = []
        
        for file in input_files:
            if (self.base_dir / file).exists():
                input_files_status.append(f"✅ {file}")
            else:
                input_files_status.append(f"⚠️ {file} не найден")
        
        # Обновляем лейблы
        nb_status_text = "Notebook файлы:" if all_nb_ok else "Notebook файлы: (есть проблемы)"
        self.nb_status_label.configure(text=f"📓 {nb_status_text}")
        self.nb_files_label.configure(text="  •  ".join(nb_files_status))
        
        self.input_status_label.configure(text="📄 Входные файлы:")
        self.input_files_label.configure(text="  •  ".join(input_files_status))
        
        # Логируем общий статус
        if all_nb_ok:
            self.log_message("Система готова к работе!", "SUCCESS")
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
            self.full_process_btn.configure(state='normal', text="🎯 Полная обработка")
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
                self.log_message(f"✅ {description} завершено за {elapsed_time:.1f} сек", "SUCCESS")
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
                    self.update_status("🎉 Обработка завершена!")
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Успех!", 
                        "Обработка завершена успешно!\n\nРезультаты в папке '4.excel_final'"
                    ))
                else:
                    self.update_status(f"⚠️ Завершено с ошибками ({successful}/{total})")
                
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
                    self.root.after(0, lambda: messagebox.showinfo("Успех", f"Этап {stage_num} завершен!"))
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
            self.log_message(f"🗑️ Очищено {cleaned} элементов", "SUCCESS")
            messagebox.showinfo("Успех", f"Очищено {cleaned} файлов")
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
        help_text = """🚀 СПРАВКА

📋 ЭТАПЫ:
1️⃣ DOCX → CSV: Извлекает таблицу и создает расписание
2️⃣ Коды мероприятий: Сопоставляет с справочником  
3️⃣ Финальная обработка: Создает Excel файлы

📁 ФАЙЛЫ:
• test.docx - таблица курсов
• kod_tipovogo.xlsx - справочник кодов

🎯 РЕЗУЛЬТАТ: Папка 4.excel_final/
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