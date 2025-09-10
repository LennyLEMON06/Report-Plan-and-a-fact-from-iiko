# app_gui.py
import tkinter as tk
from tkinter import ttk, messagebox
import threading
from tkcalendar import DateEntry
import json
from pathlib import Path
import calendar
from datetime import datetime, timedelta
import time
import os
import sys

# Импортируем наши модули
from iiko_collector import IikoDataCollector
from report_generator import ReportGenerator

class IikoReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Получение данных из iiko")
        self.root.geometry("600x800")
        self.root.resizable(True, True)
        
        # Переменные
        self.bases_config = {}
        self.selected_bases = []
        self.selected_sheets = []
        self.start_date = None
        self.end_date = None
        
        # Сначала создаем виджет для логов
        self.create_log_widget()
        
        # Затем загружаем конфигурацию (теперь log_text существует)
        self.load_config()
        
        # Создаем остальные виджеты
        self.create_widgets()

    def load_config(self):
        """Загрузка конфигурации баз из файла"""
        try:
            self.bases_config = {"""базы из файла"""}
            self.log_message("✅ Конфигурация баз загружена из кода")
        except Exception as e:
            self.log_message(f"❌ Ошибка загрузки конфигурации: {e}")
            self.bases_config = {}

    def create_widgets(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Получение и преобразование данных из iiko", 
                               font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Фрейм для выбора баз
        bases_frame = ttk.LabelFrame(main_frame, text="Выбор баз", padding="10")
        bases_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Список баз с чекбоксами
        self.base_checkbuttons = {}
        for i, (base_name, config) in enumerate(self.bases_config.items()):
            var = tk.BooleanVar()
            self.base_checkbuttons[base_name] = var
            checkbutton = ttk.Checkbutton(bases_frame, text=base_name, variable=var)
            checkbutton.grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
        
        # Кнопки для управления выбором баз
        buttons_frame = ttk.Frame(bases_frame)
        buttons_frame.grid(row=len(self.bases_config), column=0, columnspan=2, pady=10)
        
        ttk.Button(buttons_frame, text="Выбрать все", 
                  command=self.select_all_bases).grid(row=0, column=0, padx=5)
        ttk.Button(buttons_frame, text="Снять выбор", 
                  command=self.deselect_all_bases).grid(row=0, column=1, padx=5)
        
        # Фрейм для выбора листов
        sheets_frame = ttk.LabelFrame(main_frame, text="Выбор листов", padding="10")
        sheets_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Чекбоксы для листов
        self.sheet_vars = {}
        sheets = {
            1: "Свод План-Факт",
            2: "План_факт наполняемость", 
            3: "План_факт стоимость блюд"
        }
        
        for i, (key, value) in enumerate(sheets.items()):
            var = tk.BooleanVar()
            self.sheet_vars[key] = var
            ttk.Checkbutton(sheets_frame, text=value, variable=var).grid(
                row=i, column=0, sticky=tk.W, pady=2)
            
        # Кнопки для управления выбором листов
        sheet_buttons_frame = ttk.Frame(sheets_frame)
        sheet_buttons_frame.grid(row=len(sheets), column=0, columnspan=2, pady=10)
        
        ttk.Button(sheet_buttons_frame, text="Выбрать все", 
                  command=self.select_all_sheets).grid(row=0, column=0, padx=5)
        ttk.Button(sheet_buttons_frame, text="Снять выбор", 
                  command=self.deselect_all_sheets).grid(row=0, column=1, padx=5)
        
        # Фрейм для выбора периода
        period_frame = ttk.LabelFrame(main_frame, text="Выбор периода", padding="10")
        period_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Выбор периода по умолчанию
        period_label = ttk.Label(period_frame, text="Период:")
        period_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        
        self.period_var = tk.StringVar(value="Текущий месяц")
        period_combo = ttk.Combobox(period_frame, textvariable=self.period_var, 
                                   values=["Сегодня", "Вчера", "Текущая неделя", 
                                          "Прошлая неделя", "Текущий месяц", 
                                          "Прошлый месяц", "Текущий год", 
                                          "Прошлый год", "Другой..."])
        period_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        period_combo.bind('<<ComboboxSelected>>', self.on_period_change)
        
        # Календари для выбора дат
        start_label = ttk.Label(period_frame, text="С:")
        start_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 5))
        
        self.start_calendar = DateEntry(period_frame, width=12, background='darkblue', 
                                      foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy')
        self.start_calendar.grid(row=1, column=1, sticky=tk.W, padx=(0, 10))
        
        end_label = ttk.Label(period_frame, text="По:")
        end_label.grid(row=1, column=2, sticky=tk.W, padx=(0, 5))
        
        self.end_calendar = DateEntry(period_frame, width=12, background='darkblue', 
                                    foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy')
        self.end_calendar.grid(row=1, column=3, sticky=tk.W)
        
        # Инициализация календарей
        self.setup_calendars()
        
        # Кнопка запуска
        self.run_button = ttk.Button(main_frame, text="Запустить выгрузку", 
                                    command=self.start_report_generation)
        self.run_button.grid(row=4, column=0, columnspan=2, pady=20)
        
        # Прогресс бар
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Настройка весов для растягивания
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)

    def setup_calendars(self):
        """Инициализация календарей"""
        today = datetime.now()
        self.start_calendar.set_date(today)
        self.end_calendar.set_date(today)
        # Установим начальные значения для текущего месяца
        self.on_period_change(None)

    def on_period_change(self, event=None):
        """Обработчик изменения периода"""
        period = self.period_var.get()
        if period == "Сегодня":
            today = datetime.now()
            self.start_calendar.set_date(today)
            self.end_calendar.set_date(today)
        elif period == "Вчера":
            yesterday = datetime.now() - timedelta(days=1)
            self.start_calendar.set_date(yesterday)
            self.end_calendar.set_date(yesterday)
        elif period == "Текущая неделя":
            today = datetime.now()
            start_of_week = today - timedelta(days=today.weekday())
            end_of_week = start_of_week + timedelta(days=6)
            self.start_calendar.set_date(start_of_week)
            self.end_calendar.set_date(end_of_week)
        elif period == "Прошлая неделя":
            today = datetime.now()
            start_of_last_week = today - timedelta(days=today.weekday() + 7)
            end_of_last_week = start_of_last_week + timedelta(days=6)
            self.start_calendar.set_date(start_of_last_week)
            self.end_calendar.set_date(end_of_last_week)
        elif period == "Текущий месяц":
            today = datetime.now()
            start_of_month = today.replace(day=1)
            end_of_month = today.replace(day=calendar.monthrange(today.year, today.month)[1])
            self.start_calendar.set_date(start_of_month)
            self.end_calendar.set_date(end_of_month)
        elif period == "Прошлый месяц":
            today = datetime.now()
            if today.month == 1:
                start_of_last_month = today.replace(year=today.year - 1, month=12, day=1)
                end_of_last_month = today.replace(year=today.year - 1, month=12, day=31)
            else:
                start_of_last_month = today.replace(month=today.month - 1, day=1)
                end_of_last_month = today.replace(month=today.month - 1,
                                                day=calendar.monthrange(today.year, today.month - 1)[1])
            self.start_calendar.set_date(start_of_last_month)
            self.end_calendar.set_date(end_of_last_month)
        elif period == "Текущий год":
            today = datetime.now()
            start_of_year = today.replace(month=1, day=1)
            end_of_year = today.replace(month=12, day=31)
            self.start_calendar.set_date(start_of_year)
            self.end_calendar.set_date(end_of_year)
        elif period == "Прошлый год":
            today = datetime.now()
            start_of_last_year = today.replace(year=today.year - 1, month=1, day=1)
            end_of_last_year = today.replace(year=today.year - 1, month=12, day=31)
            self.start_calendar.set_date(start_of_last_year)
            self.end_calendar.set_date(end_of_last_year)
        elif period == "Другой...":
            # Добавляем возможность ручного выбора дат
            pass

    def select_all_bases(self):
        """Выбрать все базы"""
        for var in self.base_checkbuttons.values():
            var.set(True)

    def deselect_all_bases(self):
        """Снять выбор со всех баз"""
        for var in self.base_checkbuttons.values():
            var.set(False)

    def select_all_sheets(self):
        """Выбрать все листы"""
        for var in self.sheet_vars.values():
            var.set(True)

    def deselect_all_sheets(self):
        """Снять выбор со всех листов"""
        for var in self.sheet_vars.values():
            var.set(False)

    def get_selected_bases(self):
        """Получить выбранные базы"""
        selected = []
        for base_name, var in self.base_checkbuttons.items():
            if var.get():
                selected.append(base_name)
        return selected

    def get_selected_sheets(self):
        """Получить выбранные листы"""
        selected = []
        for key, var in self.sheet_vars.items():
            if var.get():
                selected.append(key)
        return selected

    def get_start_end_dates(self):
        """Получить начальную и конечную даты"""
        try:
            start_date = self.start_calendar.get_date()
            end_date = self.end_calendar.get_date()
            if start_date > end_date:
                messagebox.showerror("Ошибка", "Дата начала не может быть позже даты окончания")
                return None, None
            return start_date, end_date
        except Exception as e:
            messagebox.showerror("Ошибка", f"Неверно выбран период: {str(e)}")
            return None, None

    def start_report_generation(self):
        """Запуск генерации отчета в отдельном потоке"""
        # Проверяем выбор баз
        self.selected_bases = self.get_selected_bases()
        if not self.selected_bases:
            messagebox.showerror("Ошибка", "Выберите хотя бы одну базу")
            return
        # Проверяем выбор листов
        self.selected_sheets = self.get_selected_sheets()
        if not self.selected_sheets:
            messagebox.showerror("Ошибка", "Выберите хотя бы один лист")
            return
        # Проверяем корректность дат
        self.start_date, self.end_date = self.get_start_end_dates()
        if self.start_date is None or self.end_date is None:
            return
        # Блокируем кнопку и запускаем прогресс
        self.run_button.config(state='disabled')
        self.progress.start()
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self.generate_report, daemon=True)
        thread.start()

    def generate_report(self):
        """Генерация отчета"""
        try:
            self.log_message("=== Начало получения данных из iiko ===")
            # Получаем данные из каждой базы
            all_data = []
            average_data_dict = {}
            for base_name in self.selected_bases:
                config = self.bases_config[base_name]
                self.log_message(f"=== Работаем с базой: {base_name} ===")
                collector = IikoDataCollector(config["url"])
                # Получаем основные данные
                data = collector.get_report_data(config["preset_id"], self.start_date, self.end_date)
                if data and 'data' in data:
                    # Добавляем имя базы к каждому элементу
                    for item in data['data']:
                        item['BaseName'] = base_name
                    all_data.extend(data['data'])
                # Получаем данные по average_id
                avg_data = collector.get_average_report_data(config["average_id"], self.start_date, self.end_date)
                if avg_data and 'data' in avg_data:
                    average_data_dict[base_name] = avg_data['data']
            if all_data:
                # Создаем папку для отчетов
                Path("Отчёты iiko").mkdir(exist_ok=True)
                # Формируем имя файла
                current_date = datetime.now()
                report_date_str = current_date.strftime("%d_%m_%Y")
                timestamp = int(time.time())
                documents_path = Path.home() / "Documents"
                reports_path = documents_path / "Отчёты iiko"
                reports_path.mkdir(exist_ok=True)
                excel_filename = reports_path / f"Свод_План_Факт_{report_date_str}_{timestamp}.xlsx"
                # Создаем отчет, используя метод из ReportGenerator
                success = self.report_generator.create_excel_report(all_data, self.bases_config, excel_filename,
                                                 self.start_date, self.end_date, average_data_dict,
                                                 self.selected_sheets)
                if success:
                    self.log_message(f"✅ Отчет создан: {excel_filename}")
                    messagebox.showinfo("Успех", f"Отчет успешно создан!\n{excel_filename}")
                else:
                    self.log_message("❌ Ошибка при создании отчета")
                    messagebox.showerror("Ошибка", "Не удалось создать отчет")
            else:
                self.log_message("❌ Нет данных для создания отчета")
                messagebox.showwarning("Предупреждение", "Нет данных для создания отчета")
        except Exception as e:
            self.log_message(f"❌ Ошибка: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
        finally:
            # Разблокируем кнопку и останавливаем прогресс
            self.root.after(0, self.finish_report_generation)

    def finish_report_generation(self):
        """Завершение генерации отчета"""
        self.progress.stop()
        self.run_button.config(state='normal')

    def log_message(self, message):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        if hasattr(self, 'log_text'):
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            self.log_text.see(tk.END)
            self.root.update_idletasks()
        else:
            print(f"[{timestamp}] {message}")

    def create_log_widget(self):
        """Создание текстового виджета для логов"""
        log_frame = ttk.LabelFrame(self.root, text="Лог выполнения", padding="10")
        log_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=5, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        # Настройка весов
        self.root.rowconfigure(1, weight=1)
