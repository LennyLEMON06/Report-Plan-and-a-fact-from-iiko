# ui.py — графический интерфейс

import tkinter as tk
from tkinter import ttk, messagebox
import threading
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import calendar
import time
from pathlib import Path
from config import BASES_CONFIG
from iiko_collector import IikoDataCollector
from report_generator import ReportGenerator

class IikoReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Получение данных из iiko")
        self.root.geometry("850x570")
        self.root.resizable(True, True)

        self.bases_config = BASES_CONFIG
        self.selected_bases = []
        self.selected_sheets = []
        self.start_date = None
        self.end_date = None

        self.create_log_widget()
        self.create_widgets()

    def create_log_widget(self):
        log_frame = ttk.LabelFrame(self.root, text="Лог выполнения", padding="10")
        log_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, height=5, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.root.rowconfigure(1, weight=1)

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        if hasattr(self, 'log_text'):
            self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            self.log_text.see(tk.END)
            self.root.update_idletasks()
        else:
            print(f"[{timestamp}] {message}")

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(0, weight=1, minsize=250)
        main_frame.columnconfigure(1, weight=2, minsize=300)
        main_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=0)
        main_frame.rowconfigure(2, weight=1)

        title_label = ttk.Label(main_frame, text="Получение и преобразование данных из iiko", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky=tk.N)

        # Левая колонка: выбор баз
        bases_frame = ttk.LabelFrame(main_frame, text="Выбор баз", padding="10")
        bases_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5), pady=(0, 10))
        bases_frame.columnconfigure(0, weight=1)
        bases_frame.rowconfigure(0, weight=1)

        canvas = tk.Canvas(bases_frame)
        scrollbar = ttk.Scrollbar(bases_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        self.base_checkbuttons = {}
        for i, base_name in enumerate(self.bases_config.keys()):
            var = tk.BooleanVar()
            self.base_checkbuttons[base_name] = var
            ttk.Checkbutton(scrollable_frame, text=base_name, variable=var).grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)

        buttons_frame = ttk.Frame(scrollable_frame)
        buttons_frame.grid(row=len(self.bases_config), column=0, columnspan=2, pady=10)
        ttk.Button(buttons_frame, text="Выбрать все", command=self.select_all_bases).grid(row=0, column=0, padx=5)
        ttk.Button(buttons_frame, text="Снять выбор", command=self.deselect_all_bases).grid(row=0, column=1, padx=5)

        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        bases_frame.rowconfigure(0, weight=1)
        bases_frame.columnconfigure(0, weight=1)

        # Правая колонка
        right_frame = ttk.Frame(main_frame)
        right_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0), pady=(0, 10))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=0)
        right_frame.rowconfigure(2, weight=0)

        # Выбор листов
        sheets_frame = ttk.LabelFrame(right_frame, text="Выбор листов", padding="10")
        sheets_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        sheets_frame.columnconfigure(0, weight=1)

        self.sheet_vars = {}
        sheets = {
            1: "Свод План-Факт",
            2: "План_факт наполняемость",
            3: "План_факт стоимость блюд",
            4: "План_факт гостепоток"
        }
        for i, (key, value) in enumerate(sheets.items()):
            var = tk.BooleanVar()
            self.sheet_vars[key] = var
            ttk.Checkbutton(sheets_frame, text=value, variable=var).grid(row=i, column=0, sticky=tk.W, pady=2)

        sheet_buttons_frame = ttk.Frame(sheets_frame)
        sheet_buttons_frame.grid(row=len(sheets), column=0, columnspan=2, pady=10)
        ttk.Button(sheet_buttons_frame, text="Выбрать все", command=self.select_all_sheets).grid(row=0, column=0, padx=5)
        ttk.Button(sheet_buttons_frame, text="Снять выбор", command=self.deselect_all_sheets).grid(row=0, column=1, padx=5)

        # Выбор периода
        period_frame = ttk.LabelFrame(right_frame, text="Выбор периода", padding="10")
        period_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(period_frame, text="Период:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        self.period_var = tk.StringVar(value="Текущий месяц")
        period_combo = ttk.Combobox(period_frame, textvariable=self.period_var,
                                   values=["Сегодня", "Вчера", "Текущая неделя", "Прошлая неделя", "Текущий месяц",
                                          "Прошлый месяц", "Текущий год", "Прошлый год", "Другой..."])
        period_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        period_combo.bind('<<ComboboxSelected>>', self.on_period_change)

        ttk.Label(period_frame, text="С:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5))
        self.start_calendar = DateEntry(period_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy')
        self.start_calendar.grid(row=1, column=1, sticky=tk.W, padx=(0, 10))

        ttk.Label(period_frame, text="По:").grid(row=1, column=2, sticky=tk.W, padx=(0, 5))
        self.end_calendar = DateEntry(period_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd.mm.yyyy')
        self.end_calendar.grid(row=1, column=3, sticky=tk.W)

        self.setup_calendars()

        # Кнопка и прогресс
        button_frame = ttk.Frame(right_frame)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        button_frame.columnconfigure(0, weight=1)

        self.run_button = ttk.Button(button_frame, text="Запустить выгрузку", command=self.start_report_generation)
        self.run_button.grid(row=0, column=0, pady=5, sticky=(tk.W, tk.E))

        self.progress = ttk.Progressbar(button_frame, mode='indeterminate')
        self.progress.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))

        # Лог (уже создан)
        log_frame = self.log_text.master
        log_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), padx=0, pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    def setup_calendars(self):
        today = datetime.now()
        self.start_calendar.set_date(today)
        self.end_calendar.set_date(today)
        self.on_period_change()

    def on_period_change(self, event=None):
        period = self.period_var.get()
        today = datetime.now()

        if period == "Сегодня":
            self.start_calendar.set_date(today)
            self.end_calendar.set_date(today)
        elif period == "Вчера":
            yesterday = today - timedelta(days=1)
            self.start_calendar.set_date(yesterday)
            self.end_calendar.set_date(yesterday)
        elif period == "Текущая неделя":
            start = today - timedelta(days=today.weekday())
            end = start + timedelta(days=6)
            self.start_calendar.set_date(start)
            self.end_calendar.set_date(end)
        elif period == "Прошлая неделя":
            start = today - timedelta(days=today.weekday() + 7)
            end = start + timedelta(days=6)
            self.start_calendar.set_date(start)
            self.end_calendar.set_date(end)
        elif period == "Текущий месяц":
            start = today.replace(day=1)
            end = today.replace(day=calendar.monthrange(today.year, today.month)[1])
            self.start_calendar.set_date(start)
            self.end_calendar.set_date(end)
        elif period == "Прошлый месяц":
            if today.month == 1:
                start = today.replace(year=today.year-1, month=12, day=1)
                end = today.replace(year=today.year-1, month=12, day=31)
            else:
                start = today.replace(month=today.month-1, day=1)
                end = today.replace(month=today.month-1, day=calendar.monthrange(today.year, today.month-1)[1])
            self.start_calendar.set_date(start)
            self.end_calendar.set_date(end)
        elif period == "Текущий год":
            start = today.replace(month=1, day=1)
            end = today.replace(month=12, day=31)
            self.start_calendar.set_date(start)
            self.end_calendar.set_date(end)
        elif period == "Прошлый год":
            start = today.replace(year=today.year-1, month=1, day=1)
            end = today.replace(year=today.year-1, month=12, day=31)
            self.start_calendar.set_date(start)
            self.end_calendar.set_date(end)

    def select_all_bases(self):
        for var in self.base_checkbuttons.values():
            var.set(True)

    def deselect_all_bases(self):
        for var in self.base_checkbuttons.values():
            var.set(False)

    def select_all_sheets(self):
        for var in self.sheet_vars.values():
            var.set(True)

    def deselect_all_sheets(self):
        for var in self.sheet_vars.values():
            var.set(False)

    def get_selected_bases(self):
        return [name for name, var in self.base_checkbuttons.items() if var.get()]

    def get_selected_sheets(self):
        return [key for key, var in self.sheet_vars.items() if var.get()]

    def get_start_end_dates(self):
        try:
            start = self.start_calendar.get_date()
            end = self.end_calendar.get_date()
            if start > end:
                messagebox.showerror("Ошибка", "Дата начала не может быть позже даты окончания")
                return None, None
            return start, end
        except Exception as e:
            messagebox.showerror("Ошибка", f"Неверно выбран период: {str(e)}")
            return None, None

    def start_report_generation(self):
        self.selected_bases = self.get_selected_bases()
        if not self.selected_bases:
            messagebox.showerror("Ошибка", "Выберите хотя бы одну базу")
            return

        self.selected_sheets = self.get_selected_sheets()
        if not self.selected_sheets:
            messagebox.showerror("Ошибка", "Выберите хотя бы один лист")
            return

        self.start_date, self.end_date = self.get_start_end_dates()
        if self.start_date is None or self.end_date is None:
            return

        self.run_button.config(state='disabled')
        self.progress.start()

        thread = threading.Thread(target=self.generate_report, daemon=True)
        thread.start()

    def generate_report(self):
        try:
            self.log_message("=== Начало получения данных из iiko ===")
            all_data = []
            average_data_dict = {}

            for base_name in self.selected_bases:
                config = self.bases_config[base_name]
                self.log_message(f"=== Работаем с базой: {base_name} ===")
                collector = IikoDataCollector(config["url"])

                data = collector.get_report_data(config["preset_id"], self.start_date, self.end_date)
                if data and 'data' in data:
                    for item in data['data']:
                        item['BaseName'] = base_name
                    all_data.extend(data['data'])

                avg_data = collector.get_report_data(config["preset_id"], self.start_date, self.end_date)
                if avg_data and 'data' in avg_data:
                    average_data_dict[base_name] = avg_data['data']

            if all_data:
                current_date = datetime.now()
                report_date_str = current_date.strftime("%d_%m_%Y")
                timestamp = int(time.time())
                documents_path = Path.home() / "Documents"
                reports_path = documents_path / "Отчёты iiko"
                reports_path.mkdir(exist_ok=True)
                excel_filename = reports_path / f"Свод_План_Факт_{report_date_str}_{timestamp}.xlsx"

                generator = ReportGenerator(log_callback=self.log_message)
                success = generator.create_excel_report(
                    all_data, self.bases_config, excel_filename,
                    self.start_date, self.end_date, average_data_dict, self.selected_sheets
                )

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
            self.root.after(0, self.finish_report_generation)

    def finish_report_generation(self):
        self.progress.stop()
        self.run_button.config(state='normal')