import requests
import hashlib
from datetime import datetime
import json
import urllib3
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Color
from datetime import timedelta
import calendar

# Отключение предупреждений о сертификатах
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class IikoDataCollector:
    def __init__(self):
        self.base_url = "https://name.iiko.it:port/resto/api"
        self.login = "login"
        self.password = "password"
        self.token = None
        self.session = requests.Session()
        self.session.verify = False

    def auth(self):
        """Аутентификация в системе iiko"""
        auth_url = f"{self.base_url}/auth"
        try:
            password_hash = hashlib.sha1(self.password.encode()).hexdigest()
            response = self.session.post(
                auth_url,
                data={'login': self.login, 'pass': password_hash},
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            )
            if response.status_code == 200:
                self.token = response.text.strip()
                print("✅ Авторизация успешна")
                return True
            print(f"❌ Ошибка авторизации: {response.status_code} - {response.text}")
            return False
        except Exception as e:
            print(f"❌ Ошибка подключения: {str(e)}")
            return False

    def get_report_data(self, date_from, date_to, save_raw=True):
        """Получение данных отчета и сохранение сырого JSON"""
        if not self.token and not self.auth():
            return None

        preset_id = "874fe793-d575-40fb-b8ff-8d78c9a280b8"
        url = f"{self.base_url}/v2/reports/olap/byPresetId/{preset_id}"
        
        params = {
            'key': self.token,
            'dateFrom': date_from.strftime('%Y-%m-%d'),
            'dateTo': date_to.strftime('%Y-%m-%d')
        }

        try:
            print(f"Отправляем запрос с параметрами: {params}")
            response = self.session.get(url, params=params)
            
            if response.status_code == 200:
                print("✅ Данные успешно получены")
                json_data = response.json()
                
                if save_raw:
                    # Создаем папку для сохранения, если ее нет
                    Path("raw_data").mkdir(exist_ok=True)
                    
                    # Формируем имя файла
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"raw_data/report_{timestamp}.json"
                    
                    # Сохраняем сырые данные
                    with open(filename, 'w', encoding='utf-8') as f:
                        json.dump(json_data, f, ensure_ascii=False, indent=2)
                    print(f"✅ Сырые данные сохранены в файл: {filename}")
                
                return json_data
            
            print(f"❌ Ошибка получения отчета: {response.status_code}")
            print(f"Ответ сервера: {response.text}")
            return None
            
        except Exception as e:
            print(f"❌ Ошибка при получении отчета: {str(e)}")
            return None

def create_excel_report(json_data, output_filename):
    """Создание Excel отчета: все кассы на одном листе, как на фото"""
    if not json_data or 'data' not in json_data:
        print("❌ Нет данных для создания отчета")
        return False
    
    data = json_data['data']
    
    # Создаем DataFrame из данных
    df = pd.DataFrame(data)
    
    # Преобразуем дату в правильный формат
    df['OpenDate.Typed'] = pd.to_datetime(df['OpenDate.Typed'])
    df['Date'] = df['OpenDate.Typed'].dt.strftime('%d.%m.%Y')
    df['DayOfWeek'] = df['DayOfWeekOpen'].str.extract(r'(\d+)\.')[0].astype(int)
    
    # Сортируем по дате и дню недели
    df = df.sort_values(['OpenDate.Typed', 'DayOfWeek'])
    
    # Получаем уникальные кассы
    cash_registers = df['CashRegisterName'].unique()
    
    # Создаем Excel файл
    wb = Workbook()
    
    # Удаляем все существующие листы
    while len(wb.sheetnames) > 0:
        del wb[wb.sheetnames[0]]
    
    ws = wb.create_sheet(title="Свод План-Факт")

    # Заголовок
    ws['A1'] = "СВОДНЫЙ ОТЧЕТ ПО ВЫРУЧКЕ"
    ws['A1'].font = Font(bold=True, size=16)
    ws.merge_cells('A1:E1')

    # Группируем данные по кассам
    # Для каждой кассы будет блок: День недели, Дата, План, Факт, % выполнения
    headers = ["День недели", "Дата", "План", "Факт", "% выполнения"]
    
    # Определяем стартовые столбцы для каждой кассы
    col_offset = 0
    for cash_register in cash_registers:
        # Заголовки для этой кассы
        start_col = col_offset + 1
        end_col = col_offset + len(headers)
        ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
        ws.cell(row=2, column=start_col, value=cash_register)
        ws.cell(row=2, column=start_col).font = Font(bold=True)
        ws.cell(row=2, column=start_col).alignment = Alignment(horizontal='center')
        ws.cell(row=2, column=start_col).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        
        # Подзаголовки
        for i, header in enumerate(headers):
            ws.cell(row=3, column=col_offset + i + 1, value=header)
            ws.cell(row=3, column=col_offset + i + 1).font = Font(bold=True)
            ws.cell(row=3, column=col_offset + i + 1).fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        col_offset += len(headers)

    # Заполняем данные
    # Сначала получим все уникальные даты
    dates = sorted(df['OpenDate.Typed'].unique())
    
    current_row = 4
    for date in dates:
        # День недели и дата
        day_of_week = df[df['OpenDate.Typed'] == date]['DayOfWeekOpen'].iloc[0]
        formatted_date = df[df['OpenDate.Typed'] == date]['Date'].iloc[0]
        
        # Записываем день недели и дату для всех касс
        col_offset = 0
        for cash_register in cash_registers:
            # День недели и дата
            ws.cell(row=current_row, column=col_offset + 1, value=day_of_week)
            ws.cell(row=current_row, column=col_offset + 2, value=formatted_date)
            
            # План — пусто
            ws.cell(row=current_row, column=col_offset + 3, value="")
            
            # Факт — DishDiscountSumInt
            cash_data = df[(df['OpenDate.Typed'] == date) & (df['CashRegisterName'] == cash_register)]
            if not cash_data.empty:
                fact_value = cash_data.iloc[0]['DishDiscountSumInt']
                ws.cell(row=current_row, column=col_offset + 4, value=fact_value)
                
                # % выполнения — формула: Факт / План
                plan_col = get_column_letter(col_offset + 3)
                fact_col = get_column_letter(col_offset + 4)
                formula = f"=IF({plan_col}{current_row}<>0, {fact_col}{current_row}/{plan_col}{current_row}, 0)"
                ws.cell(row=current_row, column=col_offset + 5, value=formula)
                
                # Форматирование процента
                ws.cell(row=current_row, column=col_offset + 5).number_format = '0.00%'
            
            col_offset += len(headers)
        
        current_row += 1

    # Добавляем итоговую строку
    total_row = current_row
    ws.cell(row=total_row, column=1, value="ИТОГО")
    ws.cell(row=total_row, column=1).font = Font(bold=True)
    
    col_offset = 0
    for cash_register in cash_registers:
        # Только для столбца "Факт" — сумма
        start_col = col_offset + 4  # Это столбец "Факт"
        col_letter = get_column_letter(start_col)
        ws.cell(row=total_row, column=start_col, value=f"=SUM({col_letter}4:{col_letter}{current_row-1})")
        ws.cell(row=total_row, column=start_col).font = Font(bold=True)
        
        # Для столбца "План" — пусто
        ws.cell(row=total_row, column=col_offset + 3, value="")
        
        # Для столбца "% выполнения" — среднее значение
        percent_col = col_offset + 5
        percent_letter = get_column_letter(percent_col)
        ws.cell(row=total_row, column=percent_col, value=f"=AVERAGE({percent_letter}4:{percent_letter}{current_row-1})")
        ws.cell(row=total_row, column=percent_col).number_format = '0.00%'
        ws.cell(row=total_row, column=percent_col).font = Font(bold=True)
        
        col_offset += len(headers)

    # Форматирование ширины колонок
    for col in range(1, col_offset + 1):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 15

    # Добавляем границы для лучшей читаемости
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    
    for row in range(2, total_row + 1):
        for col in range(1, col_offset + 1):
            ws.cell(row=row, column=col).border = thin_border

    # Условное форматирование: если % выполнения > 100%, то желтый фон
    col_offset = 0
    for cash_register in cash_registers:
        percent_col = col_offset + 5  # Столбец "% выполнения"
        percent_letter = get_column_letter(percent_col)
        
        # Диапазон данных по процентам (без заголовка и итога)
        percent_range = f"{percent_letter}4:{percent_letter}{current_row-1}"

        # Правило: если значение > 1 (100%), то желтый фон
        rule = CellIsRule(
            operator='greaterThan',
            formula=['1'],
            stopIfTrue=True,
            fill=PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # Желтый цвет
        )

        # Применяем правило к диапазону
        ws.conditional_formatting.add(percent_range, rule)

        col_offset += len(headers)

    # Сохраняем файл
    wb.save(output_filename)
    print(f"✅ Отчет сохранен: {output_filename}")
    return True

def main():
    print("=== Получение и преобразование данных из iiko ===")
    collector = IikoDataCollector()

    # Авторизация
    if not collector.auth():
        input("Нажмите Enter для выхода...")
        return

    # Устанавливаем период с 1 июля по 31 июля 2025 года
    date_from = datetime(2025, 6, 1)
    date_to = datetime(2025, 6, 30) + timedelta(days=1)  # +1 день

    # Получаем данные
    print(f"\nПолучаем данные отчета за период: {date_from.strftime('%d.%m.%Y')}-{date_to.strftime('%d.%m.%Y')}...")
    report_data = collector.get_report_data(date_from, date_to)

    if report_data:
        # Создаем папку для отчетов
        Path("reports").mkdir(exist_ok=True)
        
        # Формируем имя файла
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Создаем отчет
        excel_filename = f"reports/report_plan_fact_{timestamp}.xlsx"
        create_excel_report(report_data, excel_filename)
        
        print("\nОтчет создан:")
        print(f"1. План-Факт: {excel_filename}")

    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()