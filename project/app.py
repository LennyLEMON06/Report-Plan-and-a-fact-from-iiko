import requests
import hashlib
from datetime import datetime
import json
import urllib3
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from datetime import timedelta
import calendar

# Отключение предупреждений о сертификатах
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class IikoDataCollector:
    def __init__(self, base_url, login="login", password="password"):
        self.base_url = base_url.strip()
        self.login = login
        self.password = password
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

    def get_report_data(self, preset_id, date_from, date_to, save_raw=False):
        """Получение данных отчета"""
        if not self.token and not self.auth():
            return None

        url = f"{self.base_url}/v2/reports/olap/byPresetId/{preset_id}"
        
        params = {
            'key': self.token,
            'dateFrom': date_from.strftime('%Y-%m-%d'),
            'dateTo': date_to.strftime('%Y-%m-%d')
        }

        try:
            print(f"Отправляем запрос в {self.base_url}...")
            response = self.session.get(url, params=params)
            
            if response.status_code == 200:
                print("✅ Данные успешно получены")
                json_data = response.json()
                
                if save_raw:
                    Path("raw_data").mkdir(exist_ok=True)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"raw_data/report_{timestamp}.json"
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

def load_bases_config():
    """Загрузка конфигурации баз из файла"""
    try:
        with open("bases_config.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"❌ Ошибка загрузки bases_config.json: {e}")
        return {}

def select_bases(bases):
    """Выбор баз для выгрузки"""
    print("\nДоступные базы:")
    keys = list(bases.keys())
    for i, name in enumerate(keys, 1):
        print(f"{i}. {name}")
    
    selected = input("\nВведите номера баз через запятую (например: 1,3,5) или 'all' для всех: ").strip()
    
    if selected.lower() == "all":
        return keys
    
    try:
        indices = [int(x.strip()) - 1 for x in selected.split(",")]
        return [keys[i] for i in indices if 0 <= i < len(keys)]
    except:
        print("❌ Неверный ввод")
        return []

def create_excel_report(all_data, output_filename):
    """Создание Excel отчета: все кассы на одном листе"""
    if not all_data:
        print("❌ Нет данных для создания отчета")
        return False
    
    # Создаем DataFrame из всех данных
    df = pd.DataFrame(all_data)
    
    # Преобразуем дату в правильный формат
    df['OpenDate.Typed'] = pd.to_datetime(df['OpenDate.Typed'])
    df['Date'] = df['OpenDate.Typed'].dt.strftime('%d.%m.%Y')
    df['DayOfWeek'] = df['DayOfWeekOpen'].str.extract(r'(\d+)\.')[0].astype(int)
    
    # Сортируем по дате и дню недели
    df = df.sort_values(['OpenDate.Typed', 'DayOfWeek'])
    
    # Получаем уникальные кассы
    cash_registers = df['Store.Name'].unique()
    
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
    headers = ["День недели", "Дата", "План", "Факт", "% выполнения"]
    
    col_offset = 0
    for cash_register in cash_registers:
        start_col = col_offset + 1
        end_col = col_offset + len(headers)
        ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
        ws.cell(row=2, column=start_col, value=cash_register)
        ws.cell(row=2, column=start_col).font = Font(bold=True)
        ws.cell(row=2, column=start_col).alignment = Alignment(horizontal='center')
        ws.cell(row=2, column=start_col).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        
        for i, header in enumerate(headers):
            ws.cell(row=3, column=col_offset + i + 1, value=header)
            ws.cell(row=3, column=col_offset + i + 1).font = Font(bold=True)
            ws.cell(row=3, column=col_offset + i + 1).fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        col_offset += len(headers)

    # Заполняем данные
    dates = sorted(df['OpenDate.Typed'].unique())
    
    current_row = 4
    for date in dates:
        day_of_week = df[df['OpenDate.Typed'] == date]['DayOfWeekOpen'].iloc[0]
        formatted_date = df[df['OpenDate.Typed'] == date]['Date'].iloc[0]
        
        col_offset = 0
        for cash_register in cash_registers:
            ws.cell(row=current_row, column=col_offset + 1, value=day_of_week)
            ws.cell(row=current_row, column=col_offset + 2, value=formatted_date)
            ws.cell(row=current_row, column=col_offset + 3, value="")
            
            cash_data = df[(df['OpenDate.Typed'] == date) & (df['Store.Name'] == cash_register)]
            if not cash_data.empty:
                fact_value = cash_data.iloc[0]['DishDiscountSumInt']
                ws.cell(row=current_row, column=col_offset + 4, value=fact_value)
                
                plan_col = get_column_letter(col_offset + 3)
                fact_col = get_column_letter(col_offset + 4)
                formula = f"=IF({plan_col}{current_row}<>0, {fact_col}{current_row}/{plan_col}{current_row}, 0)"
                ws.cell(row=current_row, column=col_offset + 5, value=formula)
                ws.cell(row=current_row, column=col_offset + 5).number_format = '0.00%'
            
            col_offset += len(headers)
        
        current_row += 1

    # Итоговая строка
    total_row = current_row
    ws.cell(row=total_row, column=1, value="ИТОГО")
    ws.cell(row=total_row, column=1).font = Font(bold=True)
    
    col_offset = 0
    for cash_register in cash_registers:
        start_col = col_offset + 4
        col_letter = get_column_letter(start_col)
        ws.cell(row=total_row, column=start_col, value=f"=SUM({col_letter}4:{col_letter}{current_row-1})")
        ws.cell(row=total_row, column=start_col).font = Font(bold=True)
        
        ws.cell(row=total_row, column=col_offset + 3, value="")
        
        percent_col = col_offset + 5
        percent_letter = get_column_letter(percent_col)
        ws.cell(row=total_row, column=percent_col, value=f"=AVERAGE({percent_letter}4:{percent_letter}{current_row-1})")
        ws.cell(row=total_row, column=percent_col).number_format = '0.00%'
        ws.cell(row=total_row, column=percent_col).font = Font(bold=True)
        
        col_offset += len(headers)

    # Форматирование
    for col in range(1, col_offset + 1):
        col_letter = get_column_letter(col)
        ws.column_dimensions[col_letter].width = 15

    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))
    
    for row in range(2, total_row + 1):
        for col in range(1, col_offset + 1):
            ws.cell(row=row, column=col).border = thin_border

    # Условное форматирование
    col_offset = 0
    for cash_register in cash_registers:
        percent_col = col_offset + 5
        percent_letter = get_column_letter(percent_col)
        percent_range = f"{percent_letter}4:{percent_letter}{current_row-1}"

        rule = CellIsRule(
            operator='greaterThan',
            formula=['1'],
            stopIfTrue=True,
            fill=PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        )
        ws.conditional_formatting.add(percent_range, rule)
        col_offset += len(headers)

    # Сохраняем файл
    wb.save(output_filename)
    print(f"✅ Отчет сохранен: {output_filename}")
    return True

def main():
    print("=== Получение и преобразование данных из iiko ===")
    
    # Загружаем конфиг
    bases = load_bases_config()
    if not bases:
        print("❌ Не удалось загрузить конфигурацию баз")
        return

    # Выбираем базы
    selected_bases = select_bases(bases)
    if not selected_bases:
        print("❌ Не выбрано ни одной базы")
        return

    # Устанавливаем период
    year = 2025
    month = 6
    first_day = 1
    last_day = calendar.monthrange(year, month)[1]
    date_from = datetime(year, month, first_day)
    date_to = datetime(year, month, last_day) + timedelta(days=1)

    all_data = []
    
    # Получаем данные из каждой базы
    for base_name in selected_bases:
        config = bases[base_name]
        print(f"\n=== Работаем с базой: {base_name} ===")
        
        collector = IikoDataCollector(config["url"])
        data = collector.get_report_data(config["preset_id"], date_from, date_to)
        
        if data and 'data' in data:
            # Добавляем имя базы к каждому элементу
            for item in data['data']:
                item['BaseName'] = base_name
            all_data.extend(data['data'])

    if all_data:
        # Создаем папку для отчетов
        Path("reports").mkdir(exist_ok=True)
        
        # Формируем имя файла
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"reports/report_plan_fact_{timestamp}.xlsx"
        
        # Создаем отчет
        create_excel_report(all_data, excel_filename)
        
        print(f"\n✅ Отчет создан: {excel_filename}")
    else:
        print("❌ Нет данных для создания отчета")

    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()