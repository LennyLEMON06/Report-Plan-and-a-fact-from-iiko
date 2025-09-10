# report_generator.py
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime
import time
from pathlib import Path

# Класс или функции для создания отчетов
class ReportGenerator:
    def __init__(self, log_callback=None):
        self.log_callback = log_callback

    def log_message(self, message):
        """Функция для логирования, вызывает внешнюю функцию логирования, если задана"""
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message) # fallback, если логгер не передан

    def create_main_report_sheet(self, wb, df, cash_registers):
        """Создание основного листа Свод План-Факт"""
        ws = wb.create_sheet(title="Свод План-Факт")
        # Заголовок
        ws['A1'] = "СВОДНЫЙ ОТЧЕТ ПО ВЫРУЧКЕ"
        ws['A1'].font = Font(bold=True, size=16)
        ws.merge_cells('A1:E1')
        # Цвета для разных баз
        base_colors = {
            0: "E2EFDA",  # Зеленый
            1: "FFEB9C",  # Желтый
            2: "C6E0B4",  # Светло-зеленый
            3: "FFE699",  # Светло-желтый
            4: "A9D08E",  # Зеленый
            5: "F8CBAD",  # Оранжевый
            6: "9CC2E5",  # Голубой
            7: "D9E1F2",  # Светло-синий
        }
        # Заголовки для каждой кассы
        headers = ["День недели", "Дата", "План", "Факт", "% выполнения"]
        col_offset = 0
        base_index = 0
        for cash_register in cash_registers:
            start_col = col_offset + 1
            end_col = col_offset + len(headers)
            ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
            ws.cell(row=2, column=start_col, value=cash_register)
            ws.cell(row=2, column=start_col).font = Font(bold=True)
            ws.cell(row=2, column=start_col).alignment = Alignment(horizontal='center')
            # Уникальный цвет для каждой базы
            color_index = base_index % len(base_colors)
            color = base_colors[color_index]
            ws.cell(row=2, column=start_col).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            # Заголовки столбцов
            for i, header in enumerate(headers):
                cell = ws.cell(row=3, column=col_offset + i + 1, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
            col_offset += len(headers)
            base_index += 1
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
                # План (берем из данных или 0 если нет)
                cash_data = df[(df['OpenDate.Typed'] == date) & (df['Store.Name'] == cash_register)]
                plan_value = cash_data.iloc[0]['PlanValue'] if not cash_data.empty and 'PlanValue' in cash_data.iloc[0] else 0
                ws.cell(row=current_row, column=col_offset + 3, value=plan_value)
                # Факт
                if not cash_data.empty:
                    fact_value = cash_data.iloc[0]['DishDiscountSumInt']
                    ws.cell(row=current_row, column=col_offset + 4, value=fact_value)
                    # Процент выполнения — ВСЕГДА формула
                    plan_col = get_column_letter(col_offset + 3)
                    fact_col = get_column_letter(col_offset + 4)
                    formula = f"=IF({plan_col}{current_row}=0, 0, {fact_col}{current_row}/{plan_col}{current_row})"
                    ws.cell(row=current_row, column=col_offset + 5, value=formula)
                    ws.cell(row=current_row, column=col_offset + 5).number_format = '0.00%'
                col_offset += len(headers)
            current_row += 1
        # Итоговая строка
        total_row = current_row
        ws.cell(row=total_row, column=1, value="ИТОГО")
        ws.cell(row=total_row, column=1).font = Font(bold=True)
        # Дата выгрузки в столбце "Дата" для строки ИТОГО
        current_date = datetime.now().strftime('%d.%m.%Y')
        ws.cell(row=total_row, column=2, value=current_date)
        ws.cell(row=total_row, column=2).font = Font(bold=True)
        col_offset = 0
        for cash_register in cash_registers:
            # Сумма плана
            plan_col = get_column_letter(col_offset + 3)
            ws.cell(row=total_row, column=col_offset + 3, value=f"=SUM({plan_col}4:{plan_col}{current_row-1})")
            ws.cell(row=total_row, column=col_offset + 3).font = Font(bold=True)
            # Сумма факта
            fact_col = get_column_letter(col_offset + 4)
            ws.cell(row=total_row, column=col_offset + 4, value=f"=SUM({fact_col}4:{fact_col}{current_row-1})")
            ws.cell(row=total_row, column=col_offset + 4).font = Font(bold=True)
            # Процент выполнения (факт/план для итоговой строки)
            plan_total_col = get_column_letter(col_offset + 3)
            fact_total_col = get_column_letter(col_offset + 4)
            formula = f"=IF({plan_total_col}{current_row}=0, 0, {fact_total_col}{current_row}/{plan_total_col}{current_row})"
            ws.cell(row=total_row, column=col_offset + 5, value=formula)
            ws.cell(row=total_row, column=col_offset + 5).number_format = '0.00%'
            ws.cell(row=total_row, column=col_offset + 5).font = Font(bold=True)
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
        # Условное форматирование для процентов выполнения
        col_offset = 0
        for cash_register in cash_registers:
            percent_col = col_offset + 5
            percent_letter = get_column_letter(percent_col)
            percent_range = f"{percent_letter}4:{percent_letter}{current_row-1}"
            # Зеленый для >= 100%
            rule_green = CellIsRule(
                operator='greaterThanOrEqual',
                formula=['1'],
                stopIfTrue=False,
                fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            )
            ws.conditional_formatting.add(percent_range, rule_green)
            # Красный для < 100%
            rule_red = CellIsRule(
                operator='lessThan',
                formula=['1'],
                stopIfTrue=False,
                fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            )
            ws.conditional_formatting.add(percent_range, rule_red)
            col_offset += len(headers)

    def create_occupancy_report_sheet(self, wb, df, cash_registers, bases_config, average_data_dict):
        """Создание листа План_факт наполняемость"""
        ws = wb.create_sheet(title="План_факт наполняемость")
        # Заголовок
        ws['A1'] = "ПЛАН-ФАКТ НАПОЛНЯЕМОСТЬ"
        ws['A1'].font = Font(bold=True, size=16)
        ws.merge_cells('A1:E1')
        # Цвета для разных баз
        base_colors = {
            0: "E2EFDA",  # Зеленый
            1: "FFEB9C",  # Желтый
            2: "C6E0B4",  # Светло-зеленый
            3: "FFE699",  # Светло-желтый
            4: "A9D08E",  # Зеленый
            5: "F8CBAD",  # Оранжевый
            6: "9CC2E5",  # Голубой
            7: "D9E1F2",  # Светло-синий
        }
        # Заголовки для каждой кассы
        headers = ["День недели", "Дата", "План", "Факт", "% выполнения"]
        col_offset = 0
        base_index = 0
        for cash_register in cash_registers:
            start_col = col_offset + 1
            end_col = col_offset + len(headers)
            ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
            ws.cell(row=2, column=start_col, value=cash_register)
            ws.cell(row=2, column=start_col).font = Font(bold=True)
            ws.cell(row=2, column=start_col).alignment = Alignment(horizontal='center')
            # Уникальный цвет для каждой базы
            color_index = base_index % len(base_colors)
            color = base_colors[color_index]
            ws.cell(row=2, column=start_col).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            # Заголовки столбцов
            for i, header in enumerate(headers):
                cell = ws.cell(row=3, column=col_offset + i + 1, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
            col_offset += len(headers)
            base_index += 1
        # Заполняем данные (наполняемость)
        dates = sorted(df['OpenDate.Typed'].unique())
        current_row = 4
        for date in dates:
            day_of_week = df[df['OpenDate.Typed'] == date]['DayOfWeekOpen'].iloc[0]
            formatted_date = df[df['OpenDate.Typed'] == date]['Date'].iloc[0]
            col_offset = 0
            for cash_register in cash_registers:
                ws.cell(row=current_row, column=col_offset + 1, value=day_of_week)
                ws.cell(row=current_row, column=col_offset + 2, value=formatted_date)
                # Находим соответствующую базу для получения average_id
                cash_data = df[(df['OpenDate.Typed'] == date) & (df['Store.Name'] == cash_register)]
                base_name = cash_data.iloc[0]['BaseName'] if not cash_data.empty else None
                # План наполняемости (берем из данных или 0 если нет)
                plan_occupancy = cash_data.iloc[0].get('PlanOccupancy', 0) if not cash_data.empty else 0
                ws.cell(row=current_row, column=col_offset + 3, value=plan_occupancy)
                # Факт наполняемости (берем из average отчета)
                fact_occupancy = 0
                if base_name and base_name in average_data_dict:
                    avg_data = average_data_dict[base_name]
                    # Ищем данные за нужную дату
                    date_str = date.strftime('%Y-%m-%d')
                    for item in avg_data:
                        if item.get('OpenDate.Typed', '').startswith(date_str):
                            dish_amount = item.get('DishAmountInt', 0)
                            guest_num = item.get('GuestNum', 0)
                            if guest_num > 0:
                                fact_occupancy = round(dish_amount / guest_num, 2)
                            break
                ws.cell(row=current_row, column=col_offset + 4, value=fact_occupancy)
                # Процент выполнения — ВСЕГДА формула (даже если план = 0)
                plan_col_letter = get_column_letter(col_offset + 3)
                fact_col_letter = get_column_letter(col_offset + 4)
                formula = f"=IF({plan_col_letter}{current_row}=0, 0, {fact_col_letter}{current_row}/{plan_col_letter}{current_row})"
                ws.cell(row=current_row, column=col_offset + 5, value=formula)
                ws.cell(row=current_row, column=col_offset + 5).number_format = '0.00%'
                col_offset += len(headers)
            current_row += 1
        # Итоговая строка
        total_row = current_row
        ws.cell(row=total_row, column=1, value="ИТОГО")
        ws.cell(row=total_row, column=1).font = Font(bold=True)
        # Дата выгрузки в столбце "Дата" для строки ИТОГО
        current_date = datetime.now().strftime('%d.%m.%Y')
        ws.cell(row=total_row, column=2, value=current_date)
        ws.cell(row=total_row, column=2).font = Font(bold=True)
        col_offset = 0
        for cash_register in cash_registers:
            # Среднее плана (округленное до 2 знаков)
            plan_col = get_column_letter(col_offset + 3)
            ws.cell(row=total_row, column=col_offset + 3, value=f"=ROUND(AVERAGE({plan_col}4:{plan_col}{current_row-1}),2)")
            ws.cell(row=total_row, column=col_offset + 3).font = Font(bold=True)
            # Среднее факта (округленное до 2 знаков)
            fact_col = get_column_letter(col_offset + 4)
            ws.cell(row=total_row, column=col_offset + 4, value=f"=ROUND(AVERAGE({fact_col}4:{fact_col}{current_row-1}),2)")
            ws.cell(row=total_row, column=col_offset + 4).font = Font(bold=True)
            # Процент выполнения (факт/план для итоговой строки)
            plan_total_col = get_column_letter(col_offset + 3)
            fact_total_col = get_column_letter(col_offset + 4)
            formula = f"=IF({plan_total_col}{current_row}=0, 0, {fact_total_col}{current_row}/{plan_total_col}{current_row})"
            ws.cell(row=total_row, column=col_offset + 5, value=formula)
            ws.cell(row=total_row, column=col_offset + 5).number_format = '0.00%'
            ws.cell(row=total_row, column=col_offset + 5).font = Font(bold=True)
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
        # Условное форматирование для процентов выполнения
        col_offset = 0
        for cash_register in cash_registers:
            percent_col = col_offset + 5
            percent_letter = get_column_letter(percent_col)
            percent_range = f"{percent_letter}4:{percent_letter}{current_row-1}"
            # Зеленый для >= 100%
            rule_green = CellIsRule(
                operator='greaterThanOrEqual',
                formula=['1'],
                stopIfTrue=False,
                fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            )
            ws.conditional_formatting.add(percent_range, rule_green)
            # Красный для < 100%
            rule_red = CellIsRule(
                operator='lessThan',
                formula=['1'],
                stopIfTrue=False,
                fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            )
            ws.conditional_formatting.add(percent_range, rule_red)
            col_offset += len(headers)

    def create_dish_price_report_sheet(self, wb, df, cash_registers, bases_config, average_data_dict):
        """Создание листа План_факт стоимость блюд"""
        ws = wb.create_sheet(title="План_факт стоимость блюд")
        # Заголовок
        ws['A1'] = "ПЛАН-ФАКТ СТОИМОСТЬ БЛЮД"
        ws['A1'].font = Font(bold=True, size=16)
        ws.merge_cells('A1:E1')
        # Цвета для разных баз
        base_colors = {
            0: "E2EFDA",  # Зеленый
            1: "FFEB9C",  # Желтый
            2: "C6E0B4",  # Светло-зеленый
            3: "FFE699",  # Светло-желтый
            4: "A9D08E",  # Зеленый
            5: "F8CBAD",  # Оранжевый
            6: "9CC2E5",  # Голубой
            7: "D9E1F2",  # Светло-синий
        }
        # Заголовки для каждой кассы
        headers = ["День недели", "Дата", "План", "Факт", "% выполнения"]
        col_offset = 0
        base_index = 0
        for cash_register in cash_registers:
            start_col = col_offset + 1
            end_col = col_offset + len(headers)
            ws.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=end_col)
            ws.cell(row=2, column=start_col, value=cash_register)
            ws.cell(row=2, column=start_col).font = Font(bold=True)
            ws.cell(row=2, column=start_col).alignment = Alignment(horizontal='center')
            # Уникальный цвет для каждой базы
            color_index = base_index % len(base_colors)
            color = base_colors[color_index]
            ws.cell(row=2, column=start_col).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            # Заголовки столбцов
            for i, header in enumerate(headers):
                cell = ws.cell(row=3, column=col_offset + i + 1, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
            col_offset += len(headers)
            base_index += 1
        # Заполняем данные (стоимость блюд)
        dates = sorted(df['OpenDate.Typed'].unique())
        current_row = 4
        for date in dates:
            day_of_week = df[df['OpenDate.Typed'] == date]['DayOfWeekOpen'].iloc[0]
            formatted_date = df[df['OpenDate.Typed'] == date]['Date'].iloc[0]
            col_offset = 0
            for cash_register in cash_registers:
                ws.cell(row=current_row, column=col_offset + 1, value=day_of_week)
                ws.cell(row=current_row, column=col_offset + 2, value=formatted_date)
                # Находим соответствующую базу для получения average_id
                cash_data = df[(df['OpenDate.Typed'] == date) & (df['Store.Name'] == cash_register)]
                base_name = cash_data.iloc[0]['BaseName'] if not cash_data.empty else None
                # План стоимости блюд (берем из данных или 0 если нет)
                plan_dish_price = cash_data.iloc[0].get('PlanDishPrice', 0) if not cash_data.empty else 0
                ws.cell(row=current_row, column=col_offset + 3, value=plan_dish_price)
                # Факт стоимости блюд (берем из average отчета)
                fact_dish_price = 0
                if base_name and base_name in average_data_dict:
                    avg_data = average_data_dict[base_name]
                    # Ищем данные за нужную дату
                    date_str = date.strftime('%Y-%m-%d')
                    for item in avg_data:
                        if item.get('OpenDate.Typed', '').startswith(date_str):
                            dish_discount_sum = item.get('DishDiscountSumInt', 0)
                            dish_amount = item.get('DishAmountInt', 0)
                            if dish_amount > 0:
                                fact_dish_price = round(dish_discount_sum / dish_amount, 2)
                            break
                ws.cell(row=current_row, column=col_offset + 4, value=fact_dish_price)
                # Процент выполнения — ВСЕГДА формула (даже если план = 0)
                plan_col_letter = get_column_letter(col_offset + 3)
                fact_col_letter = get_column_letter(col_offset + 4)
                formula = f"=IF({plan_col_letter}{current_row}=0, 0, {fact_col_letter}{current_row}/{plan_col_letter}{current_row})"
                ws.cell(row=current_row, column=col_offset + 5, value=formula)
                ws.cell(row=current_row, column=col_offset + 5).number_format = '0.00%'
                col_offset += len(headers)
            current_row += 1
        # Итоговая строка
        total_row = current_row
        ws.cell(row=total_row, column=1, value="ИТОГО")
        ws.cell(row=total_row, column=1).font = Font(bold=True)
        # Дата выгрузки в столбце "Дата" для строки ИТОГО
        current_date = datetime.now().strftime('%d.%m.%Y')
        ws.cell(row=total_row, column=2, value=current_date)
        ws.cell(row=total_row, column=2).font = Font(bold=True)
        col_offset = 0
        for cash_register in cash_registers:
            # Среднее плана (округленное до 2 знаков)
            plan_col = get_column_letter(col_offset + 3)
            ws.cell(row=total_row, column=col_offset + 3, value=f"=ROUND(AVERAGE({plan_col}4:{plan_col}{current_row-1}),2)")
            ws.cell(row=total_row, column=col_offset + 3).font = Font(bold=True)
            # Среднее факта (округленное до 2 знаков)
            fact_col = get_column_letter(col_offset + 4)
            ws.cell(row=total_row, column=col_offset + 4, value=f"=ROUND(AVERAGE({fact_col}4:{fact_col}{current_row-1}),2)")
            ws.cell(row=total_row, column=col_offset + 4).font = Font(bold=True)
            # Процент выполнения (факт/план для итоговой строки)
            plan_total_col = get_column_letter(col_offset + 3)
            fact_total_col = get_column_letter(col_offset + 4)
            formula = f"=IF({plan_total_col}{current_row}=0, 0, {fact_total_col}{current_row}/{plan_total_col}{current_row})"
            ws.cell(row=total_row, column=col_offset + 5, value=formula)
            ws.cell(row=total_row, column=col_offset + 5).number_format = '0.00%'
            ws.cell(row=total_row, column=col_offset + 5).font = Font(bold=True)
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
        # Условное форматирование для процентов выполнения
        col_offset = 0
        for cash_register in cash_registers:
            percent_col = col_offset + 5
            percent_letter = get_column_letter(percent_col)
            percent_range = f"{percent_letter}4:{percent_letter}{current_row-1}"
            # Зеленый для >= 100%
            rule_green = CellIsRule(
                operator='greaterThanOrEqual',
                formula=['1'],
                stopIfTrue=False,
                fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            )
            ws.conditional_formatting.add(percent_range, rule_green)
            # Красный для < 100%
            rule_red = CellIsRule(
                operator='lessThan',
                formula=['1'],
                stopIfTrue=False,
                fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            )
            ws.conditional_formatting.add(percent_range, rule_red)
            col_offset += len(headers)

    def create_excel_report(self, all_data, bases_config, output_filename, start_date, end_date, average_data_dict, selected_sheets):
        """Создание Excel отчета с выбранными листами"""
        if not all_data:
            self.log_message("❌ Нет данных для создания отчета")
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
        # Создаем выбранные листы
        if 1 in selected_sheets:
            self.create_main_report_sheet(wb, df, cash_registers)
        if 2 in selected_sheets:
            self.create_occupancy_report_sheet(wb, df, cash_registers, bases_config, average_data_dict)
        if 3 in selected_sheets:
            self.create_dish_price_report_sheet(wb, df, cash_registers, bases_config, average_data_dict)
        # Сохраняем файл
        wb.save(output_filename)
        self.log_message(f"✅ Отчет сохранен: {output_filename}")
        return True