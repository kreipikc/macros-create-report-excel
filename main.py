import sys
from datetime import date
from typing import Optional, List
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, numbers, Font, Alignment
from pydantic import ValidationError
from babel.dates import format_date
from config import POST_CELL, REPORT_MOTH_CELL, EMPLOYEE_CELL, START_ROW_READ, START_ROW_WRITE, COUNT_ROW_AFTER_CHECKS
from schemas import ChecksDefault, AdditionalInfo
from utils import create_check, sum_money_all_checks, convert_to_words, validate_check, get_absolute_path


def read_input_additional_info(file_path: str) -> Optional[AdditionalInfo]:
    """Read additional information from an Excel file.

    Args:
        file_path (str): The path to the Excel file containing additional information.

    Returns:
        Optional[AdditionalInfo]: An AdditionalInfo object containing employee, report month, and post data.

    Raises:
        ValidationError: If the data in the Excel file is invalid.
    """
    try:
        workbook = load_workbook(file_path, data_only=True)
        sheet = workbook.active

        data = AdditionalInfo(
            employee=sheet[EMPLOYEE_CELL].value,
            report_month=sheet[REPORT_MOTH_CELL].value,
            post=sheet[POST_CELL].value
        )

        return data
    except ValidationError as e:
        print(f"Validation error: {e}")
        return None


def read_input_checks(file_path: str) -> List[ChecksDefault]:
    """Read check data from an Excel file and validate it.

    Args:
        file_path (str): The path to the Excel file containing check data.

    Returns:
        List[ChecksDefault]: A list of ChecksDefault objects created from the Excel data.

    Raises:
        Exception: If any check data is invalid or missing required fields.
    """
    workbook = load_workbook(file_path, data_only=True)
    sheet = workbook.active

    data = []
    for row in sheet.iter_rows(min_row=START_ROW_READ, values_only=True):
        if all(cell is None for cell in row):
            break

        check = create_check(row)
        if check:
            data.append(check)

    validate_check(data)

    # for row in data:
    #     print(row)

    return data


def create_report(checks: List[ChecksDefault], info_data: AdditionalInfo, path_save: str) -> None:
    """Create a report by filling a template with check data and additional information.

    Args:
        checks (List[ChecksDefault]): A list of ChecksDefault objects to be included in the report.
        info_data (AdditionalInfo): Additional information to be included in the report.
        path_save (str): The path to save the report.

    Returns:
        None

    Raises:
        Exception: If there is an error while creating the report.
    """
    workbook = load_workbook(get_absolute_path("templates/template_advance_report.xlsx"))
    sheet = workbook.active
    locale_date = 'ru_RU'
    sys.stdout.reconfigure(encoding='utf-8')

    sum_checks = sum_money_all_checks(checks)
    rubles = int(sum_checks)
    kopecks = int((sum_checks - rubles) * 100)

    sheet['R9'] = rubles
    sheet['X9'] = kopecks
    sheet['J13'] = date.today().strftime('%d.%m.%Y')
    sheet['O15'] = format_date(date.today(), format='d MMMM yyyy г.', locale=locale_date)
    sheet['F21'] = info_data.employee
    sheet['Q23'] = f"Расходы {format_date(info_data.report_month, format='MMMM yyyy', locale=locale_date)}"
    sheet['J33'] = sum_checks
    sheet['J39'] = convert_to_words(rubles, kopecks)
    sheet['I55'] = info_data.employee
    sheet['D56'] = format_date(date.today(), format='d MMMM yyyy г.', locale=locale_date)
    sheet['K56'] = convert_to_words(rubles, kopecks)

    border = Border(
        left=Side(border_style='thin', color='000000'),
        right=Side(border_style='thin', color='000000'),
        top=Side(border_style='thin', color='000000'),
        bottom=Side(border_style='thin', color='000000')
    )

    for idx, check in enumerate(checks, start=START_ROW_WRITE):
        sheet.insert_rows(idx)
        sheet.row_dimensions[idx].height = 23

        sheet.merge_cells(f'B{idx}:C{idx}')
        sheet[f'B{idx}'].number_format = numbers.FORMAT_NUMBER
        sheet[f'B{idx}'] = check.number_str

        sheet.merge_cells(f'D{idx}:E{idx}')
        sheet[f'D{idx}'] = check.date.strftime('%d.%m.%Y') if check.date is not None else None
        sheet.column_dimensions['D'].width = 3.83 * 3

        sheet.merge_cells(f'F{idx}:G{idx}')
        sheet[f'F{idx}'] = check.id_check if check.id_check is not None else None
        sheet.column_dimensions['F'].width = 3.83 * 3

        sheet.merge_cells(f'H{idx}:K{idx}')
        sheet[f'H{idx}'] = check.type.value
        sheet.column_dimensions['H'].width = 3.83 * 3 + 10

        sheet.merge_cells(f'L{idx}:N{idx}')
        sheet[f'L{idx}'].number_format = numbers.FORMAT_NUMBER
        sheet[f'L{idx}'] = check.sum_check
        sheet.column_dimensions['L'].width = 3.83 * 3

        sheet.merge_cells(f'O{idx}:Q{idx}')
        sheet.column_dimensions['O'].width = 3.83 * 3

        sheet.merge_cells(f'R{idx}:T{idx}')
        sheet[f'R{idx}'].number_format = numbers.FORMAT_NUMBER
        sheet[f'R{idx}'] = check.sum_check
        sheet.column_dimensions['R'].width = 3.83 * 3

        sheet.merge_cells(f'U{idx}:W{idx}')
        sheet.column_dimensions['U'].width = 3.83 * 3

        sheet.merge_cells(f'X{idx}:Y{idx}')
        sheet.column_dimensions['X'].width = 3.83 * 3

        for row in sheet[f'A{idx}:Y{idx}']:
            for cell in row:
                cell.border = border


    # Заполнение "Итого" и данных на этой строке
    new_block_data_row = START_ROW_WRITE + len(checks)

    # Все строки ниже чеков с размером 11
    for i in range(COUNT_ROW_AFTER_CHECKS):
        sheet.row_dimensions[new_block_data_row + i].height = 11

    sheet.merge_cells(f'H{new_block_data_row}:K{new_block_data_row}')
    sheet[f'H{new_block_data_row}'] = "Итого"

    sheet.merge_cells(f'L{new_block_data_row}:N{new_block_data_row}')
    sheet[f'L{new_block_data_row}'] = f"=SUM(L{START_ROW_WRITE}:L{new_block_data_row - 1})"

    sheet.merge_cells(f'O{new_block_data_row}:Q{new_block_data_row}')

    sheet.merge_cells(f'R{new_block_data_row}:T{new_block_data_row}')
    sheet[f'R{new_block_data_row}'] = f"=SUM(R{START_ROW_WRITE}:R{new_block_data_row - 1})"

    sheet.merge_cells(f'U{new_block_data_row}:W{new_block_data_row}')

    sheet.merge_cells(f'X{new_block_data_row}:Y{new_block_data_row}')

    for row in sheet[f'L{new_block_data_row}:W{new_block_data_row}']:
        for cell in row:
            cell.border = border


    # Создание мест для подписей
    sheet.merge_cells(f'B{new_block_data_row + 2}:E{new_block_data_row + 2}')
    sheet[f'B{new_block_data_row + 2}'] = "Подотчетное лицо"

    sheet.merge_cells(f'F{new_block_data_row + 2}:L{new_block_data_row + 2}')
    for row in sheet[f'F{new_block_data_row + 2}:L{new_block_data_row + 2}']:
        for cell in row:
            cell.border = Border(bottom=Side(border_style='thin', color='000000'))

    sheet.merge_cells(f'F{new_block_data_row + 3}:L{new_block_data_row + 3}')
    sheet[f'F{new_block_data_row + 3}'] = "подпись"
    cell_signature = sheet[f'F{new_block_data_row + 3}']
    cell_signature.font = Font(name='Arial', size=6, bold=False, color='000000')
    cell_signature.alignment = Alignment(horizontal='center', vertical='center')

    sheet.merge_cells(f'N{new_block_data_row + 2}:Y{new_block_data_row + 2}')
    for row in sheet[f'N{new_block_data_row + 2}:Y{new_block_data_row + 2}']:
        for cell in row:
            cell.border = Border(bottom=Side(border_style='thin', color='000000'))

    sheet.merge_cells(f'N{new_block_data_row + 3}:Y{new_block_data_row + 3}')
    sheet[f'N{new_block_data_row + 3}'] = "расшифровка подписи"
    cell_decrypt_signature = sheet[f'N{new_block_data_row + 3}']
    cell_decrypt_signature.font = Font(name='Arial', size=6, bold=False, color='000000')
    cell_decrypt_signature.alignment = Alignment(horizontal='center', vertical='center')

    sheet.merge_cells(f'B{new_block_data_row + 6}:E{new_block_data_row + 6}')
    sheet[f'B{new_block_data_row + 6}'] = "Руководитель"

    sheet.merge_cells(f'F{new_block_data_row + 6}:L{new_block_data_row + 6}')
    for row in sheet[f'F{new_block_data_row + 6}:L{new_block_data_row + 6}']:
        for cell in row:
            cell.border = Border(bottom=Side(border_style='thin', color='000000'))

    sheet.merge_cells(f'N{new_block_data_row + 6}:Y{new_block_data_row + 6}')
    for row in sheet[f'N{new_block_data_row + 6}:Y{new_block_data_row + 6}']:
        for cell in row:
            cell.border = Border(bottom=Side(border_style='thin', color='000000'))

    sheet[f'N{new_block_data_row + 2}'] = info_data.employee
    sheet[f'N{new_block_data_row + 6}'] = info_data.employee

    workbook.save(path_save)


def main(path_input: str, path_save: str) -> None:
    """Main function to process an Excel file and generate a report.

    Args:
        path_input (str): The file path to the Excel file containing check data and additional information.
        path_save (str): The path to save the report.

    Returns:
        None
    """
    print("Старт сканирования данных...")
    checks_all = read_input_checks(path_input)
    info = read_input_additional_info(path_input)
    print("Сканирование завершено!")
    print("Старт создания отчета...")
    create_report(checks_all, info, path_save)
    print("Создание отчета завершено!")
    print("Можете закрывать консоль.")


if __name__ == "__main__":
    print("Запуск скрипта...")
    if len(sys.argv) > 2:
        file_path_input = sys.argv[1]
        file_path_save = sys.argv[2]
        main(file_path_input, file_path_save)
    else:
        print("Ошибка. Не переданы пути для работы скрипта.")