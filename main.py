"""
Скрипт преобразования данных из файла Excel в файл JSON.
"""
import json

from openpyxl import load_workbook

FILE_LIMITS = "limits.json"
FILE_INPUT = "input.xlsx"
FILE_OUTPUT = "output.json"
ERROR_NOT_FOUND_LIMITS_FILE = 2
ERROR_NOT_FOUND_INPUT_FILE = 3


def load_limits_file(filename=FILE_LIMITS):
    """Загрузка JSON-файла с настройками ограничений."""
    with open(filename) as json_file:
        result = json.load(json_file)
    return result


def load_excel_file(limits_dict, filename=FILE_INPUT):
    """
    Загрузка данных из Excel-файла.
    """
    wb = load_workbook(filename)
    ws = wb[limits_dict['sheet']]

    result = {}
    for row in range(limits_dict['begin_row'], limits_dict['end_row'] + 1):
        if row not in limits_dict['exclude_rows']:
            unique_key = str(ws[limits_dict['unique_key_column'] + str(row)].value)
            result[unique_key] = {}

            for column_name, column_value in limits_dict['columns'].items():
                if column_value != limits_dict['unique_key_column']:
                    result[unique_key][column_name] = ws[column_value + str(row)].value
    return result


def save_json_file(data_dict, filename=FILE_OUTPUT):
    with open(filename, 'w') as outfile:
        json.dump(data_dict, outfile, indent=4, ensure_ascii=False)


if __name__ == '__main__':
    # Чтение JSON-файла с настройками ограничений.
    limits = dict()
    try:
        limits = load_limits_file()
    except FileNotFoundError:
        print(f"Ошибка: не найден файл {FILE_LIMITS}")
        exit(ERROR_NOT_FOUND_LIMITS_FILE)

    # Чтение входного Excel-файла.
    data = dict()
    try:
        data = load_excel_file(limits)
    except FileNotFoundError:
        print(f"Ошибка: не найден файл {FILE_INPUT}")
        exit(ERROR_NOT_FOUND_INPUT_FILE)

    # Сохранение JSON-файла
    save_json_file(data)
