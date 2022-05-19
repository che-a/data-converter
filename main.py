!#/usr/bin/env python3

"""
Скрипт преобразования всех или части данных из Excel-таблицы в JSON-файл.
"""
import json
from pprint import pprint

from openpyxl import load_workbook


def load_limits_file(filename='limits.json'):
    """
    Загрузка файла ограничений столбцов и строк
    """
    with open(filename) as json_file:
        data = json.load(json_file)
    return data


def load_excel_file(filename='input.xlsx'):
    """
    Загрузка данных из Excel-файла
    """
    limits = load_limits_file()
    wb = load_workbook(filename)
    ws = wb[limits['sheet']]

    data = {}
    for row in range(limits['begin_row'], limits['end_row'] + 1):
        if row not in limits['exclude_rows']:
            unique_key = ws[limits['unique_key_column'] + str(row)].value
            data[unique_key] = {}

            for column_name, column_value in limits['columns'].items():
                if column_value != limits['unique_key_column']:
                    data[unique_key][column_name] = ws[column_value + str(row)].value

    return data


if __name__ == '__main__':
    pprint(load_excel_file())
