import os
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import openpyxl

def select_excel_file():
    Tk().withdraw()
    filename = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return filename

def find_txp_files(directory):
    return [f for f in os.listdir(directory) if f.endswith('.txp')]

def extract_numbers(txp_files):
    pattern = r"x=(\d+\.?\d*)\.txp"
    numbers = [re.search(pattern, file).group(1) for file in txp_files if re.search(pattern, file)]
    # 文字列から実数へ変換し、数値としてソート
    sorted_numbers = sorted([float(number) for number in numbers])
    return sorted_numbers

def write_numbers_to_excel(directory, excel_file):
    txp_files = find_txp_files(directory)
    numbers = extract_numbers(txp_files)

    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active

    for i, number in enumerate(numbers, start=2):
        ws[f'A{i}'] = number

    wb.save(excel_file)

# 実行
excel_file = select_excel_file()
if excel_file:
    directory = os.path.dirname(excel_file)
    write_numbers_to_excel(directory, excel_file)
