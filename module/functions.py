import openpyxl
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment

import pandas as pd

from tkinter.filedialog import askopenfilename, asksaveasfilename

import shutil
import sys
import os

from tkinter import Tk
Tk().withdraw()


# %%

# файл с ошибками
errors_file = 'errors.txt'
ERRORS = []


def write_errors():
    """ Записываем ошибки в файл и открываем этот файл в Блокноте"""
    if not ERRORS:
        ERRORS.append('Ошибок не выявлено!')

    with open(errors_file, "w") as file:
        for k in ERRORS:
            file.write(str(k) + '\n\n')

    # Открываем файл ошибок в Блокноте
    notepad = r'%windir%\system32\notepad.exe'
    file = notepad + ' ' + os.path.abspath(errors_file)
    os.system(file)


# %%
def coordinate(cell):
    """Конвртер координат: A10 ==> 10, 1 """
    data = coordinate_from_string(cell)
    row = data[1]
    col = column_index_from_string(data[0])
    return row, col


# %%
def load_matrica(matrica, sheet_name, index_col=1, file_dir=r'./Шаблоны/'):
    """ загружвем данные из матрицы """
    # file_matrica = r'./Шаблоны/Матрица_3_1.xlsx'
    matrica = file_dir + matrica
    df_matrica = pd.read_excel(matrica, sheet_name=sheet_name, index_col=index_col)

    return df_matrica


# %%
def load_report(file_name, sheet_name='TDSheet'):
    """ загружвем данные из таблиц БухОтчетности """
    df = pd.read_excel(file_name, sheet_name=sheet_name, header=None)
    # устанавливаем начальный индекс не c 0, а c 1
    df.index += 1
    df.columns += 1
    return df


# %%

def load_xbrl(file_shablon: str, file_dir = r'./Шаблоны/'):
    """ Загрузка данных из шаблона в новый файл xbrl"""
    # file_shablon - имя файла-шиблона

    # название нового файла-отчетности xbrl
    print(f'Создание файла отчетности....')
    file_new_name = asksaveasfilename(filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
    # Добавляем расширение файла
    file_new_name = file_new_name + '.xlsx'
    # запоминаем путь к файлу
    dir_name = os.path.dirname(file_new_name)
    # отбрасываем путь к файлу
    file_new = os.path.basename(file_new_name)

    # --------------------------------------------
    # Создаем новый файл отчетности xbrl, создав копию шаблона
    shutil.copyfile(file_dir + file_shablon, r'./' + file_new)
    print(f'создан файл: {file_new_name}')

    # Загружаем данные из файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=file_new)

    return wb, file_new_name


# %%

def find_row(df, string, string_col=1):
    """найти в таблице номер строки, содержащей 'string' """

    # количество строк в df
    index_max = df.shape[0]  # или так: list(df.index.values)

    for row in range(1, index_max + 1):
        title = str(df.loc[row, string_col])

        if title == str(string):  # title.startswith(str(string))
            return row

    print('.......ERROR!.......')
    ERRORS.append(f'Раздел: "{string}" в файле не найден')
    write_errors()
    sys.exit("Ошибка!")

# %%
