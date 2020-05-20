import openpyxl
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.styles import Alignment, PatternFill, Font
from datetime import date
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

def load_xbrl(file_shablon: str, file_dir=r'./Шаблоны/'):
    """ Загрузка данных из шаблона в новый файл xbrl"""
    # file_shablon - имя файла-шиблона

    # название нового файла-отчетности xbrl
    print(f'Создание файла отчетности....')
    file_new_name = asksaveasfilename(filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
    # Добавляем расширение файла
    file_new_name = file_new_name + '.xlsx'
    # # запоминаем путь к файлу
    # dir_name = os.path.dirname(file_new_name)
    # # отбрасываем путь к файлу
    # file_new = os.path.basename(file_new_name)

    # --------------------------------------------
    # Создаем новый файл отчетности xbrl, создав копию шаблона
    shutil.copyfile(file_dir + file_shablon, file_new_name)
    print(f'создан файл: {file_new_name}')

    # Загружаем данные из файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=file_new_name)

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
def Codesofsheets(wb_xbrl):
    """ Список кодов и наименования листов """
    codes = {}
    for sheet in wb_xbrl.sheetnames:
        ws_xbrl = wb_xbrl[sheet]
        ws_xbrl_cell = ws_xbrl.cell(3, 1)
        codes[ws_xbrl_cell.value] = sheet
        # print(sheet, ws_xbrl_cell.value)

    # исключаем из списка вкладку '_dropDownSheet'
    for code in codes:
        if codes[code] == '_dropDownSheet':
            codes.pop(code)
            break

    return codes


# %%

def delNullSheets(wb_xbrl, df_matrica, sheetsCodes, codesNull):
    """Удаляем незаполненные вкладки"""

    # sheetsCodes.pop('Добавочный капитал ')  # это не ошибка!

    for code in sheetsCodes:
        if code not in df_matrica["URL"].values:
            sheetName = sheetsCodes[code]
            wb_xbrl.remove(wb_xbrl[sheetName])

    for code in codesNull:
        sheetName = sheetsCodes[code]
        wb_xbrl.remove(wb_xbrl[sheetName])


# %%
def addPeriod(df_matrica, wb_xbrl):
    """Добавляем в формы периоды"""

    # Коды вкладок
    sheetsCodesID = Codesofsheets(wb_xbrl)

    for code in sheetsCodesID:
        sheetName = sheetsCodesID[code]
        per1_cell = df_matrica[df_matrica["URL"] == code]["per1_cell"][0]
        per2_cell = df_matrica[df_matrica["URL"] == code]["per2_cell"][0]
        period1 = df_matrica[df_matrica["URL"] == code]["period1"][0]
        period2 = df_matrica[df_matrica["URL"] == code]["period2"][0]
        ws = wb_xbrl[sheetName]
        if per1_cell != '-':
            ws[per1_cell].value = period1
        if per2_cell != '-':
            ws[per2_cell].value = period2

    pass


# %%
def addInfoSheets(wb_xbrl):
    """ Добавляем формы с общими данными"""

    sheets = wb_xbrl._sheets
    # Загружаем данные из файла таблицы 'fileInfo'
    wb_fileInfo = loadInfoSheets()
    for sheet in wb_fileInfo._sheets:
        sheets.append(sheet)

    # или можно так вставить:
    # lenSheets = len(sheets)
    # sheets.insert(lenSheets, ws1_fileInfo)
    # sheets.insert(lenSheets+1, ws2_fileInfo)


def loadInfoSheets():
    """ Загружаем данные из файла с общими сведениями"""
    file_dir = r'./Шаблоны/'
    fileInfo = file_dir + "Шаблон_БухОтч_3_2_Общие сведения.xlsx"
    # Загружаем данные из файла таблицы 'fileInfo'
    wb_fileInfo = openpyxl.load_workbook(filename=fileInfo)

    sheetsIgnor = ['_dropDownSheet', "wqeqwqw"]  # список листов, которые добавлять не нужно
    # исключаем ненужные листы
    for sheet in sheetsIgnor:
        try:
            wb_fileInfo.remove(wb_fileInfo[sheet])
        except KeyError:
            print(f'ВНИМАНИЕ! В файле "{os.path.basename(fileInfo)}" отсутствует лист "{sheet}"')

    return wb_fileInfo


def correctStyle(fileName):
    """ Исправляем форматы ячеек в формах с общими данными"""
    # (исправить до записи в файл не получается: сохраняется некорректно)

    # Загружаем данные из файла отчетности
    wb_xbrl = openpyxl.load_workbook(filename=fileName)
    # Загружаем данные из файла таблицы 'fileInfo'
    wb_fileInfo = loadInfoSheets()

    for sheet in wb_fileInfo._sheets:
        ws_xbrl = wb_xbrl[sheet.title]
        # print('===>', sheet.title)
        for row in sheet.rows:
            for cell in row:
                rgb = cell.fill.fgColor.rgb
                # без фона (белый фон)
                color_0 = openpyxl.styles.colors.Color(rgb='00000000', indexed=None, auto=None,
                                                       theme=None, tint=0.0, type='rgb')
                # серый фон
                color_1 = openpyxl.styles.colors.Color(rgb=None, indexed=22, auto=None,
                                                       theme=None, tint=0.0, type='indexed')
                # перекрашиваем ячейки
                if rgb == '00000000':
                    ws_xbrl[cell.coordinate].fill = PatternFill(patternType=None, fgColor=color_0)
                else:
                    ws_xbrl[cell.coordinate].fill = PatternFill(patternType='solid', fgColor=color_1)

                # Копируем цвет шрифта
                if cell.font.color != None:  # проверяем установлен ли цвет шрифта
                    fontColor = cell.font.color.rgb
                    # # красный цвет
                    # color_font = openpyxl.styles.colors.Color(rgb='FFFF0000')
                    if type(fontColor) == str:
                        color_font = openpyxl.styles.colors.Color(rgb=fontColor)
                        ws_xbrl[cell.coordinate].font = Font(color=color_font)

    wb_xbrl.save(fileName)


# %%
def inputPeriod():
    """ Ввод периода отчетности """
    reports = {'0': 'годовая',
               '1': '1-ый квартал',
               '2': '2-ой квартал',
               '3': '3-ий квартал',
               '4': '4-ый квартал'}
    period = -1
    attempt = 1
    while period not in reports:
        period = input(f"Период отчетности:\n"
                       f"0 ==> {reports['0']}\n"
                       f"1 ==> {reports['1']}\n"
                       f"2 ==> {reports['2']}\n"
                       f"3 ==> {reports['3']}\n"
                       f"4 ==> {reports['4']}\n")
        if period not in reports:
            print(f'Выбор "{period}" - не верный. Попоробуйте ещё.\n')
            attempt += 1
            print(f'(попытка № {attempt})')

        else:
            print(f'Выбрана отчетность: {reports[period]}')

    return int(period)

# %%
def datesInPeriods(nomberOfPeriod):
    """ Даты в периодах"""
    year = 2020
    periods = {0:[date(year, 1,1), date(year,12,31)],
               1:[date(year, 1,1), date(year, 3,31)],
               2:[date(year, 4,1), date(year, 6,30)],
               3:[date(year, 7,1), date(year, 9,30)],
               4:[date(year,10,1), date(year,12,31)]
               }
    # print(periods[2][0].strftime("%Y-%m-%d"))
    return periods[nomberOfPeriod]

# %%

def insertPeriodInFile(period):
    """ Вставляем даты выбранного периода в файл"""

    # Загружаем данные из файла с периодами"""
    file_dir = r'./Шаблоны/'
    fileName = file_dir + "Периоды.xlsx"
    wb = openpyxl.load_workbook(filename=fileName)
    ws = wb['Периоды']
    # Вписываем даты начали и конци периода
    ws['B3'] = period[0].strftime("%Y-%m-%d")
    ws['B4'] = period[1].strftime("%Y-%m-%d")

    wb.save(fileName)

# %%
if __name__ == "__main__":
    period = inputPeriod()
    dates = datesInPeriods(period)
    insertPeriodInFile(dates)


