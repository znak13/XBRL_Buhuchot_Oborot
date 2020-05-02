import openpyxl
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Alignment, PatternFill, Font

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
    codes = {}
    for sheet in wb_xbrl.sheetnames:
        ws_xbrl = wb_xbrl[sheet]
        ws_xbrl_cell = ws_xbrl.cell(3, 1)
        codes[ws_xbrl_cell.value] = sheet
        # print(sheet, ws_xbrl_cell.value)
    return codes


# %%

def delNullSheets(wb_xbrl, df_matrica, sheetsCodes, codesNull):
    """Удаляем незаполненные вкладки"""

    # удаляем из списка вкладку '_dropDownSheet'
    sheetsCodes.pop('Добавочный капитал ')

    for code in sheetsCodes:
        if code not in df_matrica["URL"].values:
            sheetName = sheetsCodes[code]
            wb_xbrl.remove(wb_xbrl[sheetName])

    for code in codesNull:
        sheetName = sheetsCodes[code]
        wb_xbrl.remove(wb_xbrl[sheetName])


# %%
def addPeriod(df_matrica, wb_xbrl, sheetsCodes):
    """Добавляем в формы периоды"""

    for code in df_matrica["URL"].values:
        sheetName = sheetsCodes[code]
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
    # Загружаем данные из файла таблицы 'fileInfo'
    file_dir = r'./Шаблоны/'
    fileInfo = file_dir + "Шаблон_БухОтч_3_2_Общие сведения.xlsx"
    # Загружаем данные из файла таблицы 'fileInfo'
    wb_fileInfo = openpyxl.load_workbook(filename=fileInfo)

    sheetsIgnor = ['_dropDownSheet']
    for sheet in sheetsIgnor:
        wb_fileInfo.remove(wb_fileInfo[sheet])

    return wb_fileInfo


def correctStyle(fileName):
    """ Исправляем форматы ячеек в формах с общими данными"""
    # (исправить до записи в файл не получается:
    # корректировка сохраняется некорректно)

    # Загружаем данные из файла отчетности
    wb_xbrl = openpyxl.load_workbook(filename=fileName)
    # Загружаем данные из файла таблицы 'fileInfo'
    wb_fileInfo = loadInfoSheets()

    for sheet in wb_fileInfo._sheets:
        ws_xbrl = wb_xbrl[sheet.title]
        print('===>', sheet.title)
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
                if cell.font.color != None: # проверяем установлен ли цвет шрифта
                    fontColor = cell.font.color.rgb
                    # # красный цвет
                    # color_font = openpyxl.styles.colors.Color(rgb='FFFF0000')
                    if type(fontColor) == str:
                        color_font = openpyxl.styles.colors.Color(rgb=fontColor)
                        ws_xbrl[cell.coordinate].font = Font(color=color_font)

    wb_xbrl.save(fileName)




# %%
if __name__ == "__main__":
    """
    dirName = r'd:\Clouds\YandexDisk\Git\XBRL_Buhuchot_Oborot\Шаблоны'
    file_shablon = dirName + '\\' + 'Шаблон_БухОтч_3_2_год.xlsx'
    wb_xbrl = openpyxl.load_workbook(filename=file_shablon)

    # sheetCode = Codesofsheets(wb_xbrl)
    # file_matrica = 'Матрица_3_2_год.xlsx'
    # sheet_name = 'БухОтч'
    # df_matrica = load_matrica(file_matrica, sheet_name, file_dir=dirName + "\\")
    #
    # for sheet in df_matrica.index.values.tolist():
    #     code = df_matrica.loc[sheet, 'URL']
    #     sheetName = sheetCode[code]

    file_dir = r'../Шаблоны/'
    fileInfo = file_dir + "Шаблон_БухОтч_3_2_Общие сведения.xlsx"
    # fileInfo = "Шаблон_БухОтч_3_2_Общие сведения.xlsx"
    # Загружаем данные из файла таблицы 'fileInfo'
    wb_fileInfo = openpyxl.load_workbook(filename=fileInfo)
    ws1_fileInfo = wb_fileInfo['Информация об отчитывающейся ор']
    ws2_fileInfo = wb_fileInfo['Основная деятельность некредитн']

    sheets = wb_xbrl._sheets
    sheets.append(ws1_fileInfo)
    sheets.append(ws2_fileInfo)

    #------------------------------------------------------------
    # Копирование формата ячеек
    from copy import copy
    ws2 = wb_xbrl['Информация об отчитывающейся ор']
    for row_1,row_2 in zip (ws1_fileInfo.rows, ws2.rows):
        for cell_1,cell_2 in zip(row_1, row_2):
            # cell_2.fill.start_color.index = "22"
            # cell_2.fill = PatternFill("solid")
            # if cell_1.fill.patternType == 'solid':
            #     print(cell_1.coordinate, cell_1.fill)
            # cell_1.fill.patternType = 'None'
            print(cell_1.coordinate, cell_1.fill)
            # pass
    #------------------------------------------------------------
    """
    dirName = r'd:\Clouds\YandexDisk\Git\XBRL_Buhuchot_Oborot\Отчетность\БухОтч'
    file = dirName + '\\' + '555.xlsx'
    wb_xbrl = openpyxl.load_workbook(filename=file)
    ws2 = wb_xbrl['Информация об отчитывающейся ор']
    my_red = openpyxl.styles.colors.Color(rgb='00FF0000')
    for row in ws2.rows:
        for cell in row:
            # print(f'{cell.coordinate} ==> {cell.fill}')
            # cell.fill = PatternFill (fgColor=my_red)
            # print(f'{cell.coordinate} ==> {cell.fill}')
            # input()
            # print(f'{cell.coordinate} ==> {ws2[cell.coordinate].value}')
            print("====> ", cell.coordinate, ws2[cell.coordinate].fill)
