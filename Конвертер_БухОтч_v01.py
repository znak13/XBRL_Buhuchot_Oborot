from module.analiz_data import *
from module.functions import *


# %%

def convert_balans():
    """ Конвертер данных БухОтчетности"""

    # название файла - Матрица
    # file_matrica = 'Матрица_1_3_1_long.xlsx'
    file_matrica = 'Матрица_3_2_год.xlsx'

    sheet_name = 'БухОтч'   # имя вкладки в Матрице

    # название файла - Шаблон
    # file_shablon = 'Шаблон_БухОтч_1_3_1_long.xlsx'
    file_shablon = 'Шаблон_БухОтч_3_2_год.xlsx'

    # Загружаем данные из Матрицы
    df_matrica = load_matrica(file_matrica, sheet_name)
    # Создаем новый файл отчетности на основе файла-шаблона и загружаем из него данные
    wb_xbrl, file_new_name = load_xbrl(file_shablon)
    # Коды вкладок
    sheetsCodes = Codesofsheets(wb_xbrl)

    codesNull =[] # коды вкладок, для которых отсутствуют файлы отчетности
    # перебираем все вкладки в созданном файле отчетности
    for sheet in df_matrica.index.values.tolist():
        print(f'...загружаем форму: "{sheet}"')

        # загружаем данные из нужного файла бух.отчетности
        # название файла отчетности из Матрицы
        file_name = str(df_matrica.loc[sheet, 'file'])
        file_dir = r'./Отчетность/БухОтч/'
        file_report = file_dir + file_name

        # Проверяем есть ли файл.
        if not os.path.exists(file_report):
            # Если файла нет, то выбираем следующую вкладку
            codesNull.append(df_matrica.loc[sheet, 'URL'])
            continue

        # загрузка данныз из файла 'report'
        df_report = load_report(file_report)
        # print(file)

        # находим номера первой и последней строк в таблице с данными
        string_begin = str(df_matrica.loc[sheet, 'string'])
        string_end = str(df_matrica.loc[sheet, 'end'])

        begin_row_df_report = find_row(df_report, string_begin) + 1
        end_row_df_report = find_row(df_report, string_end)

        # df_report
        begin_col_df_report = df_matrica.loc[sheet, 'begin_col']
        end_col_df_report = df_matrica.loc[sheet, 'end_col']

        # wb_xbrl: координаты верхней левой ячейки с данными
        begin_row_wb_xbrl, begin_col_wb_xbrl = coordinate(df_matrica.loc[sheet, 'xbrl_begin'])

        # количество строк и столбцов для копирования
        row_range = end_row_df_report - begin_row_df_report + 1
        col_range = end_col_df_report - begin_col_df_report + 1

        #==============================================================
        # Реальное имя вкладки
        code = df_matrica.loc[sheet, 'URL']
        sheetName = sheetsCodes[code]
        #==============================================================

        # загружаем нужную вкладку из файла отчетности
        ws_xbrl = wb_xbrl[sheetName]

        for row in range(row_range):
            for col in range(col_range):
                data_report = analiz_data_all(
                    df_report.loc[begin_row_df_report + row, begin_col_df_report + col]
                )
                # копируем данные
                if data_report != "0.00" and data_report != "Х" and \
                        data_report != "nan" and data_report != "0":
                    ws_xbrl_cell = ws_xbrl.cell(begin_row_wb_xbrl + row, begin_col_wb_xbrl + col)
                    ws_xbrl_cell.value = data_report
                    # Форматируем ячейку
                    ws_xbrl_cell.alignment = Alignment(horizontal='right')

    # Удаляем незаполненные вкладки
    delNullSheets(wb_xbrl, df_matrica, sheetsCodes, codesNull)
    # Сохраняем в файл отчетности xbrl
    wb_xbrl.save(file_new_name)


# %%

if __name__ == "__main__":
    convert_balans()

    # Записываем ошибки
    write_errors()

    print('......!ОК!......')
