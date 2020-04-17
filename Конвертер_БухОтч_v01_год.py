from module.analiz_data import *
from module.functions import *


# %%

def convert_balans():
    """ Конвертер данных БухОтчетности"""

    file_matrica = 'Матрица_3_2_год.xlsx'        # название файла - Матрица
    sheet_name   = 'БухОтч'                         # имя вкладки в Матрице
    file_shablon = 'Шаблон_БухОтч_3_2_год.xlsx'  # название файла - Шаблон

    # Загружаем данные из Матрицы
    df_matrica = load_matrica(file_matrica, sheet_name)
    # Создаем новый файл отчетности на основе файла-шаблона и загружаем из него данные
    wb_xbrl, file_new_name = load_xbrl(file_shablon)
    #
    # перебираем все вкладки в созданном файле отчетности
    for sheet in df_matrica.index.values.tolist():
        print(f'...загружаем форму: "{sheet}"')

        # загружаем данные из нужного файла бух.отчетности
        # название файла отчетности из Матрицы
        file_name = str(df_matrica.loc[sheet, 'file'])
        file_dir = r'./Отчетность/БухОтч/'
        file_report = file_dir + file_name
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

        # загружаем нужную вкладку из файла отчетности
        ws_xbrl = wb_xbrl[sheet]

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

    # Сохраняем в файл отчетности xbrl
    wb_xbrl.save(file_new_name)


# %%

if __name__ == "__main__":
    convert_balans()

    # Записываем ошибки
    write_errors()

    print('......!ОК!......')
