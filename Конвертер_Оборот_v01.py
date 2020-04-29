
from module.analiz_data import *
from module.functions import *


# %%

def compare_str(string):
    """ Сравниваем наименование показателей"""

    # если строка начинается с ...
    if string.startswith('Итого'):
        # если строка содержит...
        if string.find('актив') > 0 and string.find('Балансовые счета') > 0:
            return True, 'актив', 'Балансовые счета'
        # иначе, если строка содержит...
        elif string.find('пассив') > 0 and string.find('Балансовые счета') > 0:
            return True, 'пассив', 'Балансовые счета'

    return False


# %%

def convert_oborot():
    """ Конвертер данных по Оборотке"""

    print()
    print(f'ВНИМАНИЕ!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n'
          f'Не забыть поменять периоды в Шаблоне! ')
    print()

    file_matrica = 'Матрица_3_1.xlsx'           # название файла - Матрица
    sheet_name   = 'Оборот'                     # имя вкладки в Матрице
    file_shablon = 'Шаблон_Оборот_3_1_(квартал).xlsx'   # название файла - Шаблон
    # file_shablon = 'Шаблон_Оборот_3_1_(месяц).xlsx'       # название файла - Шаблон

    # Загружаем данные из Матрицы
    df_matrica = load_matrica(file_matrica, sheet_name)
    # Создаем новый файл отчетности на основе файла-шаблона и загружаем из него данные
    wb_xbrl, file_new_name = load_xbrl(file_shablon)

    # перебираем все вкладки в созданном файле отчетности
    # (названия файлов отчетности указано в Матрице)
    for sheet in df_matrica.index.values.tolist():
        print(f'...загружаем форму: "{sheet}"')

        # загружаем данные из нужного файла бух.отчетности
        # название файла отчетности из Матрицы
        file_name = str(df_matrica.loc[sheet, 'file'])
        file_dir = r'./Отчетность/Оборотка/'
        file_report = file_dir + file_name
        # загрузка данныз из файла 'report'
        df_report = load_report(file_report)

        # 'report': текст в заголовке столбца с данными в таблице
        string_begin = str(df_matrica.loc[sheet, 'string'])
        # 'report': номер первой строки в таблице
        begin_row_df_report = find_row(df_report, string_begin, string_col=2) + 1
        # 'report': номер последней строки в таблице
        end_row_df_report = len(df_report.index)

        # wb_xbrl: координаты верхней левой ячейки с данными
        begin_row_wb_xbrl, begin_col_wb_xbrl = coordinate(df_matrica.loc[sheet, 'xbrl_begin'])
        # wb_xbrl: начальная ячейка с показателем
        start_row = begin_row_wb_xbrl
        start_col = begin_col_wb_xbrl - 1

        # wb_xbrl: загружаем нужную вкладку из файла отчетности
        ws_xbrl = wb_xbrl[sheet]

        # перебор строк в файле 'report'
        for row_report in range(begin_row_df_report, end_row_df_report + 1):
            indicator = str(df_report.loc[row_report, 2]).replace('.', '')

            # пребор строк в xbrl файле
            for row_xbrl in range(start_row, ws_xbrl.max_row + 1):
                cell = ws_xbrl.cell(row_xbrl, start_col).value

                # проверяем совпадение показателей
                compare = compare_str(indicator)
                if cell.find(indicator) >= 0 or \
                        (compare and cell.find(compare[1]) > 0 and cell.find(compare[2]) > 0):
                    # print(f'{indicator} ==> {cell} ')

                    # кол-во колонок в таблице xbrl
                    col_total = df_matrica.loc[sheet, 'col_total']
                    # перебор колонок в таблице xbrl
                    for i in range(int(col_total)):
                        col = 'col_' + str(i + 1)
                        # преобразуем данные
                        data_reropt = analiz_data_all(
                            df_report.loc[row_report, (df_matrica.loc[sheet, col])]
                        )
                        # копируем данные
                        if data_reropt != '0.00':  # исключаем нулевые значения
                            ws_xbrl_cell = ws_xbrl.cell(row_xbrl, begin_col_wb_xbrl + i)
                            ws_xbrl_cell.value = data_reropt
                            # Форматируем ячейку
                            ws_xbrl_cell.alignment = Alignment(horizontal='right')
                    break
        print(f'.....готово')
    # Сохраняем в файл отчетности xbrl
    wb_xbrl.save(file_new_name)


# %%

if __name__ == "__main__":
    convert_oborot()

    # Записываем ошибки
    write_errors()

    print('......!ОК!......')

