# import xlwings as xw
# import openpyxl
import builtins
from module.functions import *
from module import BuhOtch
from module import logger
from module.periods import Period
import module.period_selection as selection

from module.globals import *


# %%

if __name__ == "__main__":

    # Выбор периода отчетности
    year, quarter, dir_QuarterReports, fileNewName = selection.main()
    dir_QuarterReports += '/'
    full_fileNewName = dir_QuarterReports + fileNewName

    # Расчет периодов отчетности
    period = Period(year, quarter)
    # устанавливаем 'period' как глобальную переменную (включая модули)
    builtins.period = period

    # название файла - Шаблон
    if period.number:
        file_shablon = file_shablon_quarter
    else:
        file_shablon = file_shablon_year

    # Создаем новый файл отчетности xbrl, создав копию шаблона
    shutil.copyfile(dir_shablon + file_shablon, full_fileNewName)
    print(f'создан файл: {full_fileNewName}')

    # ....................................
    # Включаем логировние
    log = logger.create_log(path=dir_QuarterReports,
                     file_log=fileNewName + log_endName,
                     file_debug=fileNewName + debug_endName
                     )
    # устанавливаем 'log' как глобальную переменную (включая модули)
    builtins.log = log

    # Загружаем данные из нового файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=full_fileNewName)
    # Формируем формы Бух.отчетности
    BuhOtch.buhOtchot(wb, dir_QuarterReports, dir_shablon, period)

    # Сохраняем результат
    wb.save(full_fileNewName)
