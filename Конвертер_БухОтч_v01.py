# import xlwings as xw
# import openpyxl
import builtins
from module.functions import *
from module import BuhOtch
from module import logger
from module.periods import Period

from module.globals import *


# %%

if __name__ == "__main__":

    # Ввод периода отчетности
    period = Period()
    period.set()
    builtins.period = period

    # название файла - Шаблон
    if period.number:
        file_shablon = file_shablon_quarter
    else:
        file_shablon = file_shablon_year

    # Создаем новый файл отчетности на основе файла-шаблона
    file_new_name = load_xbrl(file_shablon, file_dir=dir_shablon)

    # путь к папке с файлами текущей отчетности
    dir_file_report = os.path.dirname(file_new_name) + '/'
    # имя файла без пути
    file_new_name_only = os.path.basename(file_new_name)

    # ....................................
    # Включаем логировние
    log = logger.create_log(path=dir_file_report,
                     file_log=file_new_name_only + log_endName,
                     file_debug=file_new_name_only + debug_endName
                     )
    # устанавливаем 'log' как глобальную переменную (включая модули)
    builtins.log = log

    # Загружаем данные из нового файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=file_new_name)
    # Формируем формы Бух.отчетности
    BuhOtch.buhOtchot(wb, dir_file_report, dir_shablon, period)

    # Сохраняем результат
    wb.save(file_new_name)
