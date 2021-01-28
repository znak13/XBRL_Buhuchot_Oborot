# import xlwings as xw
import openpyxl
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
    year, month, dir_QuarterReports, fileNewName = selection.main()
    dir_QuarterReports += '/'
    full_fileNewName = dir_QuarterReports + fileNewName

    # ....................................
    # Включаем логировние
    log = logger.create_log(path=dir_QuarterReports,
                            file_log=fileNewName + log_endName,
                            file_debug=fileNewName + debug_endName
                            )
    # устанавливаем 'log' как глобальную переменную (включая модули)
    builtins.log = log

    # Нужно проверить вобрана ли именно квартальная отчетность!!!

    # Расчет необходимых дат и периодов отчетности
    period = Period(year, month)
    # устанавливаем 'period' как глобальную переменную (включая модули)
    builtins.period = period

    # название файла - Шаблон
    if period.month in [3, 6, 9]:

        # переменная введена дополнительно,
        # т.к. она используется в дальнейших расчетах
        period.number = True  # означает квартальную отчетность

        file_shablon = file_shablon_quarter
    elif period.month == 0:

        # переменная введена дополнительно,
        # т.к. она используется в дальнейших расчетах
        period.number = False  # означает годовую отчетность

        file_shablon = file_shablon_year
    else:
        log.error(f'Ошибка в выборе периода! Отчет не сформирован! \n'
                  f'(Проверьте корректность указания периода в случае формирования бух.отчетности:\n'
                  f' - для годовой бух.отчетности необходимо выбрать первый пункт: "ГОДОВАЯ отчетность",\n'
                  f' - бух.отчетность может быть только квартальная(!).')
        sys.exit()

    # Создаем новый файл отчетности xbrl, создав копию шаблона
    shutil.copyfile(dir_shablon + file_shablon, full_fileNewName)
    # print(f'создан файл: {full_fileNewName}')
    log.info(f'создан файл: {full_fileNewName}')

    # Загружаем данные из нового файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=full_fileNewName)

    # Формируем формы Бух.отчетности
    BuhOtch.buhOtchot(wb, dir_QuarterReports, period)

    # Сохраняем результат
    wb.save(full_fileNewName)

    print('Вроде всё ОК!......')