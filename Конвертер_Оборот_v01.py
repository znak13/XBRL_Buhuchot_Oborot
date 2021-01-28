import openpyxl
from module.Oborot import oborotka
from module.periods import Period
from module.functions import *
import builtins
import module.period_selection as selection
from module import logger

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

    # month = 3
    if month in [3, 6, 9, 12]:
        report_type = 'quarter'     # выбрана квартальная отчетность
    else:
        report_type = 'month'       # выбрана месячная отчетность

    # Расчет периодов в отчетности
    period = Period(year, month)
    # устанавливаем 'period' как глобальную переменную (включая модули)
    builtins.period = period

    # Создаем новый файл отчетности xbrl, создав копию шаблона
    if report_type == 'month':
        # шаблон для квартальной отчетности
        shablon = file_shablon_oborot_month
    else:
        # шаблон для месячной отчетности
        shablon = file_shablon_oborot_quarter



    # Создаем новый файл отчетности xbrl, создав копию шаблона
    shutil.copyfile(dir_shablon + shablon, full_fileNewName)
    log.info(f'создан файл: {full_fileNewName}')


    # Загружаем данные из этого файла xbrl
    wb_xbrl = openpyxl.load_workbook(filename=full_fileNewName)

    # Формируем формы отчетности
    oborotka(wb_xbrl, full_fileNewName, report_type)

    # Сохраняем в файл отчетности xbrl
    wb_xbrl.save(full_fileNewName)

    print('Вроде всё ОК!......')

    # ----------------------------------------------
