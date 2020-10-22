import openpyxl
from module.Oborot import oborotka
from module.periods import Period_2
from module.functions import *
import builtins
import module.period_selection_2 as selection

# %%


if __name__ == "__main__":

    # Выбор периода отчетности
    year, month, dir_QuarterReports, fileNewName = selection.main()
    dir_QuarterReports += '/'
    full_fileNewName = dir_QuarterReports + fileNewName

    # month = 3
    if month in [3, 6, 9, 12]:
        report_type = 'quarter'     # выбрана квартальная отчетность
    else:
        report_type = 'month'       # выбрана месячная отчетность

    # Расчет периодов в отчетности
    period = Period_2(year, month)
    # устанавливаем 'period' как глобальную переменную (включая модули)
    builtins.period = period

    # Создаем новый файл отчетности xbrl, создав копию шаблона
    if report_type == 'month':
        # шаблон для квартальной отчетности
        shablon = file_shablon_oborot_month
    else:
        # шаблон для месячной отчетности
        shablon = file_shablon_oborot_quarter

    shutil.copyfile(dir_shablon + shablon, full_fileNewName)
    print(f'создан файл: {full_fileNewName}')

    # Загружаем данные из этого файла xbrl
    wb_xbrl = openpyxl.load_workbook(filename=full_fileNewName)

    oborotka(wb_xbrl, full_fileNewName, report_type)

    # Сохраняем в файл отчетности xbrl
    wb_xbrl.save(full_fileNewName)

    print('Вроде всё ОК!......')

    # ----------------------------------------------
