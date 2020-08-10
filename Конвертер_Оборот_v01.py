import openpyxl
from module.Oborot import oborotka
from module.periods import Period
# from module.analiz_data import *
from module.functions import *
import builtins
import module.period_selection as selection

# %%


if __name__ == "__main__":

    # Выбор периода отчетности
    year, quarter, dir_QuarterReports, fileNewName = selection.main()
    dir_QuarterReports += '/'
    full_fileNewName = dir_QuarterReports + fileNewName

    # Ввод периода отчетности
    period = Period(year, quarter)
    # устанавливаем 'period' как глобальную переменную (включая модули)
    builtins.period = period

    # Создаем новый файл отчетности xbrl, создав копию шаблона
    shutil.copyfile(dir_shablon + file_shablon_oborot, full_fileNewName)
    print(f'создан файл: {full_fileNewName}')

    # Загружаем данные из этого файла xbrl
    wb_xbrl = openpyxl.load_workbook(filename=full_fileNewName)

    oborotka(wb_xbrl, full_fileNewName)

    # Сохраняем в файл отчетности xbrl
    wb_xbrl.save(full_fileNewName)

    print('Вроде всё ОК!......')

    # ----------------------------------------------
