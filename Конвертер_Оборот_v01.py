from module.Oborot import oborotka
from module.periods import Period
from module.analiz_data import *
from module.functions import *
import builtins


# %%


if __name__ == "__main__":

    # Ввод периода отчетности
    period = Period()
    period.set()
    builtins.period = period

    # Создаем новый файл отчетности xbrl на основе файла-шаблона
    FileNewName = load_xbrl(file_shablon_oborot, file_dir=dir_shablon)
    # Загружаем данные из этого файла xbrl
    wb_xbrl = openpyxl.load_workbook(filename=FileNewName)

    oborotka(wb_xbrl, FileNewName)

    # Сохраняем в файл отчетности xbrl
    wb_xbrl.save(FileNewName)

    print('Вроде всё ОК!......')

    # ----------------------------------------------
