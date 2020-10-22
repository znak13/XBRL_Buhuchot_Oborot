# Модуль расчета периодов

from datetime import timedelta, datetime

# ----------------------------------------------
# def choose():
#     """ Ввод периода отчетности """
#     reports = {'0': 'годовая',
#                '1': '1-ый квартал',
#                '2': '2-ой квартал',
#                '3': '3-ий квартал'}
#                # '4': '4-ый квартал'}
#     period = -1
#     attempt = 1 # номер попытки
#     while period not in reports:
#         period = input(f"Период отчетности:\n"
#                        f"0 ==> {reports['0']}\n"
#                        f"1 ==> {reports['1']}\n"
#                        f"2 ==> {reports['2']}\n"
#                        f"3 ==> {reports['3']}\n")
#                        # f"4 ==> {reports['4']}\n")
#
#         # Проверка корректности ввода периода
#         if period not in reports:
#             print(f'Выбор "{period}" - не верный. Попоробуйте ещё.\n')
#             attempt += 1
#             print(f'(попытка № {attempt})')
#
#         else:
#             print(f'Выбрана отчетность: {reports[period]}')
#
#     return int(period)
# ----------------------------------------------
# def dates(year = 2020):
#     # Существующие периоды
#     return  [datetime(year, 1, 1), datetime(year, 12, 31)], \
#             [datetime(year, 1, 1), datetime(year, 3, 31)], \
#             [datetime(year, 4, 1), datetime(year, 6, 30)], \
#             [datetime(year, 7, 1), datetime(year, 9, 30)], \
#             [datetime(year, 10, 1), datetime(year, 12, 31)]

def dates_2(year = 2020):
    # Существующие периоды
    return  [datetime(year, 1, 1), datetime(year, 12, 31)], \
            [datetime(year, 1, 1), datetime(year, 1, 31)], \
            [datetime(year, 2, 1), datetime(year, 3, 1) - timedelta(1)], \
            [datetime(year, 1, 1), datetime(year, 3, 31)], \
            [datetime(year, 4, 1), datetime(year, 4, 30)], \
            [datetime(year, 5, 1), datetime(year, 5, 31)], \
            [datetime(year, 4, 1), datetime(year, 6, 30)], \
            [datetime(year, 7, 1), datetime(year, 7, 31)], \
            [datetime(year, 8, 1), datetime(year, 8, 31)], \
            [datetime(year, 7, 1), datetime(year, 9, 30)], \
            [datetime(year, 10, 1), datetime(year, 10, 31)], \
            [datetime(year, 11, 1), datetime(year, 11, 30)], \
            [datetime(year, 10, 1), datetime(year, 12, 31)]

# ----------------------------------------------
def period_dates(current, delta):
    """ Даты в периоде"""

    period_year = current[0].year - delta
    period_begin = datetime(period_year, current[0].month, current[0].day).strftime("%Y-%m-%d")

    # конец периода расчитываем путем вычитания одного дня от первого дня следующего месяца
    # (актуально для февраля высокосного года)
    next_month = current[1].month + 1 if current[1].month != 12 else 1
    period_end = (datetime(period_year, next_month, 1) - timedelta(1)).strftime("%Y-%m-%d")

    period = period_begin + ' - ' + period_end

    # период с начала года
    period_year_begin = datetime(period_year, 1, 1).strftime("%Y-%m-%d")
    period_year_end = datetime(period_year, 12, 31).strftime("%Y-%m-%d")
    period_from_year = period_year_begin + ' - ' + period_end

    return period_begin, period_end, period, \
           period_year_begin, period_year_end, period_from_year
# ----------------------------------------------

# class Period():
#     def __init__(self, year, quarter):
#         self.year = int(year)
#         self.quarter = quarter
#         self.set()
#
#     def set(self):
#
#         # Выбор периода
#         # self.number = choose()
#         self.number = self.quarter
#
#         current = dates(year=self.year)[self.number]
#
#         # Текущий период
#         self.current_begin, self.current_end, self.current, \
#         self.current_year_begin, self.current_year_end, self.current_from_year \
#             = period_dates(current, 0)
#
#         # начало и конец последнего месяца в периоде
#         # (в случае, если в периоде несколько месяцев)
#         report_month = self.current_end.split('-')
#         report_month[2] = '01'
#         report_month = '-'.join(report_month)
#         self.report_month = report_month + ' - ' + self.current_end
#
#         # Предыдущий период
#         self.last_begin, self.last_end, self.last, \
#         self.last_year_begin, self.last_year_end, self.last_from_year \
#             = period_dates(current, 1)
#         # Позапрошлый период
#         self.before_last_begin, self.before_last_end, self.before_last, \
#         self.before_last_year_begin, self.before_last_year_end, self.before_last_from_year \
#             = period_dates(current, 2)
#
#         # смешанные периоды
#         self.current_mixed = self.current_from_year + ", " + \
#                              self.last_year_end + ", " + \
#                              self.current_end
#         self.last_mixed = self.last_from_year + ", " + \
#                           self.before_last_year_end + ", " + \
#                           self.last_end


class Period_2():
    def __init__(self, year, month):
        self.year = int(year)
        self.month = month
        self.set()

    def set(self):

        current = dates_2(year=self.year)[self.month]

        # Текущий период
        self.current_begin, self.current_end, self.current, \
        self.current_year_begin, self.current_year_end, self.current_from_year \
            = period_dates(current, 0)

        # начало и конец последнего месяца в периоде
        # (в случае, если в периоде несколько месяцев)
        report_month = self.current_end.split('-')
        report_month[2] = '01'
        report_month = '-'.join(report_month)
        self.report_month = report_month + ' - ' + self.current_end

        # Предыдущий период
        self.last_begin, self.last_end, self.last, \
        self.last_year_begin, self.last_year_end, self.last_from_year \
            = period_dates(current, 1)
        # Позапрошлый период
        self.before_last_begin, self.before_last_end, self.before_last, \
        self.before_last_year_begin, self.before_last_year_end, self.before_last_from_year \
            = period_dates(current, 2)

        # смешанные периоды
        self.current_mixed = self.current_from_year + ", " + \
                             self.last_year_end + ", " + \
                             self.current_end
        self.last_mixed = self.last_from_year + ", " + \
                          self.before_last_year_end + ", " + \
                          self.last_end



if __name__ == '__main__':

    # now = datetime.now()
    # print(now.strftime("%Y-%m-%d"))             # 2017-05-03
    # z = datetime(2021, 3, 1) - timedelta(1)

    pass
    p = Period_2(2020, 2)