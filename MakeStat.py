from collections import defaultdict
import xlrd
import pandas as pd

filename = "Отчёт по кампании Кампания от IT директора Москва-Питер 3 07.04.2022 11_09.xls"
statistics = xlrd.open_workbook("Статистика.xls")
report = xlrd.open_workbook(filename)
stat = statistics.sheet_by_index(0)
rs = report.sheet_by_index(0)
print(rs.cell_value(rowx=1, colx=0))
print(stat.cell_value(rowx=1, colx=5))
print(stat.cell_value(rowx=3, colx=6))

times = list()
intervals = {'00:00 - 00:29': 0, '00:30 - 00:59': 0, '01:00 - 01:29': 0, '01:30 - 01:59': 0, '02:00 - 02:29': 0,
             '02:30 - 02:59': 0, '03:00 - 03:29': 0, '03:30 - 03:59': 0, '04:00 - 04:29': 0, '04:30 - 04:59': 0,
             '05:00 - 05:29': 0, '05:30 - 05:59': 0, '06:00 - 06:29': 0, '06:30 - 06:59': 0, '07:00 - 07:29': 0,
             '07:30 - 07:59': 0, '08:00 - 08:29': 0, '08:30 - 08:59': 0, '09:00 - 09:29': 0, '09:30 - 09:59': 0,
             '10:00 - 10:29': 0, '10:30 - 10:59': 0, '11:00 - 11:29': 0, '11:30 - 11:59': 0, '12:00 - 12:29': 0,
             '12:30 - 12:59': 0, '13:00 - 13:29': 0, '13:30 - 13:59': 0, '14:00 - 14:29': 0, '14:30 - 14:59': 0,
             '15:00 - 15:29': 0, '15:30 - 15:59': 0, '16:00 - 16:29': 0, '16:30 - 16:59': 0, '17:00 - 17:29': 0,
             '17:30 - 17:59': 0, '18:00 - 18:29': 0, '18:30 - 18:59': 0, '19:00 - 19:29': 0, '19:30 - 19:59': 0,
             '20:00 - 20:29': 0, '20:30 - 20:59': 0, '21:00 - 21:29': 0, '21:30 - 21:59': 0, '22:00 - 22:29': 0,
             '22:30 - 22:59': 0, '23:00 - 23:29': 0, '23:30 - 23:59': 0}


def find_time(tel):
    global stat
    for row in range(stat.nrows):
        if stat.cell_value(rowx=row, colx=2) == 'Отвечен':
            if stat.cell_value(rowx=row, colx=4) == 'Робот':
                if stat.cell_value(rowx=row, colx=6)[1:] == tel:
                    time = stat.cell_value(rowx=row, colx=3)
                    time = time.split()[1].split(':')[:2]
                    return time
            else:
                if stat.cell_value(rowx=row, colx=5)[1:] == tel:
                    time = stat.cell_value(rowx=row, colx=3)
                    time = time.split()[1].split(':')[:2]
                    return time
    raise ValueError


talk_start = 0
tel_number = 0
for i in range(rs.ncols):
    if rs.cell_value(rowx=0, colx=i) == 'Начат разговор':
        talk_start = i
    if rs.cell_value(rowx=0, colx=i) == 'Телефон':
        tel_number = i

for row in range(rs.nrows):
    if rs.cell_value(rowx=row, colx=talk_start) == 'Да':
        time = find_time(rs.cell_value(rowx=row, colx=tel_number)[1:])
        if int(time[1]) < 30:
            interval = time[0] + ':00' + ' - ' + time[0] + ':29'
            intervals[interval] += 1
        else:
            interval = time[0] + ':30' + ' - ' + time[0] + ':59'
            intervals[interval] += 1

print(end='\t')
for key, val in sorted(intervals.items()):
    print(key, end='\t')

print()
print(filename.split("\\")[-1][:-4], end='\t')
for key, val in sorted(intervals.items()):
    print(val, end='\t')
