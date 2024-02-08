# постановка текущей даты вместо флага
from src.main import *
import datetime
import src.main



row = ws[table.ref][0]
row[4].value
for row in ws[table.ref]:
    if row[4].value == 1:
        # Запускается процедура total_print
        row[4].value = datetime.datetime.now().strftime('%d.%m.%Y')

wb.save(r'C:\PythonProjects\Print_AS\attachments\excel_tpl_TEST.xlsx')



# TODO нужно обернуть это в функцию
# TODO функция должна заменять значение в конкретной ячейке, тоесть нужны индексы и столбца и строки
FLAG_COLUMN_NAME = 'Печать'


wb = load_workbook(r'C:\PythonProjects\Print_AS\attachments\excel_tpl.xlsx')
ws = wb.active
table = ws.tables['Таблица1']

# запускаем процедуру просмотра строк, при соблюдении условий - начинаем действия


row = ws[table.ref][1]


index_neces_column = get_index_neces_column(table, FLAG_COLUMN_NAME)
row[index_neces_column-1].value




for row in ws[table.ref]:
    if row[column_id].value == 1:
        # Запускается процедура total_print
        row[column_id].value = datetime.datetime.now().strftime('%d.%m.%Y')


def replace_mark_date(ws, table, column_id):
    '''

    :param ws:
    :param table:
    :param column_id:
    :return:
    '''




