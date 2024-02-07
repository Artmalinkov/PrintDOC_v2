# постановка даты текущей даты вместо флага
from src.main import *
import datetime


for row in ws[table.ref]:
    for cell in row:


        if cell.value == 1:
            cell.value = datetime.datetime.now().strftime('%d.%m.%Y')
        print(cell.value)

worksheet = ws
# Перераблотать функцию total_print() для большей универсальности
def total_print_doc(worksheet, table, doc_tpl, flag_column_name):
    '''
    Сохранение и распечатка каждого документа, отмеченного флагом 'Печать'
    :return:
    '''

    # Определяем, в каком по счёту столбце находится заголовок с именем flag_column_name
    flag_column = get_column_id(table, flag_column_name)

    # Определяем где в столбце flag_column_name стоит '1' - с этой строкой нужно работать.
    for row in worksheet[table.ref]:
        index_neces_row = get_index_neces_row(flag_column, row)
        if index_neces_row != None:
            # Запускаем процедуру замены
            changed_doc, context = one_render(worksheet, table, index_neces_row, doc_tpl)

            # Если необходимо сохранить - сохраняем полученные результаты в файл
            if NEED_SAVE == True:
                save_doc_with_name(changed_doc, context)

            # Если необходимо распечатать - распечатываем полученный файл
            if NEED_PRINT == True:
                print_doc(context)

total_print_doc(ws,table,doc_tpl,FLAG_COLUMN_NAME)




