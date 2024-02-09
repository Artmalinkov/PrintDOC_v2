# постановка текущей даты вместо флага
from src.main import *
import datetime



def iteration_row(MARK, ws, table):
    '''
    Функция просмотра строк на рабочем листе в таблице.
    :param MARK: маркер по котороку определяется необходимость печати
    :param ws: рабочий лист
    :param table: рабочая таблица
    :return:
    '''

    # Запускаем процедуру просмотра строк в таблице.
    for row in ws[table.ref]:

        # Для каждой строки формируем словарь
        dict_row = get_dict_row(table, row)

        # Если в нужном поле стоит единичка...
        if dict_row[PRINT_COLUMN_NAME] == MARK:

            # Запускаем процедуру замены
            doc_tpl.render(dict_row)

            # Если нужно - сохраняем, потом если нужно - распечатываем.
            if NEED_SAVE == True:
                save_doc_with_name(doc_tpl, dict_row)

                # Если нужно - распечатываем
                if NEED_PRINT == True:
                    print_doc(dict_row)
