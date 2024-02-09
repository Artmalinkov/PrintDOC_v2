'''
Основной функционал приложения по автозамене данных в шаблонах Word
'''
import os
import datetime
import win32api
from docxtpl import DocxTemplate
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table

# Путь к рабочему файлу Excel
WORKBOOK_PATH = r'C:\PythonProjects\Print_AS\attachments\excel_tpl.xlsx'

# Наименование рабочего листа
NAME_WORKSHEET = 'Лист1'

# Наименование таблицы на рабочем листе
TABLE_NAME = 'Таблица1'

# Столбец по которому определяется необходимость печати документа
PRINT_COLUMN_NAME = 'Печать'

# Маркер по отслеживанию необходимости печати
MARK = 1

# Путь к рабочему шаблону Word
DOCTPL_PATH = r"C:\PythonProjects\Print_AS\attachments\word_tpl.docx"

# Путь к папке, в которой будут лежать результаты
RESULT_DOC_DIR = r"C:\PythonProjects\Print_AS\done"

# Необходимость выполнения распечатывания документов
NEED_PRINT = False

# Необходимость сохранения изменённого документа Word в файл
NEED_SAVE = True

# Необходимость замены маркера на сегодняшнюю дату
NEED_CHANGE_NOW_DATE = True

# Необходимость сохранения рабочей книги после манипуляций
NEED_WB_SAVE = True


def get_start(WORKBOOK_PATH: str, NAME_WORKSHEET: str, TABLE_NAME: str, TEST_DOCTPL_PATH: str):
    '''
    Функция проводит первоначальное определение основных объектов с которыми с последующем будем работать
    :param WORKBOOK_PATH:
    :param NAME_WORKSHEET:
    :param TABLE_NAME:
    :param TEST_DOCTPL_PATH:
    :return: wb, ws, table, doc_tpl
    '''

    # Загружаем рабочую книгу
    wb = load_workbook(WORKBOOK_PATH)

    # Определение рабочего листа
    ws = wb[NAME_WORKSHEET]

    # Определяем тяблицу
    table = ws.tables[TABLE_NAME]

    # Определяем шаблон Word
    doc_tpl = DocxTemplate(TEST_DOCTPL_PATH)

    return wb, ws, table, doc_tpl


def get_dict_row(table, row):
    '''
    Функция формирования словаря из значений в строке
    :param table: объект таблицы на рабочем листе
    :param row: объект строки в таблице
    :return: dict_row - сформированный словарь: ключ - заголовок столбца, значение - ячейка
    '''
    dict_row = {}
    for cell in row:
        # Преобразование даты
        if type(cell.value) == datetime.datetime:
            cell.value = cell.value.strftime('%d.%m.%Y')
        # Наполнение словаря
        dict_row[table.column_names[cell.column - 1]] = cell.value
    return dict_row


def iteration_row(wb, ws, table, doc_tpl, MARK):
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

            # Если нужно - заменяем MARK на текущую дату
            if NEED_CHANGE_NOW_DATE == True:
                change_now_date(wb, table, row, PRINT_COLUMN_NAME, NEED_WB_SAVE)


def change_now_date(wb, table, row, CHANGE_COLUMN_NAME, NEED_WB_SAVE):
    '''
    Функция заменяет значение ячейки в строке на текущую дату.
    :param wb: рабочая книга
    :param table: рабочая таблица
    :param row: исследуемая строка
    :param CHANGE_COLUMN_NAME: наименование столбца в котором нужно заменить значение
    :param NEED_WB_SAVE: определяет, нужно ли сохранять рабочую книгу после манипуляций
    :return:
    '''
    index_change_column = table.column_names.index(CHANGE_COLUMN_NAME)
    row[index_change_column].value = datetime.datetime.now().strftime('%d.%m.%Y')
    if NEED_WB_SAVE == True:
        wb.save(WORKBOOK_PATH)


def save_doc_with_name(changed_doc, dict_row):
    '''
    Функция сохранения измененённого документа
    :param changed_doc: изменённый документ
    :param context: словарь нужной строки
    :return:
    '''
    doc_name = RESULT_DOC_DIR + '\\' + f"{dict_row['Фамилия']}{dict_row['Имя']}.docx"
    changed_doc.save(doc_name)


def print_doc(dict_row):
    '''
    Функция распечатки документа
    :param dict_row: словарь нужной строки
    :return:
    '''

    doc_name = f"done/{dict_row['Фамилия']}{dict_row['Имя']}.docx"

    full_filepath = os.path.abspath(doc_name)
    win32api.ShellExecute(0, 'print', full_filepath, None, '.', 0)


def print_doc(dict_row):
    '''
    Функция распечатки документа
    :param dict_row: словарь нужной строки
    :return:
    '''
    full_filepath = RESULT_DOC_DIR + '\\' + f"{dict_row['Фамилия']}{dict_row['Имя']}.docx"
    win32api.ShellExecute(0, 'print', full_filepath, None, '.', 0)


def main():
    wb, ws, table, doc_tpl = get_start(WORKBOOK_PATH, NAME_WORKSHEET, TABLE_NAME, DOCTPL_PATH)
    iteration_row(wb, ws, table, doc_tpl, MARK)


if __name__ == '__main__':
    main()
